import os
import sys
import json
import re
import datetime
import time
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font
from openpyxl.cell import MergedCell
from bs4 import BeautifulSoup

# Force current working directory to be the script's directory
script_dir = os.path.dirname(os.path.abspath(__file__))
os.chdir(script_dir)

# Ensure input folder exists in the correct directory
os.makedirs("data_input", exist_ok=True)

CONFIG_FILE = "config.json"
RICH_SUMMARIES_FILE = "data_input/ai_summaries_rich.json"

def safe_save_workbook(wb, filepath):
    """
    Attempts to save the workbook. Bypasses openpyxl serialization bugs by converting 
    all ArrayFormula objects to normal raw formula strings (str) before saving.
    This prevents Excel from throwing sheet formula corruption warnings.
    """
    from openpyxl.worksheet.formula import ArrayFormula
    print("Optimizing formulas to prevent Excel corruption...")
    for sheetname in wb.sheetnames:
        ws = wb[sheetname]
        converted = 0
        # Use iter_rows for extremely fast cell traversal without coordinate parsing overhead
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
            for cell in row:
                val = cell.value
                if isinstance(val, ArrayFormula):
                    cell.value = val.text
                    converted += 1
        if converted > 0:
            print(f"  Converted {converted} ArrayFormulas in sheet '{sheetname}'")
            
    while True:
        try:
            wb.save(filepath)
            break
        except PermissionError:
            print(f"\n[エラー/警告] ファイルに書き込めません: '{filepath}'")
            print("Excelアプリ等でこのファイルが開いている可能性があります。")
            print("Excelを閉じてから、キーボードの【ENTER】キーを押して再試行してください。")
            input("閉じた後にENTERキーを押してください...")

def safe_write_cell(ws, row, col, value):
    """
    Safely writes a value to a cell, resolving the top-left cell if it is a MergedCell.
    """
    cell = ws.cell(row, col)
    if isinstance(cell, MergedCell):
        for merged_range in ws.merged_cells.ranges:
            if cell.coordinate in merged_range:
                top_left = ws.cell(merged_range.min_row, merged_range.min_col)
                top_left.value = value
                return top_left
    else:
        cell.value = value
        return cell

def is_empty_or_formula(val):
    """
    Determines if a cell value is empty, None, whitespace, or an Excel formula.
    """
    if val is None:
        return True
    val_str = str(val).strip()
    if not val_str:
        return True
    if val_str.startswith("="):
        return True
    return False

def normalize_date_string(date_val):
    """
    Normalizes any date object or string into standard 'YYYY/MM/DD' format.
    Handles 'YYYY-MM-DD HH:MM:SS', 'YYYY-MM-DD', etc.
    """
    if date_val is None:
        return ""
    if isinstance(date_val, datetime.datetime):
        return date_val.strftime("%Y/%m/%d")
    if isinstance(date_val, datetime.date):
        return date_val.strftime("%Y/%m/%d")
        
    date_str = str(date_val).strip()
    match = re.match(r'^(\d{4})[-/](\d{1,2})[-/](\d{1,2})', date_str)
    if match:
        return f"{match.group(1)}/{int(match.group(2)):02d}/{int(match.group(3)):02d}"
    return date_str

def load_config():
    if not os.path.exists(CONFIG_FILE):
        print(f"Error: {CONFIG_FILE} not found. Please create it first.")
        sys.exit(1)
    with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
        return json.load(f)

def read_file_with_encoding(filepath):
    encodings = ['utf-8', 'cp932', 'shift_jis', 'utf-8-sig']
    for enc in encodings:
        try:
            with open(filepath, 'r', encoding=enc) as f:
                content = f.read()
                if "機種" in content or "台番" in content or "差枚" in content:
                    print(f"File successfully read with encoding: {enc}")
                    return content
        except Exception:
            continue
            
    print("Warning: Could not detect standard encoding. Reading with fallback (errors ignored).")
    with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
        return f.read()

def parse_html_data(filepath):
    """
    Parses Pachislot data from a downloaded HTML file.
    Only targets tables with explicit '台番' or '台番号' headers to ignore summary/history tables.
    """
    results = []
    seen_machine_numbers = set()
    print(f"Parsing HTML data from: {filepath}")
    html_content = read_file_with_encoding(filepath)
    soup = BeautifulSoup(html_content, 'html.parser')
        
    tables = soup.find_all('table')
    for table in tables:
        headers = [th.get_text().strip() for th in table.find_all('th')]
        if not headers:
            first_row = table.find('tr')
            if first_row:
                headers = [td.get_text().strip() for td in first_row.find_all('td')]
                
        header_text = "".join(headers)
        
        # Strict check: Must contain BB and REG/RB, AND must explicitly mention 台番 or 台番号
        is_slot_table = (
            ("BB" in headers and "RB" in headers) and 
            any(h in header_text for h in ["台番", "台番号"])
        )
        if not is_slot_table:
            continue
            
        rows = table.find_all('tr')
        for row in rows:
            cells = [td.get_text().strip() for td in row.find_all('td')]
            if len(cells) >= 5:
                cell_text = "".join(cells)
                if any(x in cell_text for x in ["平均", "累計", "合計"]):
                    continue
                try:
                    def clean_val(val):
                        return val.replace(",", "").replace("+", "").replace(" ", "").strip()
                        
                    c0_clean = clean_val(cells[0])
                    c1_clean = clean_val(cells[1])
                    
                    if c0_clean.isdigit():
                        machine_number = int(c0_clean)
                        machine_name = cells[1]
                        g_games = int(clean_val(cells[2]))
                        diff_coins = int(clean_val(cells[3]))
                        bb_count = int(clean_val(cells[4]))
                        rb_count = int(clean_val(cells[5])) if len(cells) > 5 else 0
                    else:
                        machine_name = cells[0]
                        machine_number = int(c1_clean)
                        g_games = int(clean_val(cells[2]))
                        diff_coins = int(clean_val(cells[3]))
                        bb_count = int(clean_val(cells[4]))
                        rb_count = int(clean_val(cells[5])) if len(cells) > 5 else 0
                    
                    if machine_number in seen_machine_numbers:
                        continue
                    seen_machine_numbers.add(machine_number)
                    
                    results.append({
                        "machine_name": machine_name,
                        "machine_number": machine_number,
                        "g_games": g_games,
                        "diff_coins": diff_coins,
                        "bb_count": bb_count,
                        "rb_count": rb_count
                    })
                except ValueError:
                    continue
                    
    print(f"Successfully extracted {len(results)} unique rows from HTML tables.")
    return results

def parse_text_data(filepath):
    results = []
    print(f"Parsing Text data from: {filepath}")
    with open(filepath, 'r', encoding='utf-8') as f:
        lines = f.readlines()
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
        
        tokens = re.split(r'\t+|\s{2,}', line)
        if len(tokens) < 4:
            tokens = line.split(' ')
            if len(tokens) < 4:
                continue

        if any(x in line for x in ["機種", "台番", "平均", "累計", "合計"]):
            continue

        try:
            numbers = []
            machine_name = ""
            for token in tokens:
                token = token.replace(",", "").replace("+", "").strip()
                if re.match(r'^-?\d+$', token):
                    numbers.append(int(token))
                else:
                    if token and token not in ["-", "★"]:
                        machine_name = token
            
            if len(numbers) >= 4:
                machine_number = numbers[0]
                g_games = numbers[1]
                diff_coins = numbers[2]
                bb_count = numbers[3]
                rb_count = len(numbers) > 4 and numbers[4] or 0
                
                results.append({
                    "machine_name": machine_name,
                    "machine_number": machine_number,
                    "g_games": g_games,
                    "diff_coins": diff_coins,
                    "bb_count": bb_count,
                    "rb_count": rb_count
                })
        except Exception as e:
            continue
            
    print(f"Successfully extracted {len(results)} rows from Text.")
    return results

def copy_excel_formula(ws, src_row, dst_row, max_col):
    """
    Copies formulas from src_row to dst_row for columns > 7, dynamically updating the row numbers.
    """
    from openpyxl.worksheet.formula import ArrayFormula
    for col in range(1, max_col + 1):
        cell = ws.cell(src_row, col)
        if col > 7:
            val = cell.value
            if isinstance(val, ArrayFormula):
                val = val.text
            if isinstance(val, str) and val.startswith("="):
                new_formula = re.sub(rf'\b([A-Z]+){src_row}\b', rf'\g<1>{dst_row}', val)
                ws.cell(dst_row, col, new_formula)
            elif val is not None:
                ws.cell(dst_row, col, val)

def update_excel_data(excel_path, target_date, parsed_data):
    import excel_calc_logic
    if not os.path.exists(excel_path):
        print(f"Error: Excel file {excel_path} not found.")
        sys.exit(1)
        
    print(f"Opening Excel file: {excel_path}...")
    wb = openpyxl.load_workbook(excel_path, data_only=False)
    ws = wb['【データ】蓄積用']
    
    # Find actual max row ignoring formulas and empty rows
    real_max_row = 1
    for r in range(ws.max_row, 0, -1):
        val1 = ws.cell(r, 1).value
        val2 = ws.cell(r, 2).value
        if not is_empty_or_formula(val1) or not is_empty_or_formula(val2):
            real_max_row = r
            break
            
    max_col = ws.max_column
    
    # --- FORMULA AUTO-RESTORATION LOGIC ---
    # Python-based K-AF calculations completely bypass the need for Excel array formulas.
    # We no longer attempt to restore missing formulas in columns 11-32.
    # --------------------------------------
    
    # Check if target_date already exists in Column A (compare as dates)
    date_val = datetime.datetime.strptime(target_date, "%Y/%m/%d")
    date_exists = False
    for r in range(2, real_max_row + 1):
        cell_val = ws.cell(r, 1).value
        if isinstance(cell_val, datetime.datetime) and cell_val.date() == date_val.date():
            date_exists = True
            break
            
    if date_exists:
        print(f"Notice: Data for date {target_date} already exists in 【データ】蓄積用.")
        print("Skipping append to prevent duplication. Moving to AI analysis...")
        safe_save_workbook(wb, excel_path)  # Save the workbook safely
        return
        
    print(f"Current actual max row in 蓄積用: {real_max_row}. Appending {len(parsed_data)} lines...")
    
    # Pre-build specs, settings and history for Python K-AF calculation
    specs_dict, settings_dict = excel_calc_logic.build_maps(wb)
    processed_history = excel_calc_logic.extract_history(ws, real_max_row)
    
    start_row = real_max_row + 1
    for i, data in enumerate(parsed_data):
        current_row = start_row + i
        
        # Write values and force Date format on Column A
        dt_cell = ws.cell(current_row, 1, date_val)
        dt_cell.number_format = 'yyyy/mm/dd'
        
        ws.cell(current_row, 2, data["machine_name"])
        ws.cell(current_row, 3, data["machine_number"])
        ws.cell(current_row, 4, data["g_games"])
        ws.cell(current_row, 5, data["diff_coins"])
        ws.cell(current_row, 6, data["bb_count"])
        ws.cell(current_row, 7, data["rb_count"])
        
        # Calculate and write Big, Reg, and Combined probabilities (Cols 8, 9, 10)
        games_val = int(data["g_games"] or 0)
        bb_val = int(data["bb_count"] or 0)
        rb_val = int(data["rb_count"] or 0)
        ws.cell(current_row, 8, excel_calc_logic.format_probability(games_val, bb_val + rb_val))
        ws.cell(current_row, 9, excel_calc_logic.format_probability(games_val, bb_val))
        ws.cell(current_row, 10, excel_calc_logic.format_probability(games_val, rb_val))
        
        # Calculate K-AF columns via Python!
        row_dict = {
            "row_num": current_row,
            "date": date_val,
            "machine_name": str(data["machine_name"]) if data.get("machine_name") else "",
            "machine_number": int(data["machine_number"]),
            "g_games": int(data["g_games"] or 0),
            "diff_coins": int(data["diff_coins"] or 0),
            "bb_count": int(data["bb_count"] or 0),
            "rb_count": int(data["rb_count"] or 0)
        }
        excel_calc_logic.calculate_kaf_for_row(ws, current_row, row_dict, settings_dict, specs_dict, processed_history)

        
    print(f"Writing complete. Saving Excel file...")
    safe_save_workbook(wb, excel_path)
    print("Excel file successfully updated with new raw data and formulas.")

def prepare_ai_context(excel_path, target_date):
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    
    # 1. Dashboard stats
    dash_data = []
    if '【PC】分析ダッシュボード' in wb.sheetnames:
        dash_ws = wb['【PC】分析ダッシュボード']
        dash_data.append("### 【PC】分析ダッシュボード 現在の統計値 ###")
        for r in range(4, 38):
            row_vals = [dash_ws.cell(r, c).value for c in range(2, 6)]
            if any(row_vals):
                dash_data.append(f"Row {r}: " + " | ".join([str(x) if x is not None else "" for x in row_vals]))
    else:
        dash_data.append("### 【PC】分析ダッシュボード (削除済み - 代替Web版参照) ###")
            
    # 2. Raw records
    data_ws = wb['【データ】蓄積用']
    data_rows = []
    data_rows.append("### 【データ】蓄積用 最新直近データ (直近100台) ###")
    data_rows.append("日付 | 機種名 | 台番号 | G数 | 差枚 | BB | RB | 推定設定 | 機械割 | 期待収支")
    
    real_max_row = 1
    for r in range(data_ws.max_row, 0, -1):
        val1 = data_ws.cell(r, 1).value
        val2 = data_ws.cell(r, 2).value
        if not is_empty_or_formula(val1) or not is_empty_or_formula(val2):
            real_max_row = r
            break
            
    start_row = max(2, real_max_row - 100)
    for r in range(start_row, real_max_row + 1):
        row_vals = [
            data_ws.cell(r, 1).value, # A: 日付
            data_ws.cell(r, 2).value, # B: 機種
            data_ws.cell(r, 3).value, # C: 台番
            data_ws.cell(r, 4).value, # D: G数
            data_ws.cell(r, 5).value, # E: 差枚
            data_ws.cell(r, 6).value, # F: BB
            data_ws.cell(r, 7).value, # G: RB
            data_ws.cell(r, 33).value, # AG: 推定設定
            data_ws.cell(r, 34).value, # AH: 機械割
            data_ws.cell(r, 35).value, # AI: 期待収支
        ]
        
        if isinstance(row_vals[0], datetime.datetime):
            row_vals[0] = row_vals[0].strftime("%Y/%m/%d")
            
        data_rows.append(" | ".join([str(x) if x is not None else "" for x in row_vals]))
        
    # 3. FEEDBACK LOOP: Load AI Prediction history and actual results (last 30 slots)
    pred_history = []
    if "【AI】予想・答え合わせ" in wb.sheetnames:
        predict_ws = wb["【AI】予想・答え合わせ"]
        pred_history.append("### 過去のAI予想・答え合わせ履歴 (直近の答え合わせ結果) ###")
        pred_history.append("予想日 | 機種名 | 台番号 | 予想・狙い根拠 | 実際のG数 | 実際の差枚 | 設定スコア | 判定(〇/×)")
        
        real_max_pred = 3
        for r in range(predict_ws.max_row, 3, -1):
            if predict_ws.cell(r, 1).value is not None:
                real_max_pred = r
                break
                
        start_pred_r = max(4, real_max_pred - 30)
        for r in range(start_pred_r, real_max_pred + 1):
            row_vals = [
                predict_ws.cell(r, 1).value, # A: 予想日
                predict_ws.cell(r, 2).value, # B: 機種名
                predict_ws.cell(r, 3).value, # C: 台番号
                predict_ws.cell(r, 4).value, # D: 根拠
                predict_ws.cell(r, 5).value, # E: 実際のG数
                predict_ws.cell(r, 6).value, # F: 実際の差枚
                predict_ws.cell(r, 7).value, # G: 設定スコア
                predict_ws.cell(r, 8).value, # H: 判定
            ]
            if isinstance(row_vals[0], datetime.datetime):
                row_vals[0] = row_vals[0].strftime("%Y/%m/%d")
            pred_history.append(" | ".join([str(x) if x is not None else "" for x in row_vals]))
            
    # 4. Load manual confirmation info (Trophy, events, SNS hints, etc.)
    confirm_data = []
    if "確認情報" in wb.sheetnames:
        confirm_ws = wb["確認情報"]
        confirm_data.append("### ユーザー入力の店舗確認情報 (公約・示唆・トロフィー等の一次情報) ###")
        
        for r in range(2, confirm_ws.max_row + 1):
            row_vals = [confirm_ws.cell(r, c).value for c in range(1, confirm_ws.max_column + 1)]
            if not any(v is not None for v in row_vals):
                continue
                
            row_date = None
            other_texts = []
            for val in row_vals:
                if val is None:
                    continue
                if isinstance(val, (datetime.datetime, datetime.date)):
                    row_date = val.strftime("%Y/%m/%d")
                elif isinstance(val, str):
                    val_clean = val.strip()
                    # Match dates like YYYY/MM/DD or YYYY-MM-DD
                    if re.match(r'^\d{4}/\d{2}/\d{2}$', val_clean) or re.match(r'^\d{4}-\d{2}-\d{2}$', val_clean):
                        row_date = val_clean.replace("-", "/")
                    else:
                        if val_clean:
                            other_texts.append(val_clean)
                else:
                    other_texts.append(str(val))
                    
            if row_date and other_texts:
                confirm_data.append(f"  {row_date} : " + " / ".join(other_texts))
            elif other_texts:
                confirm_data.append(f"  (日付不明) : " + " / ".join(other_texts))
    else:
        confirm_data.append("  (入力情報なし)")
        
    wb.close()
    
    return (
        "\n".join(dash_data) + "\n\n" + 
        "\n".join(data_rows) + "\n\n" + 
        "\n".join(confirm_data) + "\n\n" + 
        "\n".join(pred_history)
    )

def run_gemini_analysis(api_key, context, target_date):
    import google.generativeai as genai
    genai.configure(api_key=api_key)
    
    models_to_try = [
        'gemini-3.1-pro-preview',
        'gemini-3.5-flash',
        'gemini-3.1-flash-lite',
        'gemini-2.0-flash'
    ]
    
    d_obj = datetime.datetime.strptime(target_date, "%Y/%m/%d")
    tomorrow_date = (d_obj + datetime.timedelta(days=1)).strftime("%Y/%m/%d")
    
    prompt = f"""
あなたはパチスロホールの設定配分を分析するプロのデータサイエンティストであり、月100万円以上を安定して稼ぐ現役パチプロ兼データアナリストです。
提供されたホールの営業データ、および「過去のAI予想・答え合わせ履歴」「ユーザーが直接入力した店舗確認情報（公約や示唆、確定情報など）」から、店長の投入クセの傾向変化を学習した上で、明日の推奨狙い台を決定してください。

一般ユーザー向けの解説は不要です。期待値を最大化するための分析のみを行ってください。

---
【直近のデータ状況】
分析対象日: {target_date} （このデータを本日追加しました）
狙い対象日 (翌日): {tomorrow_date}

以下に最新のダッシュボード値、蓄積用データの直近100台分のサマリー、ユーザー入力の最新店舗確認情報（公約やトロフィーなどの一次情報）、および過去30台分の「答え合わせ履歴（判定〇×）」を提供します。
{context}

---
【最優先・分析ルール】

① 一次情報（店舗確認情報）の最優先
「店舗確認情報」にトロフィー、確定画面、SNS示唆、LINE示唆、イベント公約、演者来店などが入力されている場合、これはAI of 過去データ推測より優先される「絶対的事実」です。これらに合致する台（例: 全台系公約、並び公約など）がある場合は、最優先で狙い台として選定してください。

② 荒波機種（スマスロ等）の補正
からくりサーカス、ヴァルヴレイヴ、戦国乙女、チバリヨ、かぐや様、東京喰種、北斗、ゴッドイーター、モンキーターンなどは、差枚数だけで設定を判断せず、ゲーム数、初当たり、設定差のある要素（小役・示唆）を重視してください。3000G未満の大量出玉は「誤爆リスク」として評価し、据え置き期待度は下げてください。

③ ノーマルタイプ（ジャグラー等）の補正
ジャグラー、ハナハナ等のノーマル機は、差枚よりも「REG確率」「合算確率」「回転数」「ブドウ逆算値」「REG先行度合い」を最重視して判断してください。

④ リスク評価とあいまいさの排除
中間設定（設定4〜5）の可能性やヒキ強による誤爆リスクも必ず根拠内で考慮してください。「設定6濃厚」という安易な表現は禁止し、「設定5〜6期待」「高設定期待」などの堅実な表現を使用してください。

⑤ 店長心理・クセの分析
店長の投入パターン（並び、全台系、据え置き、末尾、角、中央、ローテーション傾向）を分析し、店長の心理を考察してください。

---
【絶対厳守の出力フォーマット】
以下の構成で日本語で出力してください。スプレッドシートやドキュメントにそのままコピーできる構成とします。

① 【AI】営業評価と店長の心理総括
冒頭に、必ず明日の参戦評価として以下のいずれか1つを大見出しで明記してください。
- **【明日の参戦評価：行ける（勝負すべき日）】**
- **【明日の参戦評価：狙い目だけ打ちに行く（ピンポイント狙い）】**
- **【明日の参戦評価：行く価値無し（見送り推奨）】**
その後に、最新日の結果を踏まえた店長の意図（還元・回収）の考察を1分で読める文量で書いてください。

② 機種別・設定投入の「本気度」検証
「今、最も設定が狙える本命機種」と「回収用の死に機種」のリストアップと理由。

③ Excel予測スコアの「妥当性検証レポート」
現在の予測スコアが実際の店舗傾向と合っているかの検証。ズレの指摘。

④ AI独自の次回（明日）の推奨狙い目台（全機種TOP5 ＆ ジャグラーTOP5）
表面的なスコアだけでなく、「周期」「機種の強さ」「確認情報の事実」を複合して補正した推奨台と根拠。

⑤ 【AI】予想・答え合わせ コピペ用テーブル
推奨狙い目台（全機種5台 ＋ ジャグラー5台 ＝ 計10台）を、以下のMarkdown表形式で出力してください。
余計なテキストを挟まず、必ず表をそのまま出力すること。
日付には結果日（＝狙い日：{tomorrow_date}）を記入すること。

★重要★: 推奨度の強さを可視化するため、「予想・狙い根拠」の文頭には必ず【推奨Sランク - 95点】、【推奨Aランク - 80点】、【推奨Bランク - 65点】のように、推奨ランク（S/A/B）と100点満点での評価スコアを記載してください。期待値の低い台はBランク以下で表現し、自信のある台はS〜Aランクにしてください。

| 日付 | 機種名 | 台番号 | 予想・狙い根拠 |
| --- | --- | --- | --- |
| {tomorrow_date} | [機種名] | [台番号] | 【推奨[S/A/B]ランク - [点数]点】[具体的な数値や確認情報を交えた狙い根拠] |
"""
    
    for model_name in models_to_try:
        try:
            print(f"Sending prompt to Gemini API using model: {model_name}... (This may take a moment)")
            model = genai.GenerativeModel(model_name)
            response = model.generate_content(prompt)
            return response.text
        except Exception as e:
            err_str = str(e).lower()
            is_skippable = (
                any(x in err_str for x in ["429", "quota", "exhausted", "limit"]) or
                any(x in err_str for x in ["404", "not found", "no longer available", "deprecated"])
            )
            if is_skippable:
                print(f"Notice: Model {model_name} is rate-limited, deprecated or unavailable. Automatically trying next model...")
                continue
            else:
                raise e
                
    raise Exception("Error: All available Gemini models failed due to quota limits. Please try again after a few minutes.")

def write_ai_results_to_excel(excel_path, target_date, ai_text):
    d_obj = datetime.datetime.strptime(target_date, "%Y/%m/%d")
    tomorrow_date = (d_obj + datetime.timedelta(days=1)).strftime("%Y/%m/%d")
    tomorrow_date_val = datetime.datetime.strptime(tomorrow_date, "%Y/%m/%d")
    
    wb = openpyxl.load_workbook(excel_path, data_only=False)
    
    # 1. Parse AI recommendations from Markdown table
    table_pattern = r'\|?\s*(?:\d{4}/\d{2}/\d{2})\s*\|.*'
    table_lines = re.findall(table_pattern, ai_text)
    
    rows_to_insert = []
    for line in table_lines:
        cells = [c.strip() for c in line.split("|")]
        if len(cells) > 0 and cells[0] == "":
            cells.pop(0)
        if len(cells) > 0 and cells[-1] == "":
            cells.pop()
            
        if len(cells) >= 4:
            date_str = cells[0]
            machine_name = cells[1]
            try:
                machine_num = int(cells[2])
            except ValueError:
                machine_num = cells[2]
            reason = cells[3]
            
            if "機種名" in machine_name or "---" in machine_name:
                continue
                
            rows_to_insert.append((date_str, machine_name, machine_num, reason))
            
    if rows_to_insert:
        ai_ws = wb['【AI】予想・答え合わせ']
        
        # --- PREVENT DUPLICATE RECOMMENDATIONS ---
        # Scan and remove any existing rows for tomorrow_date_val to prevent infinite duplicates
        rows_to_delete = []
        for r in range(4, ai_ws.max_row + 1):
            cell_val = ai_ws.cell(r, 1).value
            normalized_cell_val = normalize_date_string(cell_val)
            normalized_tomorrow = normalize_date_string(tomorrow_date_val)
            if normalized_cell_val == normalized_tomorrow:
                rows_to_delete.append(r)
                
        if rows_to_delete:
            print(f"Removing {len(rows_to_delete)} duplicate prediction rows in 【AI】予想・答え合わせ for date {tomorrow_date}...")
            # Delete in reverse order to keep indices correct
            for r in reversed(rows_to_delete):
                ai_ws.delete_rows(r)
        # ------------------------------------------
        
        # Recalculate last row after deletion
        last_data_row = 3
        for r in range(ai_ws.max_row, 3, -1):
            val = ai_ws.cell(r, 1).value
            if not is_empty_or_formula(val):
                last_data_row = r
                break
        append_start_row = last_data_row + 1
            
        print(f"Writing {len(rows_to_insert)} recommendation rows starting at Row {append_start_row}...")
        for i, row_data in enumerate(rows_to_insert):
            curr_r = append_start_row + i
            
            # Write date and format
            dt_cell = ai_ws.cell(curr_r, 1, tomorrow_date_val)
            dt_cell.number_format = 'yyyy/mm/dd'
            
            ai_ws.cell(curr_r, 2, row_data[1])
            ai_ws.cell(curr_r, 3, row_data[2])
            ai_ws.cell(curr_r, 4, row_data[3])
    else:
        print("Warning: Could not parse recommendation table from AI response.")
        
    # 2. PERFORM PRE-CALCULATED LOOKUPS (Excel speed-up)
    print("Pre-calculating answer key data (lookup actual results)...")
    data_ws = wb['【データ】蓄積用']
    accumulated_db = {}
    for r in range(2, data_ws.max_row + 1):
        date_val = data_ws.cell(r, 1).value
        mach_num = data_ws.cell(r, 3).value
        g_games = data_ws.cell(r, 4).value
        diff_coins = data_ws.cell(r, 5).value
        setting_score = data_ws.cell(r, 29).value  # Col 29 (AC): 最終設定スコア
        
        if date_val is not None:
            date_str = normalize_date_string(date_val)
            try:
                m_num = int(str(mach_num).strip())
                accumulated_db[(date_str, m_num)] = (g_games, diff_coins, setting_score)
            except ValueError:
                continue
                
    ai_ws = wb['【AI】予想・答え合わせ']
    rewritten_count = 0
    # Scan all prediction rows (4 to max_row) and replace formulas with static values
    for r in range(4, ai_ws.max_row + 1):
        date_val = ai_ws.cell(r, 1).value
        mach_val = ai_ws.cell(r, 3).value
        m_name = str(ai_ws.cell(r, 2).value or "")
        
        if date_val is not None and mach_val is not None:
            date_str = normalize_date_string(date_val)
            try:
                m_num = int(str(mach_val).strip())
                key = (date_str, m_num)
                if key in accumulated_db:
                    g_games, diff_coins, setting_score = accumulated_db[key]
                    
                    try:
                        coins_int = int(str(diff_coins).replace(",", "").replace("+", "").strip())
                    except ValueError:
                        coins_int = 0
                        
                    score_val = 0
                    if setting_score is not None and setting_score != "":
                        try:
                            score_val = float(str(setting_score).strip())
                        except ValueError:
                            pass
                            
                    # Determine status (〇/×)
                    if "ジャグラー" in m_name:
                        status = "〇" if (score_val >= 4.5 and coins_int >= 500) else "×"
                    else:
                        status = "〇" if coins_int >= 1000 else "×"
                        
                    safe_write_cell(ai_ws, r, 5, g_games)
                    safe_write_cell(ai_ws, r, 6, diff_coins)
                    safe_write_cell(ai_ws, r, 7, setting_score if setting_score is not None else "")
                    safe_write_cell(ai_ws, r, 8, status)
                    ai_ws.cell(r, 1).number_format = 'yyyy/mm/dd'
                    rewritten_count += 1
            except ValueError:
                continue
    print(f"Pre-calculated & optimized {rewritten_count} cells in 【AI】予想・答え合わせ.")

    # 3. Write summary narrative into 【AI】総括 sheet
    sum_sheet_name = '【AI】総括'
    actual_sheet_name = None
    for name in wb.sheetnames:
        clean_name = name.replace(" ", "").replace("　", "")
        if "AI" in clean_name and "総括" in clean_name:
            actual_sheet_name = name
            break
            
    if actual_sheet_name and actual_sheet_name in wb.sheetnames:
        sum_ws = wb[actual_sheet_name]
        print(f"Found existing summary sheet: '{actual_sheet_name}'")
    else:
        sum_ws = wb.create_sheet(sum_sheet_name)
        print(f"Created new summary sheet: '{sum_sheet_name}'")
        
    # Find actual max row in Col E or F
    real_max_sum_row = 1
    for r in range(sum_ws.max_row, 0, -1):
        val_e = sum_ws.cell(r, 5).value
        val_f = sum_ws.cell(r, 6).value
        if not is_empty_or_formula(val_e) or not is_empty_or_formula(val_f):
            real_max_sum_row = r
            break
            
    next_sum_row = real_max_sum_row + 2
    if real_max_sum_row == 1 and sum_ws.cell(1, 5).value is None and sum_ws.cell(1, 6).value is None:
        next_sum_row = 1
        
    # Clean the summary text for Excel (strictly 1 line, no paragraphs)
    text_lines = []
    for line in ai_text.split("\n"):
        line_strip = line.strip()
        if line_strip.startswith("|") or line_strip.startswith("---") or line_strip.startswith("=== AI"):
            continue
        if line_strip:
            text_lines.append(line_strip)
    single_line_summary = "".join(text_lines)
    
    # Write target date (HTML date) into Column E (5) and formatted text in F (6)
    dt_sum_cell = safe_write_cell(sum_ws, next_sum_row, 5, d_obj)
    if dt_sum_cell:
        dt_sum_cell.number_format = 'yyyy/mm/dd'
        
    safe_write_cell(sum_ws, next_sum_row, 6, single_line_summary)
    
    safe_save_workbook(wb, excel_path)
    print(f"AI Analysis results written successfully into sheets.")
    
    # 4. Save the RAW rich formatted text (with paragraphs) to a JSON file
    rich_summaries = {}
    if os.path.exists(RICH_SUMMARIES_FILE):
        try:
            with open(RICH_SUMMARIES_FILE, "r", encoding="utf-8") as f:
                rich_summaries = json.load(f)
        except Exception:
            rich_summaries = {}
            
    # Clean markdown tables only, preserving paragraphs and formatting
    clean_lines = []
    for line in ai_text.split("\n"):
        line_strip = line.strip()
        if line_strip.startswith("|") or line_strip.startswith("---"):
            continue
        clean_lines.append(line)
        
    # Normalize target_date to match standard format
    norm_target_date = normalize_date_string(target_date)
    rich_summaries[norm_target_date] = "\n".join(clean_lines).strip()
    
    with open(RICH_SUMMARIES_FILE, "w", encoding="utf-8") as f:
        json.dump(rich_summaries, f, ensure_ascii=False, indent=2)
    print(f"Saved rich paragraph summary for Web dashboard to: {RICH_SUMMARIES_FILE}")

def generate_html_dashboard(excel_path, store_name):
    """
    Reads data from the Excel workbook and generates a fully interactive, lightweight
    HTML dashboard with rich CSS styling, search filters, and statistics.
    This runs entirely in the browser with ZERO formula recalculation lag.
    """
    print(f"Generating HTML Interactive Dashboard for {store_name}...")
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    
    # 1. Read Raw Slot Records (last 1500 records)
    records = []
    if "【データ】蓄積用" in wb.sheetnames:
        ws = wb["【データ】蓄積用"]
        real_max = 1
        for r in range(ws.max_row, 0, -1):
            val1 = ws.cell(r, 1).value
            val2 = ws.cell(r, 2).value
            if val1 is not None or val2 is not None:
                real_max = r
                break
        
        start_r = max(2, real_max - 1500)
        for r in range(start_r, real_max + 1):
            dt_val = ws.cell(r, 1).value
            dt_str = normalize_date_string(dt_val)
                
            records.append({
                "date": dt_str,
                "name": str(ws.cell(r, 2).value or ""),
                "number": int(ws.cell(r, 3).value or 0),
                "games": int(ws.cell(r, 4).value or 0),
                "diff": int(ws.cell(r, 5).value or 0),
                "bb": int(ws.cell(r, 6).value or 0),
                "rb": int(ws.cell(r, 7).value or 0),
                "comb_ratio": str(ws.cell(r, 8).value or "-"),
                "bb_ratio": str(ws.cell(r, 9).value or "-"),
                "rb_ratio": str(ws.cell(r, 10).value or "-")
            })
            
    # 2. Read AI recommendations
    predictions = []
    if "【AI】予想・答え合わせ" in wb.sheetnames:
        ws = wb["【AI】予想・答え合わせ"]
        real_max = 3
        for r in range(ws.max_row, 3, -1):
            if ws.cell(r, 1).value is not None:
                real_max = r
                break
        for r in range(4, real_max + 1):
            dt_val = ws.cell(r, 1).value
            dt_str = normalize_date_string(dt_val)
                
            predictions.append({
                "date": dt_str,
                "name": str(ws.cell(r, 2).value or ""),
                "number": str(ws.cell(r, 3).value or ""),
                "reason": str(ws.cell(r, 4).value or ""),
                "games": ws.cell(r, 5).value,
                "diff": ws.cell(r, 6).value,
                "score": ws.cell(r, 7).value,
                "result": str(ws.cell(r, 8).value or "")
            })
            
    # 3. Read AI Summaries (Try rich paragraph file first, fallback to Excel)
    summaries = []
    rich_summaries = {}
    if os.path.exists(RICH_SUMMARIES_FILE):
        try:
            with open(RICH_SUMMARIES_FILE, "r", encoding="utf-8") as f:
                rich_summaries = json.load(f)
        except Exception:
            pass
            
    sum_sheet_name = None
    for name in wb.sheetnames:
        clean_name = name.replace(" ", "").replace("　", "")
        if "AI" in clean_name and "総括" in clean_name:
            sum_sheet_name = name
            break
            
    if sum_sheet_name:
        ws = wb[sum_sheet_name]
        real_max = 1
        for r in range(ws.max_row, 0, -1):
            if ws.cell(r, 5).value is not None or ws.cell(r, 6).value is not None:
                real_max = r
                break
        for r in range(1, real_max + 1):
            dt_val = ws.cell(r, 5).value
            dt_str = normalize_date_string(dt_val)
                
            val_f = ws.cell(r, 6).value
            if dt_str or val_f:
                summary_text = rich_summaries.get(dt_str, str(val_f or ""))
                summaries.append({
                    "date": dt_str,
                    "text": summary_text
                })
                
    wb.close()
    
    # Generate HTML content with updated 1-column layout, marked.js, and segmented tables
    html_template = f"""<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>【{store_name}】データ分析 AIダッシュボード</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2.0.0"></script>
    <script src="https://cdn.jsdelivr.net/npm/marked/marked.min.js"></script>
    <link href="https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;600;700&family=Noto+Sans+JP:wght@300;400;700&display=swap" rel="stylesheet">
    <style>
        body {{
            font-family: 'Outfit', 'Noto Sans JP', sans-serif;
            background-color: #0f172a;
            color: #f8fafc;
        }}
        .glass {{
            background: rgba(30, 41, 59, 0.7);
            backdrop-filter: blur(12px);
            border: 1px solid rgba(255, 255, 255, 0.05);
        }}
        .prose h1, .prose h2, .prose h3 {{
            color: #ffffff;
            font-weight: 700;
            margin-top: 1.25em;
            margin-bottom: 0.5em;
        }}
        .prose h2 {{ font-size: 1.25rem; border-left: 3px solid #10b981; padding-left: 0.5rem; }}
        .prose p {{ margin-bottom: 1em; color: #cbd5e1; font-size: 0.875rem; }}
        .prose ul {{ list-style-type: disc; padding-left: 1.25rem; margin-bottom: 1em; color: #cbd5e1; font-size: 0.875rem; }}
        .prose li {{ margin-bottom: 0.25rem; }}
        .prose strong {{ color: #f43f5e; font-weight: 600; }}
    </style>
</head>
<body class="min-h-screen p-4 md:p-8">
    <div class="max-w-4xl mx-auto space-y-6">
        
        <!-- Header -->
        <header class="flex flex-col md:flex-row justify-between items-start md:items-center p-6 glass rounded-2xl gap-4">
            <div>
                <h1 class="text-3xl font-bold bg-gradient-to-r from-emerald-400 to-cyan-400 bg-clip-text text-transparent">
                    {store_name} 分析ダッシュボード
                </h1>
                <p class="text-slate-400 text-sm mt-1">AI Data Science & Prediction Suite</p>
            </div>
            <div class="flex items-center gap-3">
                <label for="date-select" class="text-sm text-slate-400">表示日付:</label>
                <select id="date-select" class="bg-slate-800 border border-slate-700 text-white rounded-lg px-3 py-2 focus:ring-2 focus:ring-emerald-400 focus:outline-none">
                    <!-- Dates populated dynamically -->
                </select>
            </div>
        </header>

        <!-- Stats Overview Grid -->
        <div class="grid grid-cols-2 md:grid-cols-4 gap-4">
            <div class="p-4 md:p-6 glass rounded-2xl text-center md:text-left">
                <h3 class="text-slate-400 text-xs font-semibold uppercase tracking-wider">本日総差枚</h3>
                <p id="stat-total-diff" class="text-xl md:text-2xl font-bold mt-2 text-white">0 枚</p>
            </div>
            <div class="p-4 md:p-6 glass rounded-2xl text-center md:text-left">
                <h3 class="text-slate-400 text-xs font-semibold uppercase tracking-wider">本日平均差枚</h3>
                <p id="stat-avg-diff" class="text-xl md:text-2xl font-bold mt-2 text-white">0 枚</p>
            </div>
            <div class="p-4 md:p-6 glass rounded-2xl text-center md:text-left">
                <h3 class="text-slate-400 text-xs font-semibold uppercase tracking-wider">本日勝率</h3>
                <p id="stat-win-rate" class="text-xl md:text-2xl font-bold mt-2 text-white">0%</p>
            </div>
            <div class="p-4 md:p-6 glass rounded-2xl text-center md:text-left">
                <h3 class="text-slate-400 text-xs font-semibold uppercase tracking-wider">AI予測 通算勝率</h3>
                <p id="stat-ai-accuracy" class="text-xl md:text-2xl font-bold mt-2 text-emerald-400">0%</p>
            </div>
        </div>

        <!-- ① AIによる営業総括・心理分析 -->
        <section class="p-6 glass rounded-2xl space-y-4">
            <div class="flex items-center gap-2">
                <span class="p-1.5 bg-emerald-500/10 text-emerald-400 rounded-lg text-sm">💡</span>
                <h2 class="text-xl font-bold">AIによる営業総括・心理分析</h2>
            </div>
            <!-- Markdown parsed and loaded dynamically -->
            <div id="ai-summary-text" class="prose max-w-none p-4 bg-slate-900/50 rounded-xl border border-slate-800">
                該当日の総括データがありません。
            </div>
        </section>

        <!-- ② 主要機種の差枚状況ランキング (グラフ) -->
        <section class="p-6 glass rounded-2xl space-y-4">
            <div class="flex items-center gap-2">
                <span class="p-1.5 bg-indigo-500/10 text-indigo-400 rounded-lg text-sm">📊</span>
                <h2 class="text-xl font-bold font-semibold">主要機種の差枚状況ランキング（平均値）</h2>
            </div>
            <div class="h-80 md:h-96 w-full">
                <canvas id="chart-machines"></canvas>
            </div>
        </section>

        <!-- ③ 次回（明日）のAI推奨台 (Segmented) -->
        <section class="p-6 glass rounded-2xl space-y-4">
            <div class="flex items-center gap-2">
                <span class="p-1.5 bg-cyan-500/10 text-cyan-400 rounded-lg text-sm">🎯</span>
                <h2 class="text-xl font-bold">次回（明日）のAI推奨台</h2>
            </div>
            
            <div class="grid grid-cols-1 md:grid-cols-2 gap-6">
                <!-- Juggler Section -->
                <div class="space-y-2">
                    <h3 class="text-emerald-400 font-bold text-sm border-b border-emerald-500/30 pb-1">🤡 ジャグラー系 (5台)</h3>
                    <div class="overflow-x-auto">
                        <table class="w-full text-left border-collapse text-xs">
                            <thead>
                                <tr class="border-b border-slate-800 text-slate-400">
                                    <th class="py-2 px-2">機種名</th>
                                    <th class="py-2 px-2">台番</th>
                                    <th class="py-2 px-2">根拠</th>
                                </tr>
                            </thead>
                            <tbody id="pred-juggler-body" class="divide-y divide-slate-800/50">
                                <!-- Populated dynamically -->
                            </tbody>
                        </table>
                    </div>
                </div>

                <!-- Others/Smaslot Section -->
                <div class="space-y-2">
                    <h3 class="text-cyan-400 font-bold text-sm border-b border-cyan-500/30 pb-1">⚡ スマスロ・その他 (5台)</h3>
                    <div class="overflow-x-auto">
                        <table class="w-full text-left border-collapse text-xs">
                            <thead>
                                <tr class="border-b border-slate-800 text-slate-400">
                                    <th class="py-2 px-2">機種名</th>
                                    <th class="py-2 px-2">台番</th>
                                    <th class="py-2 px-2">根拠</th>
                                </tr>
                            </thead>
                            <tbody id="pred-others-body" class="divide-y divide-slate-800/50">
                                <!-- Populated dynamically -->
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>
        </section>

        <!-- ④ 前日のAI予測答え合わせ (Segmented) -->
        <section class="p-6 glass rounded-2xl space-y-4">
            <div class="flex items-center gap-2">
                <span class="p-1.5 bg-amber-500/10 text-amber-400 rounded-lg text-sm">📈</span>
                <h2 class="text-xl font-bold">前日のAI予測答え合わせ</h2>
            </div>
            
            <div class="grid grid-cols-1 md:grid-cols-2 gap-6">
                <!-- Juggler Section -->
                <div class="space-y-2">
                    <h3 class="text-emerald-400 font-bold text-sm border-b border-emerald-500/30 pb-1">🤡 ジャグラー系 結果</h3>
                    <div class="overflow-x-auto">
                        <table class="w-full text-left border-collapse text-xs">
                            <thead>
                                <tr class="border-b border-slate-800 text-slate-400">
                                    <th class="py-2 px-2">台番 (機種)</th>
                                    <th class="py-2 px-2">差枚</th>
                                    <th class="py-2 px-2">設定スコア</th>
                                    <th class="py-2 px-2">結果</th>
                                </tr>
                            </thead>
                            <tbody id="ans-juggler-body" class="divide-y divide-slate-800/50">
                                <!-- Populated dynamically -->
                            </tbody>
                        </table>
                    </div>
                </div>

                <!-- Others/Smaslot Section -->
                <div class="space-y-2">
                    <h3 class="text-cyan-400 font-bold text-sm border-b border-cyan-500/30 pb-1">⚡ スマスロ・その他 結果</h3>
                    <div class="overflow-x-auto">
                        <table class="w-full text-left border-collapse text-xs">
                            <thead>
                                <tr class="border-b border-slate-800 text-slate-400">
                                    <th class="py-2 px-2">台番 (機種)</th>
                                    <th class="py-2 px-2">差枚</th>
                                    <th class="py-2 px-2">設定スコア</th>
                                    <th class="py-2 px-2">結果</th>
                                </tr>
                            </thead>
                            <tbody id="ans-others-body" class="divide-y divide-slate-800/50">
                                <!-- Populated dynamically -->
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>
        </section>

        <!-- ⑤ 本日（選択日）の全台データ一覧 -->
        <section class="p-6 glass rounded-2xl space-y-4">
            <div class="flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
                <div class="flex items-center gap-2">
                    <span class="p-1.5 bg-slate-500/10 text-slate-400 rounded-lg text-sm">📋</span>
                    <h2 class="text-xl font-bold">本日（選択日）の全台データ一覧</h2>
                </div>
                <input type="text" id="search-input" placeholder="機種名や台番号で検索..." class="w-full md:w-64 bg-slate-800 border border-slate-700 text-white rounded-lg px-3 py-1.5 focus:ring-2 focus:ring-emerald-400 focus:outline-none text-sm">
            </div>
            <div class="overflow-x-auto max-h-96 overflow-y-auto">
                <table class="w-full text-left border-collapse text-sm">
                    <thead class="sticky top-0 bg-slate-900 border-b border-slate-800 text-slate-400 z-10">
                        <tr>
                            <th class="py-3 px-2">台番号</th>
                            <th class="py-3 px-2">機種名</th>
                            <th class="py-3 px-2">G数</th>
                            <th class="py-3 px-2">差枚</th>
                            <th class="py-3 px-2">BB</th>
                            <th class="py-3 px-2">RB</th>
                            <th class="py-3 px-2">合算</th>
                            <th class="py-3 px-2">BB確率</th>
                            <th class="py-3 px-2">REG確率</th>
                        </tr>
                    </thead>
                    <tbody id="raw-data-table-body" class="divide-y divide-slate-800/50">
                        <!-- Populated dynamically -->
                    </tbody>
                </table>
            </div>
        </section>

    </div>

    <!-- Embedding JSON Data -->
    <script>
        const rawRecords = {json.dumps(records, ensure_ascii=False)};
        const rawPredictions = {json.dumps(predictions, ensure_ascii=False)};
        const rawSummaries = {json.dumps(summaries, ensure_ascii=False)};
        
        let machineChart = null;
        Chart.register(ChartDataLabels);

        document.addEventListener('DOMContentLoaded', () => {{
            initializeDashboard();
        }});

        function initializeDashboard() {{
            const dates = [...new Set(rawRecords.map(r => r.date))].sort().reverse();
            const dateSelect = document.getElementById('date-select');
            
            dates.forEach(d => {{
                const opt = document.createElement('option');
                opt.value = d;
                opt.textContent = d;
                dateSelect.appendChild(opt);
            }});

            dateSelect.addEventListener('change', (e) => {{
                updateDashboardForDate(e.target.value);
            }});

            if (dates.length > 0) {{
                dateSelect.value = dates[0];
                updateDashboardForDate(dates[0]);
            }}
            
            calculateOverallAIAccuracy();
        }}

        function isJugglerMachine(name) {{
            const jugKeywords = ['ジャグラー', 'ジャグ', 'マイJ', 'ファンキー', 'アイム', 'ゴーゴー', 'ハッピー', 'ミラクル'];
            return jugKeywords.some(keyword => name.includes(keyword));
        }}

        // Deduplicates records by taking the LAST registered prediction for a unique (date, number)
        function deduplicatePredictions(preds) {{
            const uniqueMap = {{}};
            preds.forEach(p => {{
                const key = p.date + '_' + p.number;
                uniqueMap[key] = p; // Keep the latest one
            }});
            return Object.values(uniqueMap);
        }}

        function updateDashboardForDate(targetDate) {{
            const filteredRecords = rawRecords.filter(r => r.date === targetDate);
            
            // 1. Update Stat boxes
            const totalDiff = filteredRecords.reduce((acc, curr) => acc + curr.diff, 0);
            const avgDiff = filteredRecords.length > 0 ? Math.round(totalDiff / filteredRecords.length) : 0;
            const winningSlots = filteredRecords.filter(r => r.diff > 0).length;
            const winRate = filteredRecords.length > 0 ? Math.round((winningSlots / filteredRecords.length) * 100) : 0;

            document.getElementById('stat-total-diff').textContent = totalDiff.toLocaleString() + ' 枚';
            document.getElementById('stat-total-diff').className = 'text-xl md:text-2xl font-bold mt-2 ' + (totalDiff >= 0 ? 'text-emerald-400' : 'text-rose-400');
            document.getElementById('stat-avg-diff').textContent = avgDiff.toLocaleString() + ' 枚';
            document.getElementById('stat-win-rate').textContent = winRate + '%';

            // 2. Load AI Summary text (Markdown render!)
            const summaryObj = rawSummaries.find(s => s.date === targetDate);
            if (summaryObj) {{
                document.getElementById('ai-summary-text').innerHTML = marked.parse(summaryObj.text);
            }} else {{
                document.getElementById('ai-summary-text').innerHTML = '<p class="text-slate-500">この日のAI営業総括データはありません。</p>';
            }}

            // 3. Load Predictions for the next day (Segmented Juggler vs Others & Deduplicated!)
            const d = new Date(targetDate);
            d.setDate(d.getDate() + 1);
            const tomorrowStr = d.getFullYear() + '/' + String(d.getMonth() + 1).padStart(2, '0') + '/' + String(d.getDate()).padStart(2, '0');
            
            const dedupedPredictions = deduplicatePredictions(rawPredictions);
            const filteredPredictions = dedupedPredictions.filter(p => p.date === tomorrowStr);
            
            const predJugglerBody = document.getElementById('pred-juggler-body');
            const predOthersBody = document.getElementById('pred-others-body');
            predJugglerBody.innerHTML = '';
            predOthersBody.innerHTML = '';
            
            const jugPredictions = filteredPredictions.filter(p => isJugglerMachine(p.name)).slice(0, 5);
            const otherPredictions = filteredPredictions.filter(p => !isJugglerMachine(p.name)).slice(0, 5);
            
            if (jugPredictions.length === 0) {{
                predJugglerBody.innerHTML = '<tr><td colspan="3" class="py-3 text-center text-slate-500">ジャグラー推奨台はありません。</td></tr>';
            }} else {{
                jugPredictions.forEach(p => {{
                    const tr = document.createElement('tr');
                    tr.className = 'hover:bg-slate-800/30';
                    tr.innerHTML = `
                        <td class="py-2.5 px-2 font-semibold text-slate-200">${{p.name}}</td>
                        <td class="py-2.5 px-2 text-emerald-400 font-bold font-mono">${{p.number}}</td>
                        <td class="py-2.5 px-2 text-xs text-slate-300">${{p.reason}}</td>
                    `;
                    predJugglerBody.appendChild(tr);
                }});
            }}

            if (otherPredictions.length === 0) {{
                predOthersBody.innerHTML = '<tr><td colspan="3" class="py-3 text-center text-slate-500">スマスロ・その他推奨台はありません。</td></tr>';
            }} else {{
                otherPredictions.forEach(p => {{
                    const tr = document.createElement('tr');
                    tr.className = 'hover:bg-slate-800/30';
                    tr.innerHTML = `
                        <td class="py-2.5 px-2 font-semibold text-slate-200">${{p.name}}</td>
                        <td class="py-2.5 px-2 text-cyan-400 font-bold font-mono">${{p.number}}</td>
                        <td class="py-2.5 px-2 text-xs text-slate-300">${{p.reason}}</td>
                    `;
                    predOthersBody.appendChild(tr);
                }});
            }}

            // 4. Load Answers for targetDate predictions (Segmented & Deduplicated!)
            const filteredAnswers = dedupedPredictions.filter(p => p.date === targetDate);
            
            const ansJugglerBody = document.getElementById('ans-juggler-body');
            const ansOthersBody = document.getElementById('ans-others-body');
            ansJugglerBody.innerHTML = '';
            ansOthersBody.innerHTML = '';
            
            const jugAnswers = filteredAnswers.filter(a => isJugglerMachine(a.name)).slice(0, 5);
            const otherAnswers = filteredAnswers.filter(a => !isJugglerMachine(a.name)).slice(0, 5);
            
            const renderAnswerRow = (a, isJuggler) => {{
                const diffVal = a.diff !== null && a.diff !== undefined ? parseInt(a.diff) : null;
                const diffText = diffVal !== null ? (diffVal >= 0 ? '+' + diffVal : diffVal) + ' 枚' : '-';
                const diffClass = diffVal !== null ? (diffVal >= 0 ? 'text-emerald-400 font-bold' : 'text-rose-400') : 'text-slate-400';
                
                let scoreText = '-';
                if (a.score !== null && a.score !== undefined && a.score !== "") {{
                    const val = parseFloat(a.score);
                    scoreText = isNaN(val) ? String(a.score) : val.toFixed(1);
                }}
                
                let finalResult = a.result;
                if (isJuggler) {{
                    if (a.games === null || a.games === undefined || a.games === "" || a.games === 0) {{
                        finalResult = '-';
                    }} else {{
                        const scoreVal = a.score !== null && a.score !== undefined ? parseFloat(a.score) : 0;
                        if (!isNaN(scoreVal) && scoreVal >= 4.5 && diffVal !== null && diffVal >= 500) {{
                            finalResult = '〇';
                        }} else {{
                            finalResult = '×';
                        }}
                    }}
                }}
                
                const resultBadge = finalResult === '〇' ? '<span class="px-2 py-0.5 bg-emerald-500/20 text-emerald-400 text-[10px] rounded-full font-bold">〇</span>' : 
                                   (finalResult === '×' ? '<span class="px-2 py-0.5 bg-rose-500/20 text-rose-400 text-[10px] rounded-full font-bold">×</span>' : '-');
                return `
                    <td class="py-2.5 px-2 text-slate-200 font-semibold">${{a.number}}番 (${{a.name}})</td>
                    <td class="py-2.5 px-2 font-mono ${{diffClass}}">${{diffText}}</td>
                    <td class="py-2.5 px-2 text-slate-300 font-mono">${{scoreText}}</td>
                    <td class="py-2.5 px-2">${{resultBadge}}</td>
                `;
            }};

            if (jugAnswers.length === 0) {{
                ansJugglerBody.innerHTML = '<tr><td colspan="4" class="py-3 text-center text-slate-500">答え合わせはありません。</td></tr>';
            }} else {{
                jugAnswers.forEach(a => {{
                    const tr = document.createElement('tr');
                    tr.className = 'hover:bg-slate-800/30';
                    tr.innerHTML = renderAnswerRow(a, true);
                    ansJugglerBody.appendChild(tr);
                }});
            }}

            if (otherAnswers.length === 0) {{
                ansOthersBody.innerHTML = '<tr><td colspan="4" class="py-3 text-center text-slate-500">答え合わせはありません。</td></tr>';
            }} else {{
                otherAnswers.forEach(a => {{
                    const tr = document.createElement('tr');
                    tr.className = 'hover:bg-slate-800/30';
                    tr.innerHTML = renderAnswerRow(a, false);
                    ansOthersBody.appendChild(tr);
                }});
            }}

            // 5. Populate Raw Data Table
            populateRawTable(filteredRecords);

            // Setup Search
            const searchInput = document.getElementById('search-input');
            const newSearchInput = searchInput.cloneNode(true);
            searchInput.parentNode.replaceChild(newSearchInput, searchInput);
            newSearchInput.value = '';
            newSearchInput.addEventListener('input', (e) => {{
                const q = e.target.value.toLowerCase();
                const searched = filteredRecords.filter(r => 
                    r.name.toLowerCase().includes(q) || String(r.number).includes(q)
                );
                populateRawTable(searched);
            }});

            // 6. Draw Chart
            drawMachineChart(filteredRecords);
        }}

        function populateRawTable(data) {{
            const tbody = document.getElementById('raw-data-table-body');
            tbody.innerHTML = '';
            if (data.length === 0) {{
                tbody.innerHTML = '<tr><td colspan="9" class="py-8 text-center text-slate-500">データが見つかりません。</td></tr>';
                return;
            }}
            data.forEach(r => {{
                const tr = document.createElement('tr');
                tr.className = 'hover:bg-slate-800/30 transition-colors';
                const diffClass = r.diff >= 0 ? 'text-emerald-400 font-bold' : 'text-rose-400';
                const diffText = r.diff >= 0 ? '+' + r.diff : r.diff;
                tr.innerHTML = `
                    <td class="py-3 px-2 text-slate-400 font-semibold font-mono">${{r.number}}</td>
                    <td class="py-3 px-2 font-bold text-slate-200">${{r.name}}</td>
                    <td class="py-3 px-2 text-slate-300 font-mono">${{r.games}} G</td>
                    <td class="py-3 px-2 font-mono ${{diffClass}}">${{diffText}} 枚</td>
                    <td class="py-3 px-2 text-slate-400 font-mono">${{r.bb}}</td>
                    <td class="py-3 px-2 text-slate-400 font-mono">${{r.rb}}</td>
                    <td class="py-3 px-2 text-slate-300 font-mono">${{r.comb_ratio}}</td>
                    <td class="py-3 px-2 text-slate-400 font-mono">${{r.bb_ratio}}</td>
                    <td class="py-3 px-2 text-slate-400 font-mono">${{r.rb_ratio}}</td>
                `;
                tbody.appendChild(tr);
            }});
        }}

        function calculateOverallAIAccuracy() {{
            const deduped = deduplicatePredictions(rawPredictions);
            const evaluated = deduped.filter(p => p.result === '〇' || p.result === '×');
            if (evaluated.length === 0) {{
                document.getElementById('stat-ai-accuracy').textContent = '- %';
                return;
            }}
            const correct = evaluated.filter(p => p.result === '〇').length;
            const accuracy = Math.round((correct / evaluated.length) * 100);
            document.getElementById('stat-ai-accuracy').textContent = accuracy + '%';
        }}

        function drawMachineChart(records) {{
            const machineData = {{}};
            records.forEach(r => {{
                if (!machineData[r.name]) {{
                    machineData[r.name] = {{ totalDiff: 0, count: 0 }};
                }}
                machineData[r.name].totalDiff += r.diff;
                machineData[r.name].count += 1;
            }});

            const chartItems = Object.keys(machineData).map(k => ({{
                name: k,
                avgDiff: Math.round(machineData[k].totalDiff / machineData[k].count)
            }})).sort((a, b) => b.avgDiff - a.avgDiff).slice(0, 10);

            const labels = chartItems.map(item => item.name);
            const dataVals = chartItems.map(item => item.avgDiff);

            const ctx = document.getElementById('chart-machines').getContext('2d');
            if (machineChart) {{
                machineChart.destroy();
            }}

            machineChart = new Chart(ctx, {{
                type: 'bar',
                data: {{
                    labels: labels,
                    datasets: [{{
                        label: '平均差枚数 (枚)',
                        data: dataVals,
                        backgroundColor: dataVals.map(v => v >= 0 ? 'rgba(52, 211, 153, 0.6)' : 'rgba(248, 113, 113, 0.6)'),
                        borderColor: dataVals.map(v => v >= 0 ? 'rgba(52, 211, 153, 1)' : 'rgba(248, 113, 113, 1)'),
                        borderWidth: 1.5,
                        borderRadius: 6
                    }}]
                }},
                options: {{
                    responsive: true,
                    maintainAspectRatio: false,
                    layout: {{
                        padding: {{
                            top: 25
                        }}
                    }},
                    plugins: {{
                        legend: {{ display: false }},
                        datalabels: {{
                            anchor: 'end',
                            align: 'top',
                            color: '#e2e8f0',
                            font: {{
                                weight: 'bold',
                                size: 10
                            }},
                            formatter: function(value) {{
                                return (value >= 0 ? '+' : '') + value;
                            }}
                        }}
                    }},
                    scales: {{
                        y: {{
                            grid: {{ color: 'rgba(255, 255, 255, 0.05)' }},
                            ticks: {{ color: '#94a3b8' }}
                        }},
                        x: {{
                            grid: {{ display: false }},
                            ticks: {{ 
                                color: '#94a3b8',
                                font: {{ size: 10 }}
                            }}
                        }}
                    }}
                }}
            }});
        }}
    </script>
</body>
</html>
"""
    
    output_html_path = os.path.join(os.path.dirname(excel_path), f"分析ダッシュボード_{store_name}.html")
    with open(output_html_path, "w", encoding="utf-8") as f:
        f.write(html_template)
    print(f"HTML Interactive Dashboard generated successfully: {output_html_path}")
    
    # Also save as index.html for root page on GitHub Pages
    index_html_path = os.path.join(os.path.dirname(excel_path), "index.html")
    with open(index_html_path, "w", encoding="utf-8") as f:
        f.write(html_template)
    print(f"HTML index.html generated successfully for GitHub Pages.")

def deploy_to_github():
    """
    Deploys the generated HTML dashboard and raw summaries to GitHub Pages automatically
    if the local directory is initialized as a Git repository.
    """
    print("\nDeploying dashboard to GitHub Pages...")
    try:
        if not os.path.exists(".git"):
            print("Notice: Git repository not initialized. Skipping auto-deploy.")
            return
            
        # Add generated html and rich summaries
        os.system("git add 分析ダッシュボード_*.html index.html data_input/ai_summaries_rich.json")
        
        # Commit changes
        commit_msg = f"Auto-update dashboard: {datetime.datetime.now().strftime('%Y/%m/%d %H:%M:%S')}"
        os.system(f'git commit -m "{commit_msg}"')
        
        # Push to remote (Try main first, then master)
        res = os.system("git push origin main")
        if res != 0:
            os.system("git push origin master")
            
        print("Successfully deployed to GitHub Pages! It will be live in a few seconds.")
    except Exception as e:
        print(f"Warning: Failed to deploy to GitHub: {e}")

def main():
    print("=== Pachislot Data Analyzer Auto Tool ===")
    config = load_config()
    
    api_key = config.get("GEMINI_API_KEY")
    if not api_key or api_key == "YOUR_GEMINI_API_KEY_HERE":
        print("Error: Gemini API key is missing in config.json. Please update the key.")
        sys.exit(1)
        
    allowed_exts = [".txt", ".html", ".htm"]
    files = [f for f in os.listdir("data_input") if os.path.splitext(f)[1].lower() in allowed_exts]
    if not files:
        print("\nNo input data files found in 'data_input/' folder.")
        print("Please place a text or downloaded HTML file (e.g. 'A店_20260714.html') inside 'data_input/' and try again.")
        sys.exit(0)
        
    print("\nSelect the file to process:")
    for idx, f in enumerate(files):
        print(f"{idx + 1}: {f}")
        
    if len(files) == 1:
        print(f"\nOnly one file found: '{files[0]}'. Automatically selecting it.")
        sel = 0
    else:
        # Check if stdin is interactive
        if sys.stdin.isatty():
            try:
                sel = int(input("\nEnter selection number: ")) - 1
                if sel < 0 or sel >= len(files):
                    raise ValueError()
            except ValueError:
                print("Invalid selection.")
                sys.exit(1)
        else:
            print("\nNon-interactive shell detected. Defaulting to 1st file.")
            sel = 0
        
    selected_file = files[sel]
    filepath = os.path.join("data_input", selected_file)
    
    filename_wo_ext = os.path.splitext(selected_file)[0]
    match = re.match(r'^([^_ 　-]+)[_ 　-]+(\d{8})$', filename_wo_ext)
    if not match:
        print(f"\nCould not determine Store and Date from filename '{selected_file}'.")
        store_name = input("Enter Store Name (e.g. A店): ")
        target_date_raw = input("Enter Date (e.g. 20260714): ")
    else:
        store_name = match.group(1)
        target_date_raw = match.group(2)
        
    try:
        target_date = datetime.datetime.strptime(target_date_raw, "%Y%m%d").strftime("%Y/%m/%d")
    except ValueError:
        print(f"Error: Invalid date format '{target_date_raw}'. Expected YYYYMMDD.")
        sys.exit(1)
        
    store_info = config.get("STORES", {}).get(store_name)
    if not store_info:
        print(f"Error: Store '{store_name}' is not registered in config.json.")
        sys.exit(1)
        
    excel_file = store_info.get("excel_file")
    print(f"\nTarget Store: {store_name}")
    print(f"Target Date: {target_date}")
    print(f"Target Excel: {excel_file}")
    
    ext = os.path.splitext(selected_file)[1].lower()
    if ext in [".html", ".htm"]:
        parsed_data = parse_html_data(filepath)
    else:
        parsed_data = parse_text_data(filepath)
        
    if not parsed_data:
        print("Error: No data rows parsed. Check the input file format.")
        sys.exit(1)
        
    update_excel_data(excel_file, target_date, parsed_data)
    
    print("\nExtracting Excel metrics for AI context...")
    context = prepare_ai_context(excel_file, target_date)
    
    print("\nRunning Gemini AI analysis...")
    ai_text = run_gemini_analysis(api_key, context, target_date)
    
    ref_out = os.path.join("data_input", f"{filename_wo_ext}_ai_report.md")
    with open(ref_out, "w", encoding="utf-8") as f:
        f.write(ai_text)
    print(f"AI raw report saved to: {ref_out}")
    
    print("\nWriting AI recommendations back to Excel...")
    write_ai_results_to_excel(excel_file, target_date, ai_text)
    
    # 4. Generate the fully interactive HTML Dashboard for instant check
    generate_html_dashboard(excel_file, store_name)
    
    # 5. Automatically deploy to GitHub Pages
    deploy_to_github()
    
    os.makedirs(os.path.join("data_input", "processed"), exist_ok=True)
    os.rename(filepath, os.path.join("data_input", "processed", selected_file))
    print(f"\n=== Process complete! Input file moved to processed/ backup folder. ===")

if __name__ == "__main__":
    main()
