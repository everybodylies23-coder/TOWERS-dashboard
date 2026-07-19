import datetime

def build_maps(wb):
    """Build specs and machine configuration maps from master sheets."""
    print("Building specs and machine configuration maps...")
    ws_spec = wb['【マスター】スペック']
    specs_dict = {}
    for r in range(2, 40):
        m_name = ws_spec.cell(r, 1).value
        if m_name:
            row_vals = [ws_spec.cell(r, c).value for c in range(2, 33)]
            specs_dict[m_name] = row_vals
            
    ws_setting = wb['【マスター】設定']
    settings_dict = {}
    for r in range(2, 1000):
        num = ws_setting.cell(r, 1).value
        name = ws_setting.cell(r, 2).value
        if num is not None:
            settings_dict[int(num)] = name
            
    return specs_dict, settings_dict

def format_probability(games, count):
    """Formats a ratio as '1/X.X' or '-' if invalid or zero."""
    if not games or not count or games <= 0 or count <= 0:
        return "-"
    ratio = games / count
    return f"1/{ratio:.1f}"

def extract_history(ws, max_row):
    """
    Reads existing rows from the sheet to build processed_history.
    Requires Column A (Date), B (Machine Name), C (Machine Num), D (Games), E (Diff), F (BB), G (RB).
    """
    history = []
    for r in range(2, max_row + 1):
        d_val = ws.cell(r, 1).value
        m_num = ws.cell(r, 3).value
        m_name_raw = ws.cell(r, 2).value  # Column B: machine name
        if d_val and m_num:
            d_val = _to_datetime(d_val)
            if d_val is not None:
                history.append({
                    "row_num": r,
                    "date": d_val,
                    "machine_name": str(m_name_raw) if m_name_raw else "",
                    "machine_number": int(m_num),
                    "g_games": _safe_int(ws.cell(r, 4).value),
                    "diff_coins": _safe_int(ws.cell(r, 5).value),
                    "bb_count": _safe_int(ws.cell(r, 6).value),
                    "rb_count": _safe_int(ws.cell(r, 7).value)
                })
    return history

def _to_datetime(val):
    """Convert a value to datetime."""
    if isinstance(val, datetime.datetime):
        return val
    if isinstance(val, datetime.date):
        return datetime.datetime.combine(val, datetime.time())
    if isinstance(val, str):
        for fmt in ["%Y/%m/%d %H:%M:%S", "%Y/%m/%d", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d"]:
            try:
                return datetime.datetime.strptime(val.strip(), fmt)
            except ValueError:
                continue
    return None

def _safe_int(val):
    """Safely convert a cell value to int."""
    if val is None:
        return 0
    val_str = str(val).replace("=", "").replace("+", "").replace(",", "").strip()
    if not val_str:
        return 0
    try:
        return int(float(val_str))
    except (ValueError, TypeError):
        return 0

def calculate_kaf_for_row(ws, curr_r, row_data, settings_dict, specs_dict, processed_history):
    """
    Calculates and writes K(11) through AF(32) columns for a specific row.
    
    Column mapping (matching the original template MAP formulas):
      K  (11) = 場所 (machine name from settings by number)
      L  (12) = 機械割 (mechanical payout)
      M  (13) = 前日の差枚 (previous day's diff_coins for same machine)
      N  (14) = 前々日の差枚 (2 days ago diff_coins for same machine)
      O  (15) = 連凹日 (consecutive negative days count)
      P  (16) = 末尾 (last digit of machine number)
      Q  (17) = ジャグラー識別 (1 if juggler, 0 if not)
      R  (18) = ブドウ逆算回数 (grape inverse calculation)
      S  (19) = ブドウ確率 (grape probability = G_games / R)
      T  (20) = BIG確率 (G_games / BB)
      U  (21) = REG確率 (G_games / RB)
      V  (22) = 合算確率 (G_games / (BB+RB))
      W  (23) = BIG Pt
      X  (24) = REG Pt
      Y  (25) = 合算 Pt
      Z  (26) = ブドウ Pt
      AA (27) = 平均点
      AB (28) = 重み係数
      AC (29) = 最終設定スコア
      AD (30) = 評価スコア
      AE (31) = タイスコア
      AF (32) = メモ/内訳
    """
    m_num = row_data["machine_number"]
    # Column B's machine_name is the ground truth for juggler detection
    machine_name_b = row_data.get("machine_name", "")
    # K column uses settings lookup
    m_name_setting = settings_dict.get(m_num, "×")
    
    gx = row_data["g_games"]
    dx = row_data["diff_coins"]
    bx = row_data["bb_count"]
    rx = row_data["rb_count"]
    current_date = row_data["date"]
    
    # --- Col K (11): 場所 ---
    ws.cell(curr_r, 11, m_name_setting)
    
    # --- Col L (12): 機械割 = (gx*3 + dx) / (gx*3) ---
    mechanical_payout = round((gx * 3 + dx) / (gx * 3), 4) if gx > 0 else 0
    ws.cell(curr_r, 12, mechanical_payout)
    
    # --- Find previous entries for the same machine (by C=machine_number) ---
    prev_rows = [p for p in processed_history if p["machine_number"] == m_num and p["date"] < current_date]
    prev_rows.sort(key=lambda x: x["date"])
    
    # --- Col M (13): 前日の差枚 = diff_coins of the most recent previous date ---
    cell_m = ws.cell(curr_r, 13)
    if prev_rows:
        cell_m.value = prev_rows[-1]["diff_coins"]
    else:
        cell_m.value = ""
    cell_m.number_format = '#,##0'  # Force numeric format, NOT date
    
    # --- Col N (14): 前々日の差枚 = diff_coins of the 2nd most recent previous date ---
    cell_n = ws.cell(curr_r, 14)
    if len(prev_rows) >= 2:
        cell_n.value = prev_rows[-2]["diff_coins"]
    else:
        cell_n.value = ""
    cell_n.number_format = '#,##0'  # Force numeric format, NOT date
    
    # --- Col O (15): 連凹日 ---
    # Original: COUNTIFS(dates > prev_date, dates < current_date, same machine, diff >= 0)
    if prev_rows:
        prev_date = prev_rows[-1]["date"]
        count_positive = 0
        for p in processed_history:
            if p["machine_number"] == m_num and p["date"] > prev_date and p["date"] < current_date and p["diff_coins"] >= 0:
                count_positive += 1
        ws.cell(curr_r, 15, count_positive)
    else:
        ws.cell(curr_r, 15, "")
    
    # --- Col P (16): 末尾 ---
    tail_num = int(str(m_num)[-1])
    ws.cell(curr_r, 16, tail_num)
    
    # --- Col Q (17): ジャグラー識別 ---
    # Use column B machine_name (NOT settings name) to detect Juggler
    is_jug = 1 if "ジャグラー" in machine_name_b else 0
    ws.cell(curr_r, 17, is_jug)
    
    # --- Juggler-specific calculations (R through AC) ---
    payout_coeff = ""   # R (18)
    grape_prob = ""      # S (19)
    bb_ratio = ""        # T (20)
    rb_ratio = ""        # U (21)
    comb_ratio = ""      # V (22)
    bb_setting = ""      # W (23)
    rb_setting = ""      # X (24)
    comb_setting = ""    # Y (25)
    grape_setting = ""   # Z (26)
    avg_score = ""       # AA (27)
    weight_val = ""      # AB (28)
    final_score = ""     # AC (29)
    
    # For specs lookup, use column B machine_name (same as original VLOOKUP(m, specs, ...))
    specs_key = machine_name_b
    
    if is_jug == 1 and specs_key in specs_dict:
        specs = specs_dict[specs_key]
        bb_inv = float(specs[0]) if specs[0] else 0    # spec col 2
        rb_inv = float(specs[1]) if specs[1] else 0    # spec col 3
        spec_divisor = float(specs[2]) if specs[2] else 1  # spec col 4
        
        # R (18): ブドウ逆算回数
        if gx > 0 and spec_divisor != 0:
            payout_coeff = ((gx * 3) + dx - (bx * bb_inv) - (rx * rb_inv) - (gx / 7.3 * 3) - (gx / 33 * 2)) / spec_divisor
        
        # S (19): ブドウ確率 = gx / R
        if isinstance(payout_coeff, (int, float)) and payout_coeff > 0:
            grape_prob = gx / payout_coeff
        
        # T (20): BIG確率 = gx / bx
        if bx > 0:
            bb_ratio = gx / bx
        
        # U (21): REG確率 = gx / rx
        if rx > 0:
            rb_ratio = gx / rx
        
        # V (22): 合算確率 = gx / (bx + rx)
        if (bx + rx) > 0:
            comb_ratio = gx / (bx + rx)
        
        # Setting point lookup function
        def find_setting_pt(val, spec_start):
            """
            IF(val <= spec[start], 7, IF(val <= spec[start+1], 6, ... IF(val <= spec[start+6], 1, 0)))
            """
            if val == "" or val is None:
                return ""
            for i in range(7):
                spec_val = specs[spec_start + i]
                if spec_val is not None and spec_val != "":
                    try:
                        if float(val) <= float(spec_val):
                            return 7 - i
                    except (ValueError, TypeError):
                        pass
            return 0
        
        # W (23): BIG Pt (uses T=bb_ratio, spec cols 5-11 → indices 3-9)
        bb_setting = find_setting_pt(bb_ratio, 3)
        
        # X (24): REG Pt (uses U=rb_ratio, spec cols 12-18 → indices 10-16)
        rb_setting = find_setting_pt(rb_ratio, 10)
        
        # Y (25): 合算 Pt (uses V=comb_ratio, spec cols 19-25 → indices 17-23)
        comb_setting = find_setting_pt(comb_ratio, 17)
        
        # Z (26): ブドウ Pt (uses S=grape_prob, spec cols 26-32 → indices 24-30)
        grape_setting = find_setting_pt(grape_prob, 24)
        
        # AA (27): 平均点 = (W + X + Y + Z) / 4
        pts = [bb_setting, rb_setting, comb_setting, grape_setting]
        valid_pts = [p for p in pts if isinstance(p, (int, float))]
        if len(valid_pts) == 4:
            avg_score = sum(valid_pts) / 4
        elif valid_pts:
            avg_score = sum(valid_pts) / len(valid_pts)
        
        # AB (28): 重み係数
        if gx > 0:
            if gx >= 7000:
                weight_val = 1.0
            elif gx >= 5000:
                weight_val = 0.8
            elif gx >= 3000:
                weight_val = 0.6
            else:
                weight_val = 0.5
        
        # AC (29): 最終設定スコア = avg * weight
        if isinstance(avg_score, (int, float)) and isinstance(weight_val, (int, float)):
            final_score = avg_score * weight_val
    else:
        # Non-juggler: AB (weight) still applies if gx > 0
        if gx > 0:
            if gx >= 7000:
                weight_val = 1.0
            elif gx >= 5000:
                weight_val = 0.8
            elif gx >= 3000:
                weight_val = 0.6
            else:
                weight_val = 0.5
    
    ws.cell(curr_r, 18, payout_coeff)
    ws.cell(curr_r, 19, grape_prob)
    ws.cell(curr_r, 20, bb_ratio)
    ws.cell(curr_r, 21, rb_ratio)
    ws.cell(curr_r, 22, comb_ratio)
    ws.cell(curr_r, 23, bb_setting)
    ws.cell(curr_r, 24, rb_setting)
    ws.cell(curr_r, 25, comb_setting)
    ws.cell(curr_r, 26, grape_setting)
    ws.cell(curr_r, 27, avg_score)
    ws.cell(curr_r, 28, weight_val)
    ws.cell(curr_r, 29, final_score)
    
    # --- AD (30): 評価スコア ---
    ad_val = 0
    af_notes = []
    if gx > 0:
        base_score = 0
        if is_jug == 1:
            base_score = final_score if isinstance(final_score, (int, float)) else 0
        else:
            if gx >= 1000:
                if dx >= 3000:
                    base_score = 6.5
                elif dx >= 2000:
                    base_score = 5.5
                elif dx >= 1000:
                    base_score = 4.5
                elif dx >= 0:
                    base_score = 3.5
                else:
                    base_score = 2.0
                    
        ad_val = round(min(70.0, base_score * 10.0), 1)
        if ad_val > 0:
            af_notes.append(f"評価:{ad_val}pt")
        
        # Bonus: 据置+15
        if base_score >= 4.5:
            ad_val += 15
            af_notes.append("据置+15")
        # Bonus: 凹グ+10
        if base_score < 3.5:
            days_elapsed = ""
            if prev_rows:
                prev_date = prev_rows[-1]["date"]
                d1 = current_date if isinstance(current_date, datetime.datetime) else datetime.datetime.combine(current_date, datetime.time())
                d2 = prev_date if isinstance(prev_date, datetime.datetime) else datetime.datetime.combine(prev_date, datetime.time())
                days_elapsed = (d1 - d2).days
            slump_days = (days_elapsed + 1) if (dx < 0 and isinstance(days_elapsed, int)) else 0
            if slump_days >= 2:
                ad_val += 10
                af_notes.append("凹グ+10")
        if m_name_setting in ["準ピ", "準ピ2"]:
            ad_val += 5
            af_notes.append("準ピ+5")
        # Tail average
        this_tail = tail_num
        same_tail_rows = [r for r in processed_history if int(str(r["machine_number"])[-1]) == this_tail and r["date"] <= current_date]
        avg_diff_tail = sum(r["diff_coins"] for r in same_tail_rows) / len(same_tail_rows) if same_tail_rows else 0
        if avg_diff_tail > 0:
            ad_val += 10
            af_notes.append("末尾+10")
                
    ws.cell(curr_r, 30, ad_val)
    
    # --- AE (31): タイスコア ---
    ae_val = 0
    if ad_val > 0:
        ae_val = ad_val + (100000 - curr_r) / 100000000.0
    ws.cell(curr_r, 31, ae_val)
    
    # --- AF (32): メモ ---
    ws.cell(curr_r, 32, " | ".join(af_notes) if af_notes else "")
    
    # Append to history so subsequent rows can use this row
    processed_history.append(row_data)
