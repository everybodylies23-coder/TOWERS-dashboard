import openpyxl

path = r"C:\Users\user\Desktop\Antigravity\データ分析自動化\スタジアム_データ.xlsx"
wb = openpyxl.load_workbook(path, data_only=True)
ws = wb['【AI】予想・答え合わせ']

print("=== Predictions Headers (Row 3) ===")
for c in range(1, 10):
    print(f"Col {c} ({chr(64+c)}): {ws.cell(3, c).value}")
    
wb.close()
