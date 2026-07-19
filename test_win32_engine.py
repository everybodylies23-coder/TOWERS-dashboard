import win32com.client
import os

dir_path = r"C:\Users\user\Desktop\Antigravity\データ分析自動化"
temp_path = os.path.join(dir_path, "データ分析_テンプレート_v13.11.xlsx")
out_path = os.path.join(dir_path, "test_win32.xlsx")

print("Initializing Excel Engine...")
try:
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    
    print(f"Opening workbook: {temp_path}")
    wb = excel.Workbooks.Open(temp_path)
    
    print(f"Saving copy to: {out_path}")
    wb.SaveAs(out_path)
    
    wb.Close(SaveChanges=False)
    excel.Quit()
    print("Excel Engine run completed successfully!")
except Exception as e:
    print(f"Excel Engine run failed: {e}")
