import pandas as pd

file_path = r"D:\antigravity\oraex\cmdb\ORAEX - Consolidação GetTech 2025 (1).xlsm"

print(f"Inspecting file: {file_path}")
print("="*60)

try:
    xl = pd.ExcelFile(file_path, engine='openpyxl')
    print(f"Total Sheets: {len(xl.sheet_names)}")
    print(f"Sheet Names: {xl.sheet_names}")
    
    for sheet in xl.sheet_names:
        print(f"\n{'='*60}")
        print(f"SHEET: {sheet}")
        print(f"{'='*60}")
        
        try:
            df = pd.read_excel(file_path, sheet_name=sheet, nrows=5, engine='openpyxl')
            print(f"Columns ({len(df.columns)}):")
            for i, col in enumerate(df.columns):
                print(f"  {i+1}. {col}")
            
            print(f"\nFirst 2 rows sample:")
            if not df.empty and len(df) >= 1:
                print(df.head(2).to_string())
            else:
                print("  (empty or insufficient data)")
        except Exception as e_sheet:
            print(f"  Error reading sheet: {e_sheet}")

except Exception as e:
    print(f"CRITICAL ERROR: {e}")
