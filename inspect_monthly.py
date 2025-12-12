import pandas as pd

file_path = r"D:\antigravity\oraex\cmdb\ORAEX - Consolidação GetTech 2025 (1).xlsm"

# Focus on monthly sheets that contain GMUD data
monthly_sheets = ['FEVEREIRO-25', 'MARÇO-25', 'ABRIL-25', 'MAIO-25', 'JUNHO-25', 
                  'JULHO-25', 'AGOSTO-25', 'SETEMBRO-25', 'OUTUBRO-25', 'NOVEMBRO-25', 'DEZEMBRO-25']

for sheet in monthly_sheets:
    print(f"\n{'='*80}")
    print(f"SHEET: {sheet}")
    print(f"{'='*80}")
    
    try:
        df = pd.read_excel(file_path, sheet_name=sheet, nrows=5, engine='openpyxl')
        print(f"Columns ({len(df.columns)}):")
        for i, col in enumerate(df.columns):
            print(f"  {i+1}. '{col}'")
        
        print(f"\nSample Data (first 2 rows):")
        pd.set_option('display.max_columns', None)
        pd.set_option('display.width', None)
        if not df.empty:
            print(df.head(2).T)  # Transpose for readability
    except Exception as e:
        print(f"  Error: {e}")

# Also check the RELATÓRIO sheet as it might have summary data
print(f"\n{'='*80}")
print(f"SHEET: RELATÓRIO")
print(f"{'='*80}")
try:
    df = pd.read_excel(file_path, sheet_name='RELATÓRIO', nrows=10, engine='openpyxl')
    print(f"Columns ({len(df.columns)}):")
    for i, col in enumerate(df.columns):
        print(f"  {i+1}. '{col}'")
    print(f"\nSample:")
    print(df.head(5).T)
except Exception as e:
    print(f"  Error: {e}")
