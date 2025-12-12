import pandas as pd
import re

FILE_PATH = r"D:\antigravity\oraex\cmdb\ORAEX - Consolidação GetTech 2025 (1).xlsm"

MONTHLY_SHEETS = [
    'FEVEREIRO-25', 'MARÇO-25', 'ABRIL-25', 'MAIO-25', 'JUNHO-25',
    'JULHO-25', 'AGOSTO-25', 'SETEMBRO-25', 'OUTUBRO-25', 'NOVEMBRO-25', 'DEZEMBRO-25'
]

def load_all_gmuds():
    all_data = []
    for sheet in MONTHLY_SHEETS:
        try:
            df = pd.read_excel(FILE_PATH, sheet_name=sheet, engine='openpyxl')
            
            # Mapear colunas
            col_mapping = {}
            for col in df.columns:
                col_lower = str(col).lower().strip()
                if 'status' in col_lower and 'gmud' in col_lower:
                    col_mapping[col] = 'Status'
                elif col_lower == 'gmud':
                    col_mapping[col] = 'GMUD_ID'
                elif 'título' in col_lower or 'titulo' in col_lower:
                    col_mapping[col] = 'Titulo'
                elif 'entorno' in col_lower:
                    col_mapping[col] = 'Entorno'
                elif 'designado' in col_lower or 'responsável' in col_lower or 'responsavel' in col_lower:
                    col_mapping[col] = 'Responsavel'
            
            df = df.rename(columns=col_mapping)
            df['Mes'] = sheet.replace('-25', '')
            
            if 'GMUD_ID' in df.columns:
                df = df[df['GMUD_ID'].notna()]
                df = df[df['GMUD_ID'].astype(str).str.contains('CHG', case=False, na=False)]
            
            all_data.append(df)
            print(f"✓ {sheet}: {len(df)} registros")
        except Exception as e:
            print(f"✗ {sheet}: {e}")
    
    return pd.concat(all_data, ignore_index=True) if all_data else pd.DataFrame()

# Carregar dados
print("Carregando dados...")
df = load_all_gmuds()

print(f"\n{'='*60}")
print("ANÁLISE DE COLUNAS DISPONÍVEIS")
print(f"{'='*60}")
print(f"Colunas encontradas: {list(df.columns)}")

print(f"\n{'='*60}")
print("ANÁLISE DE TÍTULOS (primeiros 10)")
print(f"{'='*60}")
if 'Titulo' in df.columns:
    for i, titulo in enumerate(df['Titulo'].head(10)):
        print(f"{i+1}. {titulo}")

print(f"\n{'='*60}")
print("ANÁLISE DE ENTORNO")
print(f"{'='*60}")
if 'Entorno' in df.columns:
    print(df['Entorno'].value_counts())

print(f"\n{'='*60}")
print("ANÁLISE DE RESPONSÁVEIS (top 10)")
print(f"{'='*60}")
if 'Responsavel' in df.columns:
    print(df['Responsavel'].value_counts().head(10))
else:
    print("Coluna 'Responsavel' não encontrada. Colunas disponíveis:")
    for col in df.columns:
        if 'design' in col.lower() or 'respons' in col.lower() or 'exec' in col.lower():
            print(f"  - {col}")

print(f"\n{'='*60}")
print("AMOSTRA DE TÍTULOS COM 'PSU'")
print(f"{'='*60}")
if 'Titulo' in df.columns:
    psu_titles = df[df['Titulo'].str.contains('PSU', case=False, na=False)]['Titulo'].head(10)
    for t in psu_titles:
        print(f"  {t[:100]}...")
