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
            col_mapping = {}
            for col in df.columns:
                col_lower = str(col).lower().strip()
                if 'status' in col_lower and 'gmud' in col_lower:
                    col_mapping[col] = 'Status'
                elif col_lower == 'gmud':
                    col_mapping[col] = 'GMUD_ID'
                elif 'título' in col_lower or 'titulo' in col_lower:
                    col_mapping[col] = 'Titulo'
            
            df = df.rename(columns=col_mapping)
            df['Mes'] = sheet.replace('-25', '')
            
            if 'GMUD_ID' in df.columns:
                df = df[df['GMUD_ID'].notna()]
                df = df[df['GMUD_ID'].astype(str).str.contains('CHG', case=False, na=False)]
            
            all_data.append(df)
        except Exception as e:
            print(f"Erro em {sheet}: {e}")
    
    return pd.concat(all_data, ignore_index=True) if all_data else pd.DataFrame()

def categorize_gmud(titulo):
    """Categoriza GMUD pelo tipo de atividade"""
    if pd.isna(titulo):
        return 'Desconhecido'
    titulo = str(titulo).upper()
    
    if 'PSU' in titulo:
        return 'PSU Oracle'
    elif 'ODBC' in titulo:
        return 'Drivers ODBC'
    elif 'DATAGUARD' in titulo:
        return 'Dataguard'
    elif 'MONGO' in titulo:
        return 'MongoDB'
    elif 'REDIS' in titulo:
        return 'Redis'
    elif 'SQLSERVER' in titulo or 'SQL SERVER' in titulo:
        return 'SQL Server'
    elif 'POSTGRESQL' in titulo or 'POSTGRES' in titulo:
        return 'PostgreSQL'
    elif 'MYSQL' in titulo:
        return 'MySQL'
    elif 'JAVA' in titulo:
        return 'Java'
    elif 'RU ' in titulo or ' RU' in titulo:
        return 'Release Update (RU)'
    elif 'SINCRONIZAÇÃO' in titulo or 'RECONSTRUÇÃO' in titulo:
        return 'Sincronização/Reconstrução'
    elif 'VULNERABILIDADE' in titulo or 'SECURITY' in titulo:
        return 'Correção de Vulnerabilidade'
    elif 'MANUTENÇÃO' in titulo:
        return 'Manutenção Geral'
    else:
        return 'Outros'

# Carregar dados
print("Carregando dados...")
df = load_all_gmuds()

# Categorizar
df['Categoria'] = df['Titulo'].apply(categorize_gmud)

# Normalizar status
def normalize_status(status):
    if pd.isna(status):
        return 'DESCONHECIDO'
    status = str(status).strip().upper()
    if 'ENCERRADA' in status or 'FECHADA' in status or '✅' in status:
        return 'SUCESSO'
    elif 'CANCELADA' in status or '❌' in status or 'CANCELAR' in status:
        return 'CANCELADA'
    elif 'REPLANEJAR' in status or '🔄' in status or 'REAGENDADA' in status:
        return 'REPLANEJADA'
    elif 'INSUCESSO' in status:
        return 'INSUCESSO'
    else:
        return 'OUTROS'

df['Status_Norm'] = df['Status'].apply(normalize_status)

print("\n" + "="*70)
print("ANÁLISE COMPLETA DE TIPOS DE GMUDs")
print("="*70)

# Estatísticas por categoria
print("\n📊 DISTRIBUIÇÃO POR TIPO DE ATIVIDADE:")
print("-"*70)
cat_stats = df.groupby('Categoria').agg({
    'GMUD_ID': 'count',
    'Status_Norm': lambda x: (x == 'SUCESSO').sum()
}).reset_index()
cat_stats.columns = ['Categoria', 'Total', 'Sucesso']
cat_stats['Taxa'] = (cat_stats['Sucesso'] / cat_stats['Total'] * 100).round(1)
cat_stats = cat_stats.sort_values('Total', ascending=False)

for _, row in cat_stats.iterrows():
    print(f"  {row['Categoria']:30} | Total: {row['Total']:4} | Sucesso: {row['Sucesso']:4} | Taxa: {row['Taxa']:.1f}%")

print("\n" + "="*70)
print("EXEMPLOS DE TÍTULOS POR CATEGORIA")
print("="*70)

for cat in cat_stats['Categoria'].head(10):
    exemplos = df[df['Categoria'] == cat]['Titulo'].head(2).tolist()
    print(f"\n🔹 {cat}:")
    for ex in exemplos:
        print(f"   → {str(ex)[:90]}...")

print("\n" + "="*70)
print("VERIFICAÇÃO DE CONTAGEM")
print("="*70)
print(f"Total GERAL de GMUDs: {len(df)}")
print(f"Total com SUCESSO (geral): {len(df[df['Status_Norm'] == 'SUCESSO'])}")
print(f"Total PSU com SUCESSO: {len(df[(df['Categoria'] == 'PSU Oracle') & (df['Status_Norm'] == 'SUCESSO')])}")
