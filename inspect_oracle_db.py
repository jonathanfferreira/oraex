import pandas as pd

FILE_PATH = r"D:\antigravity\oraex\cmdb\ORAEX - Consolidação GetTech 2025 (1).xlsm"

print("="*70)
print("ANALISANDO ABA: GetNet - Oracle Databases")
print("="*70)

# Tentar ler a aba
try:
    df = pd.read_excel(FILE_PATH, sheet_name='GetNet - Oracle Databases', engine='openpyxl')
    
    print(f"\n📊 DIMENSÕES: {df.shape[0]} linhas x {df.shape[1]} colunas")
    
    print("\n📋 COLUNAS ENCONTRADAS:")
    print("-"*70)
    for i, col in enumerate(df.columns):
        print(f"  {i+1}. {col}")
    
    print("\n📝 PRIMEIRAS 5 LINHAS:")
    print("-"*70)
    print(df.head().to_string())
    
    print("\n📈 ESTATÍSTICAS POR COLUNA:")
    print("-"*70)
    for col in df.columns:
        if df[col].dtype == 'object':
            unique = df[col].nunique()
            top_vals = df[col].value_counts().head(5).to_dict()
            print(f"\n  {col}:")
            print(f"    → Valores únicos: {unique}")
            print(f"    → Top 5: {top_vals}")
        else:
            print(f"\n  {col}:")
            print(f"    → Min: {df[col].min()}, Max: {df[col].max()}, Mean: {df[col].mean():.2f}")
            
except Exception as e:
    print(f"Erro: {e}")
    
    # Listar todas as abas disponíveis
    print("\n📂 Listando TODAS as abas da planilha:")
    xl = pd.ExcelFile(FILE_PATH, engine='openpyxl')
    for i, sheet in enumerate(xl.sheet_names):
        print(f"  {i+1}. {sheet}")
