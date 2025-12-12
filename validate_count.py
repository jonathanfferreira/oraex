import pandas as pd

FILE_PATH = r"D:\antigravity\oraex\cmdb\ORAEX - Consolidação GetTech 2025 (1).xlsm"

print("="*70)
print("VALIDANDO CONTAGEM DE SERVIDORES (PRIMARY + STANDBY)")
print("="*70)

df = pd.read_excel(FILE_PATH, sheet_name='GetNet - Oracle Databases', engine='openpyxl')

print(f"\n📋 Colunas disponíveis:")
for i, col in enumerate(df.columns):
    print(f"  {i}: {col}")

# Filtrar apenas registros válidos
df_valid = df[df['PRIMARY HOSTNAME'].notna()].copy()

print(f"\n📊 ANÁLISE DE CONTAGEM:")
print("-"*70)

# Verificar coluna Total Servidores
if 'Total Servidores' in df.columns:
    total_col = df_valid['Total Servidores'].sum()
    print(f"  Soma da coluna 'Total Servidores': {int(total_col)}")
else:
    print("  Coluna 'Total Servidores' não encontrada, verificando alternativas...")

# Contar PRIMARY
primary_count = df_valid['PRIMARY HOSTNAME'].notna().sum()
print(f"  Contagem de PRIMARY HOSTNAME: {primary_count}")

# Contar STANDBY
if 'STANDBY HOSTNAME' in df.columns:
    standby_count = df_valid['STANDBY HOSTNAME'].notna().sum()
    print(f"  Contagem de STANDBY HOSTNAME: {standby_count}")
    print(f"  Total (PRIMARY + STANDBY): {primary_count + standby_count}")

# Verificar valores únicos na coluna Total Servidores
print(f"\n📈 Distribuição 'Total Servidores':")
print("-"*70)
total_srv = df_valid['Total Servidores'].value_counts().sort_index()
for val, count in total_srv.items():
    if pd.notna(val):
        print(f"  {int(val)} servidor(es): {count} linhas")

# Somar todo o parque
total_real = df_valid['Total Servidores'].sum()
print(f"\n🎯 TOTAL REAL DE SERVIDORES: {int(total_real)}")

# Por situação
print(f"\n📊 POR SITUAÇÃO:")
print("-"*70)

def get_situacao(val):
    if pd.isna(val): return 'Desconhecido'
    return 'Ativo' if 'Ativo' in str(val) else ('Descontinuado' if 'Descontinuado' in str(val) else str(val))

df_valid['Situacao'] = df_valid['SITUAÇÃO'].apply(get_situacao)

for sit in df_valid['Situacao'].unique():
    subset = df_valid[df_valid['Situacao'] == sit]
    count = subset['Total Servidores'].sum()
    print(f"  {sit}: {int(count)} servidores")
