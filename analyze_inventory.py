import pandas as pd

FILE_PATH = r"D:\antigravity\oraex\cmdb\ORAEX - Consolidação GetTech 2025 (1).xlsm"

print("="*70)
print("ANÁLISE COMPLETA: Inventário Oracle GetNet")
print("="*70)

df = pd.read_excel(FILE_PATH, sheet_name='GetNet - Oracle Databases', engine='openpyxl')

# Limpar dados
df = df[df['PRIMARY HOSTNAME'].notna()]

# Normalizar situação
def get_situacao(val):
    if pd.isna(val): return 'Desconhecido'
    val = str(val).strip()
    if 'Ativo' in val: return 'Ativo'
    if 'Descontinuado' in val: return 'Descontinuado'
    return val

df['Situacao'] = df['SITUAÇÃO'].apply(get_situacao)

# Normalizar entorno
def get_entorno(val):
    if pd.isna(val): return 'Desconhecido'
    val = str(val).strip()
    if 'Prod' in val: return 'Produção'
    if 'Homolog' in val: return 'Homologação'
    if 'Desenv' in val: return 'Desenvolvimento'
    if 'Trans' in val: return 'Transacional'
    if 'Descontinuado' in val: return 'Descontinuado'
    return val

df['Entorno'] = df['ENVIROMENT'].apply(get_entorno)

# Extrair versão PSU
def get_psu_version(val):
    if pd.isna(val): return None
    val = str(val).strip()
    if 'Descontinuado' in val: return 'Descontinuado'
    if '19.' in val:
        # Extrair número
        import re
        match = re.search(r'19\.(\d+)', val)
        if match:
            return f"19.{match.group(1)}"
    return val

df['PSU_Version'] = df['GRID/PSU VERSION'].apply(get_psu_version)

# Versão mais recente (19.29 é a mais atual)
LATEST_PSU = '19.29'
QUARTERS_2025 = ['19.25', '19.26', '19.27', '19.28', '19.29']

def is_outdated(version):
    if version is None or version == 'Descontinuado':
        return False
    if version not in QUARTERS_2025:
        return True  # Versão de 2024 ou anterior
    return False

def get_quarters_behind(version):
    if version is None or version == 'Descontinuado':
        return None
    try:
        current_idx = QUARTERS_2025.index(LATEST_PSU)
        if version in QUARTERS_2025:
            ver_idx = QUARTERS_2025.index(version)
            return current_idx - ver_idx
        else:
            # Versão antiga (2024)
            return 5  # Mais de 4 quarters atrás
    except:
        return None

df['Is_Outdated'] = df['PSU_Version'].apply(is_outdated)
df['Quarters_Behind'] = df['PSU_Version'].apply(get_quarters_behind)

# Filtrar apenas ativos
df_ativos = df[df['Situacao'] == 'Ativo'].copy()

print(f"\n📊 RESUMO GERAL:")
print("-"*70)
print(f"  Total de registros: {len(df)}")
print(f"  Servidores Ativos: {len(df_ativos)}")
print(f"  Servidores Descontinuados: {len(df[df['Situacao'] == 'Descontinuado'])}")

print(f"\n🌐 DISTRIBUIÇÃO POR ENTORNO (Ativos):")
print("-"*70)
entorno_counts = df_ativos['Entorno'].value_counts()
for ent, count in entorno_counts.items():
    print(f"  {ent}: {count}")

print(f"\n📦 DISTRIBUIÇÃO POR VERSÃO PSU (Ativos):")
print("-"*70)
psu_counts = df_ativos['PSU_Version'].value_counts().sort_index()
for psu, count in psu_counts.items():
    behind = get_quarters_behind(psu)
    behind_str = f" ({behind} quarter(s) atrás)" if behind and behind > 0 else " ✅ ATUAL"
    print(f"  {psu}: {count} servidores{behind_str}")

print(f"\n⚠️ ANÁLISE DE DESATUALIZAÇÃO (Ativos):")
print("-"*70)
outdated = df_ativos[df_ativos['Is_Outdated'] == True]
print(f"  Servidores com PSU anterior a 2025: {len(outdated)}")

# Por quarters atrás
for q in range(5, 0, -1):
    count = len(df_ativos[df_ativos['Quarters_Behind'] == q])
    if count > 0:
        print(f"  {q} quarter(s) atrás: {count} servidores")

print(f"\n🎯 SERVIDORES NA VERSÃO MAIS RECENTE (19.29):")
print("-"*70)
latest = df_ativos[df_ativos['PSU_Version'] == LATEST_PSU]
print(f"  Total: {len(latest)} servidores ({len(latest)/len(df_ativos)*100:.1f}%)")

print(f"\n📈 VERSÃO DB (Oracle):")
print("-"*70)
db_counts = df_ativos['DB VERSION'].value_counts()
for db, count in db_counts.items():
    print(f"  {db}: {count}")

print(f"\n🔍 SERVIDORES DESATUALIZADOS (versão < 19.25):")
print("-"*70)
very_old = df_ativos[(df_ativos['Quarters_Behind'] >= 5) | (df_ativos['Is_Outdated'] == True)]
if len(very_old) > 0:
    print(f"  Total: {len(very_old)} servidores")
    for _, row in very_old.head(10).iterrows():
        hostname = row['PRIMARY HOSTNAME']
        psu = row['PSU_Version']
        entorno = row['Entorno']
        print(f"    → {hostname} | PSU: {psu} | {entorno}")
else:
    print("  Nenhum servidor com versão muito antiga!")
