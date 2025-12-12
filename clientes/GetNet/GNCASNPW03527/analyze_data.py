import pandas as pd
import json

xl = pd.ExcelFile('GNCASNPW03527_espaco_tabelas.xlsx')

# Resumo geral
total_tabelas = 0
total_registros = 0
db_stats = {}
all_tables = []

for sheet in xl.sheet_names:
    df = pd.read_excel(xl, sheet_name=sheet)
    db_name = sheet.replace('GNCASNPW03527.', '')
    total_tabelas += len(df)
    
    # Converter Total Rows para número
    df['Total Rows Num'] = pd.to_numeric(
        df['Total Rows'].astype(str).str.replace('.', '', regex=False).str.replace(',', '', regex=False), 
        errors='coerce'
    ).fillna(0).astype(int)
    
    total_rows = df['Total Rows Num'].sum()
    total_registros += total_rows
    
    # Add database name
    df['Database'] = db_name
    all_tables.append(df)
    
    db_stats[db_name] = {
        'tabelas': len(df),
        'registros': int(total_rows),
        'top_5': df.nlargest(5, 'Total Rows Num')[['Full Table Name', 'Total Rows', 'Total Reserved Size']].values.tolist()
    }

print('=== RESUMO DO SERVIDOR GNCASNPW03527 ===')
print(f'Total de Bancos de Dados: {len(xl.sheet_names)}')
print(f'Total de Tabelas: {total_tabelas}')
print(f'Total de Registros: {total_registros:,}')
print()

for db, stats in db_stats.items():
    print(f'\n--- {db} ---')
    print(f'  Tabelas: {stats["tabelas"]}')
    print(f'  Registros: {stats["registros"]:,}')
    print(f'  Top 5 tabelas:')
    for t in stats['top_5']:
        print(f'    - {t[0]}: {t[1]} rows ({t[2]})')

# Combine all tables
all_df = pd.concat(all_tables, ignore_index=True)

# Helper para limpar e converter tamanho
def clean_size(size_str):
    if not isinstance(size_str, str): return 0
    clean = size_str.upper().replace('.', '').replace(',', '.')
    if 'GB' in clean: return float(clean.replace(' GB', '').strip()) * 1024 * 1024
    if 'MB' in clean: return float(clean.replace(' MB', '').strip()) * 1024
    if 'KB' in clean: return float(clean.replace(' KB', '').strip())
    return 0

# Calcular colunas numéricas de tamanho (KB)
all_df['TotalReserved_KB'] = all_df['Total Reserved Size'].apply(clean_size)
all_df['Unused_KB'] = all_df['Unused'].apply(clean_size)
all_df['Data_KB'] = all_df['Data'].apply(clean_size)
all_df['Index_KB'] = all_df['Indexes'].apply(clean_size)

# === ANÁLISE DE SCHEMA ===
all_df['Schema'] = all_df['Full Table Name'].apply(lambda x: x.split('.')[0] if '.' in str(x) else 'dbo')
schema_stats = all_df.groupby('Schema')['TotalReserved_KB'].sum().sort_values(ascending=False)
top_schemas = [{'label': s, 'value': v/1024} for s, v in schema_stats.head(10).items()] # MB

# === ANÁLISE DE FRAGMENTAÇÃO ===
total_unused_mb = all_df['Unused_KB'].sum() / 1024
top_fragmented = all_df.nlargest(10, 'Unused_KB')[['Database', 'Full Table Name', 'Unused', 'Total Reserved Size']].to_dict('records')

# === ANÁLISE DE OFENSORES DE ÍNDICE ===
# Filtra tabelas com tamanho relevante (> 100MB) e Indice > 60%
index_offenders = all_df[
    (all_df['TotalReserved_KB'] > 100*1024) & 
    (all_df['% Index Weight'] > 0.6)
].nlargest(10, '% Index Weight')[['Database', 'Full Table Name', 'Total Reserved Size', 'Indexes', '% Index Weight']].to_dict('records')

# Top 20 maiores tabelas
print('\n\n=== TOP 20 MAIORES TABELAS ===')
top20 = all_df.nlargest(20, 'Total Rows Num')[['Database', 'Full Table Name', 'Total Rows', 'Avg Row Size', 'Total Reserved Size', 'Data', 'Indexes', '% Index Weight']]
print(top20.to_string(index=False))

# Estatísticas de armazenamento
print('\n\n=== DISTRIBUIÇÃO POR TAMANHO ===')
print(all_df[['Database', 'Total Reserved Size']].groupby('Database').apply(lambda x: x['Total Reserved Size'].value_counts().head()).to_string())

# Salvar dados processados para o relatório
output = {
    'servidor': 'GNCASNPW03527',
    'total_databases': len(xl.sheet_names),
    'total_tabelas': total_tabelas,
    'total_registros': int(total_registros), # Converter para int nativo
    'databases': db_stats,
    'top20_tables': top20.to_dict('records'),
    'advanced_metrics': {
        'total_unused_mb': total_unused_mb,
        'top_fragmented': top_fragmented,
        'schemas': top_schemas,
        'index_offenders': index_offenders
    }
}

with open('data_summary.json', 'w', encoding='utf-8') as f:
    json.dump(output, f, ensure_ascii=False, indent=2, default=str)

print('\n\nDados salvos em data_summary.json')
