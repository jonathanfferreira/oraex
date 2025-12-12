import pandas as pd
import re
from collections import Counter

FILE_PATH = r"D:\antigravity\oraex\cmdb\ORAEX - Consolidação GetTech 2025 (1).xlsm"

MONTHLY_SHEETS = [
    'FEVEREIRO-25', 'MARÇO-25', 'ABRIL-25', 'MAIO-25', 'JUNHO-25',
    'JULHO-25', 'AGOSTO-25', 'SETEMBRO-25', 'OUTUBRO-25', 'NOVEMBRO-25', 'DEZEMBRO-25'
]

def extract_hostnames(title):
    """Extract GNCAS... hostnames from GMUD title"""
    if pd.isna(title):
        return []
    # Pattern: GNCAS followed by alphanumeric characters
    pattern = r'(gncas[a-z0-9]+)'
    matches = re.findall(pattern, str(title).lower())
    return [m.upper() for m in matches]

def normalize_status(status):
    """Normalize status values"""
    if pd.isna(status):
        return 'DESCONHECIDO'
    status = str(status).strip().upper()
    if 'ENCERRADA' in status or 'FECHADA' in status or '✅' in status:
        return 'SUCESSO'
    elif 'CANCELADA' in status or '❌' in status:
        return 'CANCELADA'
    elif 'REPLANEJAR' in status or '🔄' in status:
        return 'REPLANEJADA'
    elif 'ANDAMENTO' in status or 'EXECUÇÃO' in status:
        return 'EM ANDAMENTO'
    else:
        return status

def load_all_gmuds():
    """Load and consolidate all GMUD data from monthly sheets"""
    all_data = []
    
    for sheet in MONTHLY_SHEETS:
        try:
            df = pd.read_excel(FILE_PATH, sheet_name=sheet, engine='openpyxl')
            
            # Standardize column names (handle slight variations)
            col_mapping = {}
            for col in df.columns:
                col_lower = str(col).lower().strip()
                if 'status' in col_lower and 'gmud' in col_lower:
                    col_mapping[col] = 'Status'
                elif col_lower == 'gmud':
                    col_mapping[col] = 'GMUD_ID'
                elif 'título' in col_lower or 'titulo' in col_lower:
                    col_mapping[col] = 'Titulo'
                elif 'data' in col_lower and 'início' in col_lower:
                    col_mapping[col] = 'Data_Inicio'
                elif 'entorno' in col_lower:
                    col_mapping[col] = 'Entorno'
                elif 'cliente' in col_lower:
                    col_mapping[col] = 'Cliente'
                elif 'designado' in col_lower:
                    col_mapping[col] = 'Responsavel'
                elif 'tipo' in col_lower and 'banco' in col_lower:
                    col_mapping[col] = 'Tipo_BD'
            
            df = df.rename(columns=col_mapping)
            df['Mes_Origem'] = sheet.replace('-25', '')
            
            # Filter only rows that have a GMUD ID (CHG...)
            if 'GMUD_ID' in df.columns:
                df = df[df['GMUD_ID'].notna()]
                df = df[df['GMUD_ID'].astype(str).str.contains('CHG', case=False, na=False)]
            
            all_data.append(df)
            print(f"✓ {sheet}: {len(df)} GMUDs encontradas")
            
        except Exception as e:
            print(f"✗ {sheet}: Erro - {e}")
    
    if all_data:
        consolidated = pd.concat(all_data, ignore_index=True)
        return consolidated
    return pd.DataFrame()

def generate_metrics(df):
    """Generate summary metrics"""
    print("\n" + "="*60)
    print("MÉTRICAS CONSOLIDADAS - PSU ORACLE 2025")
    print("="*60)
    
    # Normalize status
    df['Status_Normalizado'] = df['Status'].apply(normalize_status)
    
    # Extract hostnames
    df['Hostnames'] = df['Titulo'].apply(extract_hostnames)
    df['Num_Servidores'] = df['Hostnames'].apply(len)
    
    # Basic counts
    total_gmuds = len(df)
    print(f"\n📊 TOTAL DE GMUDs: {total_gmuds}")
    
    # Status breakdown
    print("\n📈 DISTRIBUIÇÃO POR STATUS:")
    status_counts = df['Status_Normalizado'].value_counts()
    for status, count in status_counts.items():
        pct = count / total_gmuds * 100
        print(f"   {status}: {count} ({pct:.1f}%)")
    
    # Unique servers
    all_hostnames = []
    for hosts in df['Hostnames']:
        all_hostnames.extend(hosts)
    unique_servers = len(set(all_hostnames))
    total_server_updates = len(all_hostnames)
    
    print(f"\n🖥️ SERVIDORES:")
    print(f"   Servidores únicos atualizados: {unique_servers}")
    print(f"   Total de atualizações (com repetições): {total_server_updates}")
    
    # Monthly breakdown
    print("\n📅 DISTRIBUIÇÃO MENSAL:")
    monthly = df.groupby('Mes_Origem').size()
    for mes, count in monthly.items():
        print(f"   {mes}: {count} GMUDs")
    
    # Success rate
    success_count = len(df[df['Status_Normalizado'] == 'SUCESSO'])
    success_rate = success_count / total_gmuds * 100 if total_gmuds > 0 else 0
    print(f"\n✅ TAXA DE SUCESSO: {success_rate:.1f}%")
    
    return df, {
        'total_gmuds': total_gmuds,
        'status_counts': status_counts.to_dict(),
        'unique_servers': unique_servers,
        'total_server_updates': total_server_updates,
        'success_rate': success_rate,
        'monthly_counts': monthly.to_dict()
    }

if __name__ == "__main__":
    print("Carregando dados...")
    df = load_all_gmuds()
    
    if not df.empty:
        df_enriched, metrics = generate_metrics(df)
        
        # Save consolidated data
        output_path = r"D:\antigravity\oraex\cmdb\consolidated_gmuds_2025.xlsx"
        df_enriched.to_excel(output_path, index=False)
        print(f"\n💾 Dados consolidados salvos em: {output_path}")
    else:
        print("Nenhum dado encontrado!")
