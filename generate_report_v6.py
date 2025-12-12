import pandas as pd
import plotly.graph_objects as go
from datetime import datetime
import re
import base64

FILE_PATH = r"D:\antigravity\oraex\cmdb\ORAEX - Consolidação GetTech 2025 (1).xlsm"
LOGO_PATH = r"D:\antigravity\oraex\cmdb\oraex_logo.png"
OUTPUT_HTML = r"D:\antigravity\oraex\cmdb\relatorio_psu_2025_v6_completo.html"

MONTHLY_SHEETS = [
    'FEVEREIRO-25', 'MARÇO-25', 'ABRIL-25', 'MAIO-25', 'JUNHO-25',
    'JULHO-25', 'AGOSTO-25', 'SETEMBRO-25', 'OUTUBRO-25', 'NOVEMBRO-25', 'DEZEMBRO-25'
]
MONTH_ORDER = ['FEVEREIRO', 'MARÇO', 'ABRIL', 'MAIO', 'JUNHO', 'JULHO', 'AGOSTO', 'SETEMBRO', 'OUTUBRO', 'NOVEMBRO', 'DEZEMBRO']
LATEST_PSU = '19.29'
QUARTERS_2025 = ['19.25', '19.26', '19.27', '19.28', '19.29']

def get_logo_base64():
    try:
        with open(LOGO_PATH, "rb") as f:
            return base64.b64encode(f.read()).decode('utf-8')
    except:
        return None

def extract_hostnames(title):
    if pd.isna(title): return []
    return [m.upper() for m in re.findall(r'(gncas[a-z0-9]+)', str(title).lower())]

def extract_psu_version(title):
    if pd.isna(title): return None
    match = re.search(r'PSU\s*(19[.\d]+)', str(title), re.IGNORECASE)
    if match: return match.group(1)
    match2 = re.search(r'19\.(\d+)', str(title))
    if match2: return f"19.{match2.group(1)}"
    return None

def normalize_status(status):
    if pd.isna(status): return 'DESCONHECIDO'
    status = str(status).strip().upper()
    if 'ENCERRADA' in status or 'FECHADA' in status or '✅' in status: return 'SUCESSO'
    elif 'CANCELADA' in status or '❌' in status or 'CANCELAR' in status: return 'CANCELADA'
    elif 'REPLANEJAR' in status or '🔄' in status or 'REAGENDADA' in status: return 'REPLANEJADA'
    elif 'INSUCESSO' in status: return 'INSUCESSO'
    return 'OUTROS'

def normalize_responsavel(resp):
    if pd.isna(resp): return 'Não Atribuído'
    resp = str(resp).strip().title()
    names = {'Guilherme': 'Guilherme Fonseca', 'Bruno': 'Bruno Ferreira', 'Alcides': 'Alcides Souto',
             'Kaue': 'Kaue Santos', 'Rafael': 'Rafael Rabello', 'Luca': 'Luca Mozart', 'Jonathan': 'Jonathan Ferreira'}
    for key, val in names.items():
        if key in resp: return val
    return resp

def load_gmuds():
    all_data = []
    for sheet in MONTHLY_SHEETS:
        try:
            df = pd.read_excel(FILE_PATH, sheet_name=sheet, engine='openpyxl')
            col_mapping = {}
            for col in df.columns:
                col_lower = str(col).lower().strip()
                if 'status' in col_lower and 'gmud' in col_lower: col_mapping[col] = 'Status'
                elif col_lower == 'gmud': col_mapping[col] = 'GMUD_ID'
                elif 'título' in col_lower or 'titulo' in col_lower: col_mapping[col] = 'Titulo'
                elif 'entorno' in col_lower: col_mapping[col] = 'Entorno'
                elif 'designado' in col_lower or 'responsável' in col_lower or 'responsavel' in col_lower: col_mapping[col] = 'Responsavel'
            df = df.rename(columns=col_mapping)
            df['Mes'] = sheet.replace('-25', '')
            if 'GMUD_ID' in df.columns:
                df = df[df['GMUD_ID'].notna()]
                df = df[df['GMUD_ID'].astype(str).str.contains('CHG', case=False, na=False)]
            all_data.append(df)
        except: pass
    return pd.concat(all_data, ignore_index=True) if all_data else pd.DataFrame()

def load_inventory():
    df = pd.read_excel(FILE_PATH, sheet_name='GetNet - Oracle Databases', engine='openpyxl')
    df = df[df['PRIMARY HOSTNAME'].notna()]
    
    def get_situacao(val):
        if pd.isna(val): return 'Desconhecido'
        return 'Ativo' if 'Ativo' in str(val) else ('Descontinuado' if 'Descontinuado' in str(val) else str(val))
    
    def get_entorno(val):
        if pd.isna(val): return 'Outros'
        val = str(val).strip()
        if 'Prod' in val: return 'Produção'
        if 'Homolog' in val: return 'Homologação'
        if 'Desenv' in val: return 'Desenvolvimento'
        if 'Trans' in val: return 'Transacional'
        return 'Outros'
    
    def get_psu(val):
        if pd.isna(val): return None
        val = str(val).strip()
        if 'Descontinuado' in val: return 'Descontinuado'
        match = re.search(r'19\.(\d+)', val)
        return f"19.{match.group(1)}" if match else val
    
    def quarters_behind(version):
        if version is None or version == 'Descontinuado': return None
        try:
            if version in QUARTERS_2025:
                return QUARTERS_2025.index(LATEST_PSU) - QUARTERS_2025.index(version)
            return 5
        except: return None
    
    df['Situacao'] = df['SITUAÇÃO'].apply(get_situacao)
    df['Entorno'] = df['ENVIROMENT'].apply(get_entorno)
    df['PSU_Version'] = df['GRID/PSU VERSION'].apply(get_psu)
    df['Quarters_Behind'] = df['PSU_Version'].apply(quarters_behind)
    df['Hostname'] = df['PRIMARY HOSTNAME'].apply(lambda x: re.sub(r'[^\w]', '', str(x).split()[0]) if pd.notna(x) else '')
    
    return df

def generate_complete_report():
    print("Carregando dados...")
    df_gmuds = load_gmuds()
    df_inv = load_inventory()
    logo_b64 = get_logo_base64()
    
    # ========== MÉTRICAS GMUDs ==========
    df_gmuds['Status_Final'] = df_gmuds['Status'].apply(normalize_status)
    df_gmuds['Hostnames'] = df_gmuds['Titulo'].apply(extract_hostnames)
    df_gmuds['Num_Servers'] = df_gmuds['Hostnames'].apply(len)
    df_gmuds['Versao_PSU'] = df_gmuds['Titulo'].apply(extract_psu_version)
    df_gmuds['Responsavel_Norm'] = df_gmuds['Responsavel'].apply(normalize_responsavel)
    df_gmuds['Is_PSU'] = df_gmuds['Titulo'].str.contains('PSU', case=False, na=False)
    
    df_psu = df_gmuds[df_gmuds['Is_PSU']].copy()
    
    total_psu = len(df_psu)
    sucesso_psu = len(df_psu[df_psu['Status_Final'] == 'SUCESSO'])
    canceladas_psu = len(df_psu[df_psu['Status_Final'] == 'CANCELADA'])
    replanejadas_psu = len(df_psu[df_psu['Status_Final'] == 'REPLANEJADA'])
    outras_psu = total_psu - sucesso_psu - canceladas_psu - replanejadas_psu
    all_hosts = []
    for h in df_psu[df_psu['Status_Final'] == 'SUCESSO']['Hostnames']:
        all_hosts.extend(h)
    unique_servers = len(set(all_hosts))
    total_updates = len(all_hosts)
    horas_totais = total_updates * 3
    taxa_sucesso = sucesso_psu / total_psu * 100 if total_psu > 0 else 0
    
    # Entornos GMUDs
    entorno_prod_gmud = len(df_psu[df_psu['Entorno'] == 'P'])
    entorno_hml_gmud = len(df_psu[df_psu['Entorno'] == 'H'])
    entorno_dev_gmud = len(df_psu[df_psu['Entorno'] == 'D'])
    
    # Versões GMUDs
    versao_stats = df_psu.groupby('Versao_PSU').agg({'GMUD_ID': 'count', 'Status_Final': lambda x: (x == 'SUCESSO').sum()}).reset_index()
    versao_stats.columns = ['Versao', 'Total', 'Sucesso']
    versao_stats = versao_stats.dropna(subset=['Versao']).sort_values('Versao')
    
    # Executores
    executor_stats = df_psu.groupby('Responsavel_Norm').agg({'GMUD_ID': 'count', 'Status_Final': lambda x: (x == 'SUCESSO').sum(), 'Num_Servers': 'sum'}).reset_index()
    executor_stats.columns = ['Executor', 'Total_GMUDs', 'Sucesso', 'Servidores']
    executor_stats['Taxa'] = (executor_stats['Sucesso'] / executor_stats['Total_GMUDs'] * 100).round(1)
    executor_stats['Horas'] = executor_stats['Servidores'] * 3
    executor_stats = executor_stats.sort_values('Total_GMUDs', ascending=False)
    executor_stats = executor_stats[executor_stats['Executor'] != 'Não Atribuído'].head(8)
    
    # Mensal
    monthly_psu = df_psu.groupby(['Mes', 'Status_Final']).size().unstack(fill_value=0)
    monthly_psu = monthly_psu.reindex(MONTH_ORDER)
    
    # ========== MÉTRICAS INVENTÁRIO ==========
    # Usar coluna "Total Servidores" para contar PRIMARY + STANDBY corretamente
    df_ativos = df_inv[df_inv['Situacao'] == 'Ativo'].copy()
    
    # Total real usando a coluna "Total Servidores" (inclui standby)
    total_servidores = int(df_ativos['Total Servidores'].sum())
    
    # Para as versões, multiplicamos pela quantidade de servidores em cada linha
    srv_atualizados = int(df_ativos[df_ativos['PSU_Version'] == LATEST_PSU]['Total Servidores'].sum())
    srv_1q_atras = int(df_ativos[df_ativos['Quarters_Behind'] == 1]['Total Servidores'].sum())
    srv_2q_atras = int(df_ativos[df_ativos['Quarters_Behind'] == 2]['Total Servidores'].sum())
    srv_3q_mais = int(df_ativos[df_ativos['Quarters_Behind'] >= 3]['Total Servidores'].sum())
    pct_atualizados = srv_atualizados / total_servidores * 100 if total_servidores > 0 else 0
    
    # Entornos inventário (usando Total Servidores)
    inv_prod = int(df_ativos[df_ativos['Entorno'] == 'Produção']['Total Servidores'].sum())
    inv_hml = int(df_ativos[df_ativos['Entorno'] == 'Homologação']['Total Servidores'].sum())
    inv_dev = int(df_ativos[df_ativos['Entorno'] == 'Desenvolvimento']['Total Servidores'].sum())
    inv_trans = int(df_ativos[df_ativos['Entorno'] == 'Transacional']['Total Servidores'].sum())
    
    # Versões inventário (usando Total Servidores para contar corretamente)
    inv_versao = df_ativos.groupby('PSU_Version')['Total Servidores'].sum().sort_index()
    
    # Servidores críticos
    criticos = df_ativos[df_ativos['Quarters_Behind'] >= 4][['Hostname', 'PSU_Version', 'Entorno']].head(10)
    
    # ===== GRÁFICOS =====
    oraex_blue = '#0000FF'
    
    # Gráfico mensal - COM NÚMEROS VISÍVEIS
    fig_monthly = go.Figure()
    colors = {'SUCESSO': oraex_blue, 'REPLANEJADA': '#93C5FD', 'CANCELADA': '#D1D5DB'}
    for status in ['SUCESSO', 'REPLANEJADA', 'CANCELADA']:
        if status in monthly_psu.columns:
            values = monthly_psu[status]
            fig_monthly.add_trace(go.Bar(
                name=status, 
                x=monthly_psu.index, 
                y=values, 
                marker_color=colors.get(status),
                text=values,
                textposition='inside',
                textfont=dict(size=10, color='white' if status == 'SUCESSO' else '#374151')
            ))
    fig_monthly.update_layout(barmode='stack', paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
        font=dict(color='#374151', family='Inter'), height=380,
        xaxis=dict(showgrid=False), yaxis=dict(showgrid=True, gridcolor='rgba(0,0,0,0.06)'),
        legend=dict(orientation="h", y=1.12, x=0.5, xanchor='center'), margin=dict(t=50, b=50), bargap=0.3)
    
    # Gráfico executores - COM NÚMEROS VISÍVEIS
    fig_exec = go.Figure()
    exec_sucesso = executor_stats['Sucesso'].head(6)
    exec_outras = executor_stats['Total_GMUDs'].head(6) - executor_stats['Sucesso'].head(6)
    fig_exec.add_trace(go.Bar(
        y=executor_stats['Executor'].head(6), 
        x=exec_sucesso, 
        name='Concluídas', 
        orientation='h', 
        marker_color=oraex_blue,
        text=exec_sucesso.astype(int),
        textposition='inside',
        textfont=dict(size=11, color='white')
    ))
    fig_exec.add_trace(go.Bar(
        y=executor_stats['Executor'].head(6), 
        x=exec_outras, 
        name='Outras', 
        orientation='h', 
        marker_color='#E5E7EB',
        text=exec_outras.astype(int),
        textposition='inside',
        textfont=dict(size=11, color='#374151')
    ))
    fig_exec.update_layout(barmode='stack', paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
        font=dict(color='#374151', family='Inter'), height=300,
        xaxis=dict(showgrid=True, gridcolor='rgba(0,0,0,0.06)'), yaxis=dict(showgrid=False),
        legend=dict(orientation="h", y=1.08, x=0.5, xanchor='center'), margin=dict(l=130, t=40, b=30), bargap=0.25)
    
    # Gráfico inventário por versão
    fig_inv_versao = go.Figure()
    versions = []
    counts = []
    colors_inv = []
    for v in sorted(inv_versao.index):
        if v and v != 'Descontinuado':
            versions.append(v)
            counts.append(inv_versao[v])
            if v == LATEST_PSU:
                colors_inv.append('#10B981')
            elif v == '19.28':
                colors_inv.append('#60A5FA')
            else:
                colors_inv.append('#F59E0B')
    
    fig_inv_versao.add_trace(go.Bar(x=versions, y=counts, marker_color=colors_inv, text=counts, textposition='outside'))
    fig_inv_versao.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
        font=dict(color='#374151', family='Inter'), height=300,
        xaxis=dict(showgrid=False, title='Versão PSU'), yaxis=dict(showgrid=True, gridcolor='rgba(0,0,0,0.06)', title='Servidores'),
        margin=dict(t=30, b=50), bargap=0.4)
    
    # Gráfico inventário por entorno (donut)
    fig_inv_entorno = go.Figure(data=[go.Pie(
        labels=['Produção', 'Homologação', 'Desenvolvimento', 'Transacional'],
        values=[inv_prod, inv_hml, inv_dev, inv_trans],
        hole=0.6,
        marker_colors=[oraex_blue, '#60A5FA', '#93C5FD', '#DBEAFE'],
        textinfo='label+value',
        textposition='outside'
    )])
    fig_inv_entorno.update_layout(paper_bgcolor='rgba(0,0,0,0)', font=dict(color='#374151', family='Inter'),
        height=300, margin=dict(t=20, b=20, l=20, r=20), showlegend=False)
    
    logo_html = f'<img src="data:image/png;base64,{logo_b64}" alt="ORAEX" class="logo-img">' if logo_b64 else '<span class="logo-text">ORAEX</span>'
    
    # ===== HTML =====
    html = f"""
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Relatório Oracle 2025 | ORAEX</title>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap" rel="stylesheet">
    <script src="https://cdn.plot.ly/plotly-2.27.0.min.js"></script>
    <style>
        :root {{
            --oraex-blue: #0000FF;
            --oraex-blue-light: #4D4DFF;
            --oraex-blue-bg: #EEF2FF;
            --bg-white: #FFFFFF;
            --bg-light: #F9FAFB;
            --border: #E5E7EB;
            --text-dark: #111827;
            --text-gray: #6B7280;
            --text-muted: #9CA3AF;
            --success: #10B981;
            --warning: #F59E0B;
            --danger: #EF4444;
        }}
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        body {{ font-family: 'Inter', sans-serif; background: var(--bg-light); color: var(--text-dark); line-height: 1.6; }}
        
        .header {{ background: var(--bg-white); border-bottom: 1px solid var(--border); padding: 20px 0; position: sticky; top: 0; z-index: 100; box-shadow: 0 1px 3px rgba(0,0,0,0.05); }}
        .header-inner {{ max-width: 1200px; margin: 0 auto; padding: 0 40px; display: flex; align-items: center; justify-content: space-between; }}
        .logo-img {{ height: 50px; }}
        .header-meta {{ font-size: 0.85rem; color: var(--text-muted); }}
        
        .container {{ max-width: 1200px; margin: 0 auto; padding: 50px 40px; }}
        
        .hero {{ text-align: center; margin-bottom: 50px; animation: fadeIn 0.6s ease-out; }}
        @keyframes fadeIn {{ from {{ opacity: 0; transform: translateY(20px); }} to {{ opacity: 1; transform: translateY(0); }} }}
        .hero-badge {{ display: inline-flex; align-items: center; gap: 8px; padding: 8px 18px; background: var(--oraex-blue-bg); border: 1px solid rgba(0,0,255,0.2); border-radius: 100px; font-size: 0.8rem; color: var(--oraex-blue); font-weight: 600; margin-bottom: 20px; }}
        .hero-badge .dot {{ width: 8px; height: 8px; background: var(--oraex-blue); border-radius: 50%; animation: pulse 2s infinite; }}
        @keyframes pulse {{ 0%, 100% {{ opacity: 1; }} 50% {{ opacity: 0.7; }} }}
        .hero h1 {{ font-size: 2.5rem; font-weight: 800; letter-spacing: -0.03em; margin-bottom: 12px; }}
        .hero .subtitle {{ font-size: 1rem; color: var(--text-gray); }}
        
        .section-divider {{ border: none; border-top: 2px solid var(--oraex-blue); margin: 60px 0; opacity: 0.3; }}
        
        .part-header {{ text-align: center; margin-bottom: 40px; }}
        .part-header h2 {{ font-size: 1.8rem; font-weight: 700; color: var(--oraex-blue); margin-bottom: 8px; }}
        .part-header p {{ color: var(--text-gray); }}
        
        .kpi-grid {{ display: grid; grid-template-columns: repeat(5, 1fr); gap: 20px; margin-bottom: 50px; }}
        .kpi-card {{ background: var(--bg-white); border: 1px solid var(--border); border-radius: 16px; padding: 24px; transition: all 0.3s; animation: fadeIn 0.6s ease-out backwards; }}
        .kpi-card:hover {{ border-color: var(--oraex-blue); transform: translateY(-4px); box-shadow: 0 10px 40px -10px rgba(0,0,255,0.15); }}
        .kpi-card.primary {{ background: linear-gradient(135deg, var(--oraex-blue), var(--oraex-blue-light)); color: white; }}
        .kpi-card.primary .kpi-label, .kpi-card.primary .kpi-sub {{ color: rgba(255,255,255,0.8); }}
        .kpi-card.success {{ border-left: 4px solid var(--success); }}
        .kpi-card.warning {{ border-left: 4px solid var(--warning); }}
        .kpi-card.danger {{ border-left: 4px solid var(--danger); }}
        .kpi-icon {{ font-size: 1.3rem; margin-bottom: 12px; }}
        .kpi-value {{ font-size: 2.2rem; font-weight: 700; line-height: 1; margin-bottom: 4px; }}
        .kpi-card.primary .kpi-value {{ color: white; }}
        .kpi-label {{ font-size: 0.7rem; color: var(--text-muted); text-transform: uppercase; letter-spacing: 0.1em; font-weight: 600; }}
        .kpi-sub {{ margin-top: 10px; padding-top: 10px; border-top: 1px solid var(--border); font-size: 0.75rem; color: var(--text-gray); }}
        .kpi-card.primary .kpi-sub {{ border-top-color: rgba(255,255,255,0.2); }}
        
        .section {{ margin-bottom: 40px; }}
        .section-header {{ display: flex; align-items: center; gap: 14px; margin-bottom: 20px; }}
        .section-line {{ width: 4px; height: 24px; background: var(--oraex-blue); border-radius: 2px; }}
        .section-title {{ font-size: 1.15rem; font-weight: 700; }}
        
        .card {{ background: var(--bg-white); border: 1px solid var(--border); border-radius: 16px; padding: 24px; transition: all 0.3s; }}
        .card:hover {{ box-shadow: 0 4px 20px rgba(0,0,0,0.05); }}
        .card-title {{ font-size: 0.8rem; color: var(--text-gray); font-weight: 500; margin-bottom: 16px; }}
        
        .two-col {{ display: grid; grid-template-columns: 1fr 1fr; gap: 24px; }}
        .three-col {{ display: grid; grid-template-columns: repeat(3, 1fr); gap: 20px; }}
        .four-col {{ display: grid; grid-template-columns: repeat(4, 1fr); gap: 16px; }}
        
        .entorno-card {{ background: var(--bg-white); border: 1px solid var(--border); border-radius: 16px; padding: 24px; text-align: center; transition: all 0.3s; }}
        .entorno-card:hover {{ transform: translateY(-4px); box-shadow: 0 10px 30px rgba(0,0,0,0.08); }}
        .entorno-icon {{ width: 50px; height: 50px; margin: 0 auto 12px; border-radius: 12px; display: flex; align-items: center; justify-content: center; font-size: 1.2rem; font-weight: 700; }}
        .entorno-card.prod .entorno-icon {{ background: var(--oraex-blue-bg); color: var(--oraex-blue); }}
        .entorno-card.hml .entorno-icon {{ background: #FEF3C7; color: var(--warning); }}
        .entorno-card.dev .entorno-icon {{ background: #D1FAE5; color: var(--success); }}
        .entorno-card.trans .entorno-icon {{ background: #E0E7FF; color: #6366F1; }}
        .entorno-value {{ font-size: 1.8rem; font-weight: 700; }}
        .entorno-label {{ font-size: 0.8rem; color: var(--text-gray); margin-top: 4px; }}
        
        table {{ width: 100%; border-collapse: collapse; }}
        th {{ padding: 12px 14px; text-align: left; font-size: 0.65rem; text-transform: uppercase; letter-spacing: 0.1em; color: var(--text-muted); font-weight: 600; border-bottom: 2px solid var(--border); background: var(--bg-light); }}
        td {{ padding: 14px; border-bottom: 1px solid var(--border); font-size: 0.85rem; }}
        tr:hover td {{ background: var(--bg-light); }}
        
        .badge {{ display: inline-flex; padding: 4px 10px; border-radius: 100px; font-size: 0.7rem; font-weight: 600; }}
        .badge.blue {{ background: var(--oraex-blue-bg); color: var(--oraex-blue); }}
        .badge.green {{ background: #D1FAE5; color: var(--success); }}
        .badge.yellow {{ background: #FEF3C7; color: var(--warning); }}
        .badge.red {{ background: #FEE2E2; color: var(--danger); }}
        
        .progress-bar {{ width: 70px; height: 6px; background: #E5E7EB; border-radius: 100px; overflow: hidden; }}
        .progress-fill {{ height: 100%; background: var(--oraex-blue); border-radius: 100px; }}
        
        .alert {{ padding: 16px 20px; border-radius: 12px; margin-bottom: 20px; display: flex; align-items: flex-start; gap: 12px; }}
        .alert.warning {{ background: #FFFBEB; border: 1px solid #FDE68A; }}
        .alert.danger {{ background: #FEF2F2; border: 1px solid #FECACA; }}
        .alert-icon {{ font-size: 1.2rem; }}
        .alert-content {{ flex: 1; }}
        .alert-title {{ font-weight: 600; font-size: 0.9rem; margin-bottom: 4px; }}
        .alert-text {{ font-size: 0.8rem; color: var(--text-gray); }}
        
        .footer {{ text-align: center; padding: 40px 0; margin-top: 40px; border-top: 1px solid var(--border); background: var(--bg-white); }}
        .footer-logo img {{ height: 40px; }}
        .footer p {{ color: var(--text-muted); font-size: 0.85rem; margin-top: 12px; }}
        
        @media (max-width: 1100px) {{ .kpi-grid {{ grid-template-columns: repeat(3, 1fr); }} }}
        @media (max-width: 768px) {{ .kpi-grid, .two-col, .three-col, .four-col {{ grid-template-columns: 1fr; }} .hero h1 {{ font-size: 1.8rem; }} .container {{ padding: 30px 20px; }} }}
        @media print {{ body {{ background: white !important; }} .header {{ position: relative; }} }}
    </style>
</head>
<body>
    <header class="header">
        <div class="header-inner">
            {logo_html}
            <div class="header-meta">Gerado em {datetime.now().strftime('%d/%m/%Y às %H:%M')}</div>
        </div>
    </header>
    
    <div class="container">
        <div class="hero">
            <div class="hero-badge"><span class="dot"></span><span>Fevereiro — Dezembro 2025</span></div>
            <h1>Relatório Oracle GetNet</h1>
            <p class="subtitle">Consolidação Anual de Atualizações PSU e Inventário de Servidores</p>
        </div>
        
        <!-- PARTE 1: GMUDs -->
        <div class="part-header">
            <h2>📋 Atualizações PSU Executadas</h2>
            <p>Resumo das GMUDs de atualização PSU Oracle realizadas em 2025</p>
        </div>
        
        <div class="kpi-grid">
            <div class="kpi-card primary"><div class="kpi-icon">📊</div><div class="kpi-value">{total_psu:,}</div><div class="kpi-label">GMUDs PSU</div><div class="kpi-sub">Total de mudanças</div></div>
            <div class="kpi-card"><div class="kpi-icon">✅</div><div class="kpi-value">{sucesso_psu:,}</div><div class="kpi-label">Concluídas</div><div class="kpi-sub">{taxa_sucesso:.1f}% de sucesso</div></div>
            <div class="kpi-card"><div class="kpi-icon">🖥️</div><div class="kpi-value">{unique_servers:,}</div><div class="kpi-label">Servidores</div><div class="kpi-sub">Hosts únicos</div></div>
            <div class="kpi-card"><div class="kpi-icon">🔄</div><div class="kpi-value">{total_updates:,}</div><div class="kpi-label">Atualizações</div><div class="kpi-sub">Total intervenções</div></div>
            <div class="kpi-card"><div class="kpi-icon">⏱️</div><div class="kpi-value">{horas_totais:,}h</div><div class="kpi-label">Esforço</div><div class="kpi-sub">Horas trabalhadas</div></div>
        </div>
        
        <section class="section">
            <div class="section-header"><div class="section-line"></div><h2 class="section-title">Evolução Mensal</h2></div>
            <div class="card"><div id="chart-monthly"></div></div>
        </section>
        
        <section class="section">
            <div class="section-header"><div class="section-line"></div><h2 class="section-title">Resumo por Status</h2></div>
            <div class="card">
                <table>
                    <thead><tr><th>Status</th><th>Quantidade</th><th>Percentual</th></tr></thead>
                    <tbody>
                        <tr><td><span class="badge blue">✅ SUCESSO</span></td><td><strong>{sucesso_psu}</strong></td><td>{taxa_sucesso:.1f}%</td></tr>
                        <tr><td><span class="badge yellow">🔄 REPLANEJADA</span></td><td><strong>{replanejadas_psu}</strong></td><td>{replanejadas_psu/total_psu*100:.1f}%</td></tr>
                        <tr><td><span class="badge" style="background:#F3F4F6;color:#6B7280;">❌ CANCELADA</span></td><td><strong>{canceladas_psu}</strong></td><td>{canceladas_psu/total_psu*100:.1f}%</td></tr>
                        <tr><td><span class="badge" style="background:#F3F4F6;color:#6B7280;">📋 OUTRAS</span></td><td><strong>{outras_psu}</strong></td><td>{outras_psu/total_psu*100:.1f}%</td></tr>
                        <tr style="background: var(--bg-light); font-weight: 600;"><td>TOTAL</td><td><strong>{total_psu}</strong></td><td>100%</td></tr>
                    </tbody>
                </table>
            </div>
        </section>
        
        <section class="section">
            <div class="section-header"><div class="section-line"></div><h2 class="section-title">Performance da Equipe</h2></div>
            <div class="two-col">
                <div class="card"><div class="card-title">GMUDs por Executor</div><div id="chart-executor"></div></div>
                <div class="card"><div class="card-title">Ranking Detalhado</div>
                    <table><thead><tr><th>Executor</th><th>Total</th><th>Taxa</th><th>Horas</th></tr></thead><tbody>
"""
    for _, row in executor_stats.head(6).iterrows():
        html += f'<tr><td><strong>{row["Executor"]}</strong></td><td>{int(row["Total_GMUDs"])}</td><td><span class="badge blue">{row["Taxa"]:.0f}%</span></td><td style="color: var(--text-gray);">{int(row["Horas"])}h</td></tr>'
    
    html += f"""
                    </tbody></table>
                </div>
            </div>
        </section>
        
        <hr class="section-divider">
        
        <!-- PARTE 2: INVENTÁRIO -->
        <div class="part-header">
            <h2>🖥️ Inventário de Servidores Oracle</h2>
            <p>Situação atual dos servidores Oracle na infraestrutura GetNet</p>
        </div>
        
        <div class="kpi-grid">
            <div class="kpi-card primary"><div class="kpi-icon">🖥️</div><div class="kpi-value">{total_servidores}</div><div class="kpi-label">Servidores Ativos</div><div class="kpi-sub">Infraestrutura Oracle</div></div>
            <div class="kpi-card success"><div class="kpi-icon">✅</div><div class="kpi-value">{srv_atualizados}</div><div class="kpi-label">Atualizados</div><div class="kpi-sub">PSU {LATEST_PSU} ({pct_atualizados:.0f}%)</div></div>
            <div class="kpi-card"><div class="kpi-icon">📦</div><div class="kpi-value">{srv_1q_atras}</div><div class="kpi-label">1 Quarter Atrás</div><div class="kpi-sub">PSU 19.28</div></div>
            <div class="kpi-card warning"><div class="kpi-icon">⚠️</div><div class="kpi-value">{srv_2q_atras}</div><div class="kpi-label">2 Quarters Atrás</div><div class="kpi-sub">PSU 19.27</div></div>
            <div class="kpi-card danger"><div class="kpi-icon">🚨</div><div class="kpi-value">{srv_3q_mais}</div><div class="kpi-label">3+ Quarters</div><div class="kpi-sub">Atenção necessária</div></div>
        </div>
"""
    
    if srv_3q_mais > 0:
        html += f"""
        <div class="alert danger">
            <div class="alert-icon">🚨</div>
            <div class="alert-content">
                <div class="alert-title">Servidores Críticos Identificados</div>
                <div class="alert-text">{srv_3q_mais} servidores estão com PSU de 3 ou mais quarters atrás. Recomenda-se priorizar a atualização destes hosts.</div>
            </div>
        </div>
"""
    
    html += f"""
        <section class="section">
            <div class="section-header"><div class="section-line"></div><h2 class="section-title">Distribuição por Ambiente</h2></div>
            <div class="four-col">
                <div class="entorno-card prod"><div class="entorno-icon">P</div><div class="entorno-value">{inv_prod}</div><div class="entorno-label">Produção</div></div>
                <div class="entorno-card hml"><div class="entorno-icon">H</div><div class="entorno-value">{inv_hml}</div><div class="entorno-label">Homologação</div></div>
                <div class="entorno-card dev"><div class="entorno-icon">D</div><div class="entorno-value">{inv_dev}</div><div class="entorno-label">Desenvolvimento</div></div>
                <div class="entorno-card trans"><div class="entorno-icon">T</div><div class="entorno-value">{inv_trans}</div><div class="entorno-label">Transacional</div></div>
            </div>
        </section>
        
        <section class="section">
            <div class="section-header"><div class="section-line"></div><h2 class="section-title">Distribuição por Versão PSU</h2></div>
            <div class="two-col">
                <div class="card"><div class="card-title">Servidores por Versão</div><div id="chart-inv-versao"></div></div>
                <div class="card"><div class="card-title">Servidores por Ambiente</div><div id="chart-inv-entorno"></div></div>
            </div>
        </section>
"""
    
    if len(criticos) > 0:
        html += f"""
        <section class="section">
            <div class="section-header"><div class="section-line"></div><h2 class="section-title">Servidores com Atenção Necessária</h2></div>
            <div class="card">
                <table>
                    <thead><tr><th>Hostname</th><th>Versão PSU</th><th>Ambiente</th><th>Status</th></tr></thead>
                    <tbody>
"""
        for _, row in criticos.iterrows():
            html += f'<tr><td><strong>{row["Hostname"]}</strong></td><td>{row["PSU_Version"]}</td><td>{row["Entorno"]}</td><td><span class="badge red">Desatualizado</span></td></tr>'
        html += """
                    </tbody>
                </table>
            </div>
        </section>
"""
    
    html += f"""
    </div>
    
    <footer class="footer">
        <div class="footer-logo">{logo_html}</div>
        <p>Relatório Confidencial • GetNet Infrastructure Operations</p>
        <p>© 2025 ORAEX Cloud Consulting</p>
    </footer>
    
    <script>
        Plotly.newPlot('chart-monthly', {fig_monthly.to_json()}.data, {fig_monthly.to_json()}.layout, {{responsive: true, displayModeBar: false}});
        Plotly.newPlot('chart-executor', {fig_exec.to_json()}.data, {fig_exec.to_json()}.layout, {{responsive: true, displayModeBar: false}});
        Plotly.newPlot('chart-inv-versao', {fig_inv_versao.to_json()}.data, {fig_inv_versao.to_json()}.layout, {{responsive: true, displayModeBar: false}});
        Plotly.newPlot('chart-inv-entorno', {fig_inv_entorno.to_json()}.data, {fig_inv_entorno.to_json()}.layout, {{responsive: true, displayModeBar: false}});
        
        document.querySelectorAll('.kpi-value').forEach(el => {{
            const text = el.textContent;
            const num = parseInt(text.replace(/[^0-9]/g, ''));
            if (!isNaN(num) && num > 0) {{
                let current = 0;
                const step = Math.ceil(num / 40);
                const suffix = text.includes('h') ? 'h' : (text.includes('%') ? '%' : '');
                el.textContent = '0' + suffix;
                setTimeout(() => {{
                    const interval = setInterval(() => {{
                        current += step;
                        if (current >= num) {{ current = num; clearInterval(interval); }}
                        el.textContent = current.toLocaleString('pt-BR') + suffix;
                    }}, 25);
                }}, 300);
            }}
        }});
    </script>
</body>
</html>
"""
    
    with open(OUTPUT_HTML, 'w', encoding='utf-8') as f:
        f.write(html)
    
    print(f"\n✅ Relatório V6 COMPLETO gerado: {OUTPUT_HTML}")

if __name__ == "__main__":
    print("="*60)
    print("GERANDO RELATÓRIO V6 - GMUDS + INVENTÁRIO")
    print("="*60)
    generate_complete_report()
