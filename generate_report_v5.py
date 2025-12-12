import pandas as pd
import plotly.graph_objects as go
from datetime import datetime
import re
import base64

FILE_PATH = r"D:\antigravity\oraex\cmdb\ORAEX - Consolidação GetTech 2025 (1).xlsm"
LOGO_PATH = r"D:\antigravity\oraex\cmdb\oraex_logo.png"
OUTPUT_HTML = r"D:\antigravity\oraex\cmdb\relatorio_psu_2025_v5_oraex.html"

MONTHLY_SHEETS = [
    'FEVEREIRO-25', 'MARÇO-25', 'ABRIL-25', 'MAIO-25', 'JUNHO-25',
    'JULHO-25', 'AGOSTO-25', 'SETEMBRO-25', 'OUTUBRO-25', 'NOVEMBRO-25', 'DEZEMBRO-25'
]

MONTH_ORDER = ['FEVEREIRO', 'MARÇO', 'ABRIL', 'MAIO', 'JUNHO', 'JULHO', 
               'AGOSTO', 'SETEMBRO', 'OUTUBRO', 'NOVEMBRO', 'DEZEMBRO']

# Converter logo para base64 para embutir no HTML
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
    if 'Guilherme' in resp: return 'Guilherme Fonseca'
    if 'Bruno' in resp: return 'Bruno Ferreira'
    if 'Alcides' in resp: return 'Alcides Souto'
    if 'Kaue' in resp: return 'Kaue Santos'
    if 'Rafael' in resp: return 'Rafael Rabello'
    if 'Luca' in resp: return 'Luca Mozart'
    if 'Jonathan' in resp: return 'Jonathan Ferreira'
    return resp

def load_all_gmuds():
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

def generate_oraex_blue_report(df):
    print("Processando dados...")
    
    logo_b64 = get_logo_base64()
    
    df['Status_Final'] = df['Status'].apply(normalize_status)
    df['Hostnames'] = df['Titulo'].apply(extract_hostnames)
    df['Num_Servers'] = df['Hostnames'].apply(len)
    df['Versao_PSU'] = df['Titulo'].apply(extract_psu_version)
    df['Responsavel_Norm'] = df['Responsavel'].apply(normalize_responsavel)
    df['Is_PSU'] = df['Titulo'].str.contains('PSU', case=False, na=False)
    
    df_psu = df[df['Is_PSU']].copy()
    
    # Métricas
    total_psu = len(df_psu)
    sucesso_psu = len(df_psu[df_psu['Status_Final'] == 'SUCESSO'])
    canceladas = len(df_psu[df_psu['Status_Final'] == 'CANCELADA'])
    replanejadas = len(df_psu[df_psu['Status_Final'] == 'REPLANEJADA'])
    
    all_hosts = []
    for h in df_psu[df_psu['Status_Final'] == 'SUCESSO']['Hostnames']:
        all_hosts.extend(h)
    unique_servers = len(set(all_hosts))
    total_updates = len(all_hosts)
    horas_totais = total_updates * 3
    taxa_sucesso = sucesso_psu / total_psu * 100 if total_psu > 0 else 0
    
    # Entornos
    entorno_prod = len(df_psu[df_psu['Entorno'] == 'P'])
    entorno_hml = len(df_psu[df_psu['Entorno'] == 'H'])
    entorno_dev = len(df_psu[df_psu['Entorno'] == 'D'])
    
    # Versões
    versao_stats = df_psu.groupby('Versao_PSU').agg({
        'GMUD_ID': 'count',
        'Status_Final': lambda x: (x == 'SUCESSO').sum()
    }).reset_index()
    versao_stats.columns = ['Versao', 'Total', 'Sucesso']
    versao_stats = versao_stats.dropna(subset=['Versao']).sort_values('Versao')
    
    # Executores
    executor_stats = df_psu.groupby('Responsavel_Norm').agg({
        'GMUD_ID': 'count',
        'Status_Final': lambda x: (x == 'SUCESSO').sum(),
        'Num_Servers': 'sum'
    }).reset_index()
    executor_stats.columns = ['Executor', 'Total_GMUDs', 'Sucesso', 'Servidores']
    executor_stats['Taxa'] = (executor_stats['Sucesso'] / executor_stats['Total_GMUDs'] * 100).round(1)
    executor_stats['Horas'] = executor_stats['Servidores'] * 3
    executor_stats = executor_stats.sort_values('Total_GMUDs', ascending=False)
    executor_stats = executor_stats[executor_stats['Executor'] != 'Não Atribuído'].head(8)
    
    # Mensal
    monthly_psu = df_psu.groupby(['Mes', 'Status_Final']).size().unstack(fill_value=0)
    monthly_psu = monthly_psu.reindex(MONTH_ORDER)
    
    # Cor Azul Oraex (extraída da logo)
    oraex_blue = '#0000FF'
    oraex_blue_light = '#4D4DFF'
    oraex_blue_dark = '#0000CC'
    
    # Gráfico mensal
    fig_monthly = go.Figure()
    colors = {'SUCESSO': oraex_blue, 'REPLANEJADA': '#93C5FD', 'CANCELADA': '#D1D5DB'}
    for status in ['SUCESSO', 'REPLANEJADA', 'CANCELADA']:
        if status in monthly_psu.columns:
            fig_monthly.add_trace(go.Bar(
                name=status, x=monthly_psu.index, y=monthly_psu[status],
                marker_color=colors.get(status),
                marker_line_width=0
            ))
    fig_monthly.update_layout(
        barmode='stack',
        paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
        font=dict(color='#374151', family='Inter, sans-serif', size=12),
        height=380,
        xaxis=dict(showgrid=False, tickfont=dict(size=11, color='#6B7280')),
        yaxis=dict(showgrid=True, gridcolor='rgba(0,0,0,0.06)', gridwidth=1),
        legend=dict(orientation="h", y=1.12, x=0.5, xanchor='center', font=dict(size=11)),
        margin=dict(t=50, b=50, l=50, r=30),
        bargap=0.3
    )
    
    # Gráfico executores
    fig_exec = go.Figure()
    fig_exec.add_trace(go.Bar(
        y=executor_stats['Executor'].head(6),
        x=executor_stats['Sucesso'].head(6),
        name='Concluídas',
        orientation='h',
        marker_color=oraex_blue,
        marker_line_width=0
    ))
    insucesso = executor_stats['Total_GMUDs'].head(6) - executor_stats['Sucesso'].head(6)
    fig_exec.add_trace(go.Bar(
        y=executor_stats['Executor'].head(6),
        x=insucesso,
        name='Outras',
        orientation='h',
        marker_color='#E5E7EB'
    ))
    fig_exec.update_layout(
        barmode='stack',
        paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
        font=dict(color='#374151', family='Inter, sans-serif', size=12),
        height=320,
        xaxis=dict(showgrid=True, gridcolor='rgba(0,0,0,0.06)'),
        yaxis=dict(showgrid=False),
        legend=dict(orientation="h", y=1.08, x=0.5, xanchor='center'),
        margin=dict(l=130, r=30, t=40, b=30),
        bargap=0.25
    )
    
    # Logo embed
    logo_html = f'<img src="data:image/png;base64,{logo_b64}" alt="ORAEX" class="logo-img">' if logo_b64 else '<span class="logo-text">ORAEX</span>'
    
    # HTML com fundo branco e azul Oraex
    html = f"""
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Relatório PSU Oracle 2025 | ORAEX</title>
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap" rel="stylesheet">
    <script src="https://cdn.plot.ly/plotly-2.27.0.min.js"></script>
    <style>
        :root {{
            /* Paleta ORAEX - Azul e Branco */
            --oraex-blue: #0000FF;
            --oraex-blue-light: #4D4DFF;
            --oraex-blue-dark: #0000CC;
            --oraex-blue-bg: #EEF2FF;
            --oraex-blue-glow: rgba(0, 0, 255, 0.15);
            
            --bg-white: #FFFFFF;
            --bg-light: #F9FAFB;
            --bg-card: #FFFFFF;
            --border: #E5E7EB;
            --border-hover: #D1D5DB;
            
            --text-dark: #111827;
            --text-gray: #6B7280;
            --text-muted: #9CA3AF;
            
            --success: #10B981;
            --warning: #F59E0B;
            --danger: #EF4444;
        }}
        
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        
        body {{
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
            background: var(--bg-light);
            color: var(--text-dark);
            min-height: 100vh;
            line-height: 1.6;
        }}
        
        /* Header */
        .header {{
            background: var(--bg-white);
            border-bottom: 1px solid var(--border);
            padding: 20px 0;
            position: sticky;
            top: 0;
            z-index: 100;
            box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        }}
        
        .header-inner {{
            max-width: 1200px;
            margin: 0 auto;
            padding: 0 40px;
            display: flex;
            align-items: center;
            justify-content: space-between;
        }}
        
        .logo-img {{
            height: 50px;
            width: auto;
        }}
        
        .logo-text {{
            font-size: 1.8rem;
            font-weight: 800;
            color: var(--oraex-blue);
            letter-spacing: -0.02em;
        }}
        
        .header-meta {{
            font-size: 0.85rem;
            color: var(--text-muted);
        }}
        
        .container {{
            max-width: 1200px;
            margin: 0 auto;
            padding: 50px 40px;
        }}
        
        /* Hero */
        .hero {{
            text-align: center;
            margin-bottom: 60px;
            animation: fadeIn 0.6s ease-out;
        }}
        
        @keyframes fadeIn {{
            from {{ opacity: 0; transform: translateY(20px); }}
            to {{ opacity: 1; transform: translateY(0); }}
        }}
        
        .hero-badge {{
            display: inline-flex;
            align-items: center;
            gap: 8px;
            padding: 8px 18px;
            background: var(--oraex-blue-bg);
            border: 1px solid rgba(0, 0, 255, 0.2);
            border-radius: 100px;
            font-size: 0.8rem;
            color: var(--oraex-blue);
            font-weight: 600;
            margin-bottom: 20px;
        }}
        
        .hero-badge .dot {{
            width: 8px; height: 8px;
            background: var(--oraex-blue);
            border-radius: 50%;
            animation: pulse 2s infinite;
        }}
        
        @keyframes pulse {{
            0%, 100% {{ opacity: 1; transform: scale(1); }}
            50% {{ opacity: 0.7; transform: scale(0.9); }}
        }}
        
        .hero h1 {{
            font-size: 2.8rem;
            font-weight: 800;
            letter-spacing: -0.03em;
            color: var(--text-dark);
            margin-bottom: 12px;
        }}
        
        .hero .subtitle {{
            font-size: 1.1rem;
            color: var(--text-gray);
            font-weight: 400;
        }}
        
        /* KPI Grid */
        .kpi-grid {{
            display: grid;
            grid-template-columns: repeat(5, 1fr);
            gap: 20px;
            margin-bottom: 60px;
        }}
        
        .kpi-card {{
            background: var(--bg-white);
            border: 1px solid var(--border);
            border-radius: 16px;
            padding: 28px;
            transition: all 0.3s ease;
            animation: fadeIn 0.6s ease-out backwards;
            position: relative;
        }}
        
        .kpi-card:nth-child(1) {{ animation-delay: 0.1s; }}
        .kpi-card:nth-child(2) {{ animation-delay: 0.15s; }}
        .kpi-card:nth-child(3) {{ animation-delay: 0.2s; }}
        .kpi-card:nth-child(4) {{ animation-delay: 0.25s; }}
        .kpi-card:nth-child(5) {{ animation-delay: 0.3s; }}
        
        .kpi-card:hover {{
            border-color: var(--oraex-blue);
            box-shadow: 0 10px 40px -10px var(--oraex-blue-glow);
            transform: translateY(-4px);
        }}
        
        .kpi-card.primary {{
            background: linear-gradient(135deg, var(--oraex-blue) 0%, var(--oraex-blue-dark) 100%);
            border: none;
            color: white;
        }}
        
        .kpi-card.primary .kpi-label,
        .kpi-card.primary .kpi-sub {{
            color: rgba(255,255,255,0.8);
        }}
        
        .kpi-icon {{
            font-size: 1.4rem;
            margin-bottom: 16px;
        }}
        
        .kpi-value {{
            font-size: 2.4rem;
            font-weight: 700;
            color: var(--text-dark);
            line-height: 1;
            margin-bottom: 6px;
            letter-spacing: -0.02em;
        }}
        
        .kpi-card.primary .kpi-value {{
            color: white;
        }}
        
        .kpi-label {{
            font-size: 0.75rem;
            color: var(--text-muted);
            text-transform: uppercase;
            letter-spacing: 0.1em;
            font-weight: 600;
        }}
        
        .kpi-sub {{
            margin-top: 12px;
            padding-top: 12px;
            border-top: 1px solid var(--border);
            font-size: 0.8rem;
            color: var(--text-gray);
        }}
        
        .kpi-card.primary .kpi-sub {{
            border-top-color: rgba(255,255,255,0.2);
        }}
        
        /* Sections */
        .section {{
            margin-bottom: 50px;
            animation: fadeIn 0.6s ease-out backwards;
        }}
        
        .section-header {{
            display: flex;
            align-items: center;
            gap: 14px;
            margin-bottom: 24px;
        }}
        
        .section-line {{
            width: 4px;
            height: 26px;
            background: var(--oraex-blue);
            border-radius: 2px;
        }}
        
        .section-title {{
            font-size: 1.25rem;
            font-weight: 700;
            color: var(--text-dark);
        }}
        
        /* Cards */
        .card {{
            background: var(--bg-white);
            border: 1px solid var(--border);
            border-radius: 16px;
            padding: 28px;
            transition: all 0.3s ease;
        }}
        
        .card:hover {{
            border-color: var(--border-hover);
            box-shadow: 0 4px 20px rgba(0,0,0,0.05);
        }}
        
        .card-title {{
            font-size: 0.85rem;
            color: var(--text-gray);
            font-weight: 500;
            margin-bottom: 20px;
        }}
        
        .two-col {{ display: grid; grid-template-columns: 1fr 1fr; gap: 24px; }}
        .three-col {{ display: grid; grid-template-columns: repeat(3, 1fr); gap: 20px; }}
        
        /* Entorno */
        .entorno-card {{
            background: var(--bg-white);
            border: 1px solid var(--border);
            border-radius: 16px;
            padding: 32px;
            text-align: center;
            transition: all 0.3s ease;
        }}
        
        .entorno-card:hover {{
            transform: translateY(-4px);
            box-shadow: 0 10px 30px rgba(0,0,0,0.08);
        }}
        
        .entorno-icon {{
            width: 56px; height: 56px;
            margin: 0 auto 16px;
            border-radius: 14px;
            display: flex; align-items: center; justify-content: center;
            font-size: 1.3rem; font-weight: 700;
        }}
        
        .entorno-card.prod .entorno-icon {{ background: var(--oraex-blue-bg); color: var(--oraex-blue); }}
        .entorno-card.hml .entorno-icon {{ background: #FEF3C7; color: var(--warning); }}
        .entorno-card.dev .entorno-icon {{ background: #D1FAE5; color: var(--success); }}
        
        .entorno-value {{ font-size: 2rem; font-weight: 700; color: var(--text-dark); }}
        .entorno-label {{ font-size: 0.85rem; color: var(--text-gray); margin-top: 4px; }}
        
        /* Tables */
        table {{ width: 100%; border-collapse: collapse; }}
        
        th {{
            padding: 14px 16px;
            text-align: left;
            font-size: 0.7rem;
            text-transform: uppercase;
            letter-spacing: 0.1em;
            color: var(--text-muted);
            font-weight: 600;
            border-bottom: 2px solid var(--border);
            background: var(--bg-light);
        }}
        
        td {{
            padding: 16px;
            border-bottom: 1px solid var(--border);
            font-size: 0.9rem;
        }}
        
        tr:hover td {{ background: var(--bg-light); }}
        
        .badge {{
            display: inline-flex;
            padding: 5px 12px;
            border-radius: 100px;
            font-size: 0.75rem;
            font-weight: 600;
        }}
        
        .badge.blue {{ background: var(--oraex-blue-bg); color: var(--oraex-blue); }}
        .badge.green {{ background: #D1FAE5; color: var(--success); }}
        .badge.gray {{ background: #F3F4F6; color: var(--text-gray); }}
        
        .progress-bar {{
            width: 80px; height: 6px;
            background: #E5E7EB;
            border-radius: 100px;
            overflow: hidden;
        }}
        
        .progress-fill {{
            height: 100%;
            background: var(--oraex-blue);
            border-radius: 100px;
        }}
        
        /* Footer */
        .footer {{
            text-align: center;
            padding: 40px 0;
            margin-top: 40px;
            border-top: 1px solid var(--border);
            background: var(--bg-white);
        }}
        
        .footer-logo {{
            margin-bottom: 12px;
        }}
        
        .footer-logo img {{
            height: 40px;
        }}
        
        .footer p {{
            color: var(--text-muted);
            font-size: 0.85rem;
        }}
        
        @media (max-width: 1100px) {{
            .kpi-grid {{ grid-template-columns: repeat(3, 1fr); }}
        }}
        
        @media (max-width: 768px) {{
            .kpi-grid, .two-col, .three-col {{ grid-template-columns: 1fr; }}
            .hero h1 {{ font-size: 2rem; }}
            .container {{ padding: 30px 20px; }}
        }}
        
        @media print {{
            body {{ background: white !important; }}
            .header {{ position: relative; }}
        }}
    </style>
</head>
<body>
    <!-- Header -->
    <header class="header">
        <div class="header-inner">
            {logo_html}
            <div class="header-meta">
                Gerado em {datetime.now().strftime('%d/%m/%Y às %H:%M')}
            </div>
        </div>
    </header>
    
    <div class="container">
        <!-- Hero -->
        <div class="hero">
            <div class="hero-badge">
                <span class="dot"></span>
                <span>Fevereiro — Dezembro 2025</span>
            </div>
            <h1>Relatório PSU Oracle</h1>
            <p class="subtitle">Consolidação Anual de Atualizações • GetNet Infrastructure</p>
        </div>
        
        <!-- KPIs -->
        <div class="kpi-grid">
            <div class="kpi-card primary">
                <div class="kpi-icon">📊</div>
                <div class="kpi-value" data-count="{total_psu}">{total_psu:,}</div>
                <div class="kpi-label">GMUDs PSU</div>
                <div class="kpi-sub">Total de mudanças</div>
            </div>
            <div class="kpi-card">
                <div class="kpi-icon">✅</div>
                <div class="kpi-value" data-count="{sucesso_psu}">{sucesso_psu:,}</div>
                <div class="kpi-label">Concluídas</div>
                <div class="kpi-sub">{taxa_sucesso:.1f}% de sucesso</div>
            </div>
            <div class="kpi-card">
                <div class="kpi-icon">🖥️</div>
                <div class="kpi-value" data-count="{unique_servers}">{unique_servers:,}</div>
                <div class="kpi-label">Servidores</div>
                <div class="kpi-sub">Hosts únicos</div>
            </div>
            <div class="kpi-card">
                <div class="kpi-icon">🔄</div>
                <div class="kpi-value" data-count="{total_updates}">{total_updates:,}</div>
                <div class="kpi-label">Atualizações</div>
                <div class="kpi-sub">Total de intervenções</div>
            </div>
            <div class="kpi-card">
                <div class="kpi-icon">⏱️</div>
                <div class="kpi-value" data-count="{horas_totais}">{horas_totais:,}h</div>
                <div class="kpi-label">Esforço</div>
                <div class="kpi-sub">Horas trabalhadas</div>
            </div>
        </div>
        
        <!-- Entornos -->
        <section class="section">
            <div class="section-header">
                <div class="section-line"></div>
                <h2 class="section-title">Distribuição por Ambiente</h2>
            </div>
            <div class="three-col">
                <div class="entorno-card prod">
                    <div class="entorno-icon">P</div>
                    <div class="entorno-value">{entorno_prod}</div>
                    <div class="entorno-label">Produção</div>
                </div>
                <div class="entorno-card hml">
                    <div class="entorno-icon">H</div>
                    <div class="entorno-value">{entorno_hml}</div>
                    <div class="entorno-label">Homologação</div>
                </div>
                <div class="entorno-card dev">
                    <div class="entorno-icon">D</div>
                    <div class="entorno-value">{entorno_dev}</div>
                    <div class="entorno-label">Desenvolvimento</div>
                </div>
            </div>
        </section>
        
        <!-- Versões -->
        <section class="section">
            <div class="section-header">
                <div class="section-line"></div>
                <h2 class="section-title">Versões PSU Aplicadas</h2>
            </div>
            <div class="card">
                <table>
                    <thead>
                        <tr>
                            <th>Versão PSU</th>
                            <th>Total</th>
                            <th>Sucesso</th>
                            <th>Taxa</th>
                        </tr>
                    </thead>
                    <tbody>
"""
    
    for _, row in versao_stats.iterrows():
        taxa = row['Sucesso'] / row['Total'] * 100 if row['Total'] > 0 else 0
        html += f"""
                        <tr>
                            <td><strong>{row['Versao']}</strong></td>
                            <td>{int(row['Total'])}</td>
                            <td><span class="badge blue">{int(row['Sucesso'])}</span></td>
                            <td>
                                <div style="display: flex; align-items: center; gap: 10px;">
                                    <div class="progress-bar"><div class="progress-fill" style="width: {taxa}%;"></div></div>
                                    <span style="color: var(--text-gray); font-size: 0.85rem;">{taxa:.0f}%</span>
                                </div>
                            </td>
                        </tr>
"""
    
    html += f"""
                    </tbody>
                </table>
            </div>
        </section>
        
        <!-- Evolução Mensal -->
        <section class="section">
            <div class="section-header">
                <div class="section-line"></div>
                <h2 class="section-title">Evolução Mensal</h2>
            </div>
            <div class="card">
                <div id="chart-monthly"></div>
            </div>
        </section>
        
        <!-- Executores -->
        <section class="section">
            <div class="section-header">
                <div class="section-line"></div>
                <h2 class="section-title">Performance da Equipe</h2>
            </div>
            <div class="two-col">
                <div class="card">
                    <div class="card-title">GMUDs por Executor</div>
                    <div id="chart-executor"></div>
                </div>
                <div class="card">
                    <div class="card-title">Ranking Detalhado</div>
                    <table>
                        <thead>
                            <tr>
                                <th>Executor</th>
                                <th>Total</th>
                                <th>Taxa</th>
                                <th>Horas</th>
                            </tr>
                        </thead>
                        <tbody>
"""
    
    for _, row in executor_stats.head(6).iterrows():
        html += f"""
                            <tr>
                                <td><strong>{row['Executor']}</strong></td>
                                <td>{int(row['Total_GMUDs'])}</td>
                                <td><span class="badge blue">{row['Taxa']:.0f}%</span></td>
                                <td style="color: var(--text-gray);">{int(row['Horas'])}h</td>
                            </tr>
"""
    
    html += f"""
                        </tbody>
                    </table>
                </div>
            </div>
        </section>
    </div>
    
    <!-- Footer -->
    <footer class="footer">
        <div class="footer-logo">
            {logo_html}
        </div>
        <p>Relatório Confidencial • GetNet Infrastructure Operations</p>
        <p style="margin-top: 4px;">© 2025 ORAEX Cloud Consulting</p>
    </footer>
    
    <script>
        Plotly.newPlot('chart-monthly', {fig_monthly.to_json()}.data, {fig_monthly.to_json()}.layout, {{responsive: true, displayModeBar: false}});
        Plotly.newPlot('chart-executor', {fig_exec.to_json()}.data, {fig_exec.to_json()}.layout, {{responsive: true, displayModeBar: false}});
        
        // Animação de contagem
        document.querySelectorAll('.kpi-value').forEach(el => {{
            const text = el.textContent;
            const num = parseInt(text.replace(/[^0-9]/g, ''));
            if (!isNaN(num) && num > 0) {{
                let current = 0;
                const step = Math.ceil(num / 40);
                const suffix = text.includes('h') ? 'h' : '';
                el.textContent = '0' + suffix;
                setTimeout(() => {{
                    const interval = setInterval(() => {{
                        current += step;
                        if (current >= num) {{
                            current = num;
                            clearInterval(interval);
                        }}
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
    
    print(f"\n✅ Relatório V5 ORAEX (Azul/Branco) gerado: {OUTPUT_HTML}")

if __name__ == "__main__":
    print("="*60)
    print("GERANDO RELATÓRIO PSU 2025 - V5 ORAEX AZUL/BRANCO")
    print("="*60)
    df = load_all_gmuds()
    if not df.empty:
        generate_oraex_blue_report(df)
