import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import re
import json

FILE_PATH = r"D:\antigravity\oraex\cmdb\ORAEX - Consolidação GetTech 2025 (1).xlsm"
OUTPUT_HTML = r"D:\antigravity\oraex\cmdb\relatorio_psu_2025_v3_premium.html"

MONTHLY_SHEETS = [
    'FEVEREIRO-25', 'MARÇO-25', 'ABRIL-25', 'MAIO-25', 'JUNHO-25',
    'JULHO-25', 'AGOSTO-25', 'SETEMBRO-25', 'OUTUBRO-25', 'NOVEMBRO-25', 'DEZEMBRO-25'
]

MONTH_ORDER = ['FEVEREIRO', 'MARÇO', 'ABRIL', 'MAIO', 'JUNHO', 'JULHO', 
               'AGOSTO', 'SETEMBRO', 'OUTUBRO', 'NOVEMBRO', 'DEZEMBRO']

ENTORNO_MAP = {'P': 'Produção', 'H': 'Homologação', 'D': 'Desenvolvimento', 'T': 'Transacional'}

def extract_hostnames(title):
    if pd.isna(title): return []
    pattern = r'(gncas[a-z0-9]+)'
    return [m.upper() for m in re.findall(pattern, str(title).lower())]

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
    elif 'ANDAMENTO' in status or 'EXECUÇÃO' in status or 'IMPLEMENT' in status: return 'EM ANDAMENTO'
    elif 'PROGRAM' in status or 'NOVO' in status or 'AUTORIZAR' in status or 'CAB' in status or 'AVALIAR' in status: return 'PENDENTE'
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

def generate_premium_report(df):
    print("Processando dados...")
    
    # Enrich
    df['Status_Final'] = df['Status'].apply(normalize_status)
    df['Hostnames'] = df['Titulo'].apply(extract_hostnames)
    df['Num_Servers'] = df['Hostnames'].apply(len)
    df['Versao_PSU'] = df['Titulo'].apply(extract_psu_version)
    df['Responsavel_Norm'] = df['Responsavel'].apply(normalize_responsavel)
    df['Entorno_Nome'] = df['Entorno'].map(ENTORNO_MAP).fillna('Outros')
    df['Is_PSU'] = df['Titulo'].str.contains('PSU', case=False, na=False)
    
    df_psu = df[df['Is_PSU']].copy()
    
    # Metrics
    total_psu = len(df_psu)
    sucesso_psu = len(df_psu[df_psu['Status_Final'] == 'SUCESSO'])
    canceladas = len(df_psu[df_psu['Status_Final'] == 'CANCELADA'])
    replanejadas = len(df_psu[df_psu['Status_Final'] == 'REPLANEJADA'])
    insucesso = len(df_psu[df_psu['Status_Final'] == 'INSUCESSO'])
    
    all_hosts = []
    for h in df_psu[df_psu['Status_Final'] == 'SUCESSO']['Hostnames']:
        all_hosts.extend(h)
    unique_servers = len(set(all_hosts))
    total_updates = len(all_hosts)
    horas_totais = total_updates * 3
    taxa_sucesso = sucesso_psu / total_psu * 100 if total_psu > 0 else 0
    
    # Version stats
    versao_stats = df_psu.groupby('Versao_PSU').agg({
        'GMUD_ID': 'count',
        'Status_Final': lambda x: (x == 'SUCESSO').sum()
    }).reset_index()
    versao_stats.columns = ['Versao', 'Total', 'Sucesso']
    versao_stats = versao_stats.dropna(subset=['Versao']).sort_values('Versao')
    
    # Entorno stats
    entorno_prod = len(df_psu[df_psu['Entorno'] == 'P'])
    entorno_hml = len(df_psu[df_psu['Entorno'] == 'H'])
    entorno_dev = len(df_psu[df_psu['Entorno'] == 'D'])
    
    # Executor stats
    executor_stats = df_psu.groupby('Responsavel_Norm').agg({
        'GMUD_ID': 'count',
        'Status_Final': lambda x: (x == 'SUCESSO').sum(),
        'Num_Servers': 'sum'
    }).reset_index()
    executor_stats.columns = ['Executor', 'Total_GMUDs', 'Sucesso', 'Servidores']
    executor_stats['Insucesso'] = executor_stats['Total_GMUDs'] - executor_stats['Sucesso']
    executor_stats['Taxa'] = (executor_stats['Sucesso'] / executor_stats['Total_GMUDs'] * 100).round(1)
    executor_stats['Horas'] = executor_stats['Servidores'] * 3
    executor_stats = executor_stats.sort_values('Total_GMUDs', ascending=False)
    executor_stats = executor_stats[executor_stats['Executor'] != 'Não Atribuído'].head(8)
    
    # Monthly
    monthly_psu = df_psu.groupby(['Mes', 'Status_Final']).size().unstack(fill_value=0)
    monthly_psu = monthly_psu.reindex(MONTH_ORDER)
    
    # Charts
    fig_monthly = go.Figure()
    colors = {'SUCESSO': '#10b981', 'REPLANEJADA': '#f59e0b', 'CANCELADA': '#ef4444'}
    for status in ['SUCESSO', 'REPLANEJADA', 'CANCELADA']:
        if status in monthly_psu.columns:
            fig_monthly.add_trace(go.Bar(name=status, x=monthly_psu.index, y=monthly_psu[status], marker_color=colors.get(status)))
    fig_monthly.update_layout(barmode='stack', paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
        font=dict(color='#94a3b8', family='Outfit'), height=400,
        xaxis=dict(showgrid=False, tickfont=dict(size=11)), yaxis=dict(showgrid=True, gridcolor='rgba(148,163,184,0.1)'),
        legend=dict(orientation="h", y=1.15, font=dict(size=11)), margin=dict(t=40, b=40))
    
    # Executor chart
    fig_exec = go.Figure()
    fig_exec.add_trace(go.Bar(y=executor_stats['Executor'].head(6), x=executor_stats['Sucesso'].head(6), name='Sucesso', orientation='h', marker_color='#10b981'))
    fig_exec.add_trace(go.Bar(y=executor_stats['Executor'].head(6), x=executor_stats['Insucesso'].head(6), name='Outros', orientation='h', marker_color='#64748b'))
    fig_exec.update_layout(barmode='stack', paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
        font=dict(color='#94a3b8', family='Outfit'), height=350,
        xaxis=dict(showgrid=True, gridcolor='rgba(148,163,184,0.1)'), yaxis=dict(showgrid=False),
        legend=dict(orientation="h", y=1.1), margin=dict(l=120))
    
    # HTML
    html = f"""
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Relatório PSU Oracle 2025 - Oraex Premium</title>
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link href="https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;500;600;700;800&display=swap" rel="stylesheet">
    <script src="https://cdn.plot.ly/plotly-2.27.0.min.js"></script>
    <style>
        :root {{
            --bg-primary: #0a0a0f;
            --bg-secondary: #12121a;
            --bg-card: rgba(255,255,255,0.03);
            --border: rgba(255,255,255,0.06);
            --text-primary: #f1f5f9;
            --text-secondary: #94a3b8;
            --text-muted: #64748b;
            --accent-red: #ef4444;
            --accent-green: #10b981;
            --accent-blue: #3b82f6;
            --accent-purple: #8b5cf6;
            --accent-amber: #f59e0b;
            --glow-red: rgba(239,68,68,0.4);
        }}
        
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        
        body {{
            font-family: 'Outfit', sans-serif;
            background: var(--bg-primary);
            color: var(--text-primary);
            min-height: 100vh;
            overflow-x: hidden;
        }}
        
        /* Animated Background */
        .bg-mesh {{
            position: fixed;
            top: 0; left: 0; right: 0; bottom: 0;
            background: 
                radial-gradient(ellipse 80% 50% at 20% 40%, rgba(120,0,255,0.15), transparent),
                radial-gradient(ellipse 60% 40% at 80% 20%, rgba(255,0,80,0.1), transparent),
                radial-gradient(ellipse 50% 30% at 50% 80%, rgba(0,200,255,0.08), transparent);
            pointer-events: none;
            z-index: 0;
        }}
        
        .container {{
            position: relative;
            z-index: 1;
            max-width: 1400px;
            margin: 0 auto;
            padding: 60px 40px;
        }}
        
        /* Hero Header */
        .hero {{
            text-align: center;
            margin-bottom: 80px;
            animation: fadeInUp 0.8s ease-out;
        }}
        
        @keyframes fadeInUp {{
            from {{ opacity: 0; transform: translateY(30px); }}
            to {{ opacity: 1; transform: translateY(0); }}
        }}
        
        @keyframes countUp {{
            from {{ opacity: 0; transform: scale(0.5); }}
            to {{ opacity: 1; transform: scale(1); }}
        }}
        
        @keyframes shimmer {{
            0% {{ background-position: -200% center; }}
            100% {{ background-position: 200% center; }}
        }}
        
        @keyframes pulse {{
            0%, 100% {{ opacity: 1; }}
            50% {{ opacity: 0.7; }}
        }}
        
        @keyframes borderGlow {{
            0%, 100% {{ border-color: rgba(239,68,68,0.3); }}
            50% {{ border-color: rgba(239,68,68,0.6); }}
        }}
        
        .logo {{
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 16px;
            margin-bottom: 30px;
        }}
        
        .logo-icon {{
            width: 70px; height: 70px;
            background: linear-gradient(135deg, #ef4444 0%, #b91c1c 100%);
            border-radius: 20px;
            display: flex; align-items: center; justify-content: center;
            font-size: 2rem; font-weight: 800; color: white;
            box-shadow: 0 20px 40px -10px var(--glow-red);
            animation: pulse 3s ease-in-out infinite;
        }}
        
        .logo-text {{
            font-size: 2.5rem; font-weight: 800;
            background: linear-gradient(90deg, #fff 0%, #94a3b8 100%);
            -webkit-background-clip: text; -webkit-text-fill-color: transparent;
        }}
        
        .hero h1 {{
            font-size: 3.5rem; font-weight: 800; letter-spacing: -0.03em;
            background: linear-gradient(90deg, #fff 0%, #64748b 50%, #fff 100%);
            background-size: 200% auto;
            -webkit-background-clip: text; -webkit-text-fill-color: transparent;
            animation: shimmer 3s linear infinite;
            margin-bottom: 16px;
        }}
        
        .hero .subtitle {{
            font-size: 1.25rem; color: var(--text-secondary); font-weight: 400;
        }}
        
        .hero .period {{
            display: inline-flex; align-items: center; gap: 8px;
            margin-top: 20px; padding: 10px 20px;
            background: var(--bg-card); border: 1px solid var(--border);
            border-radius: 100px; font-size: 0.9rem; color: var(--text-muted);
        }}
        
        .hero .period .dot {{
            width: 8px; height: 8px; background: var(--accent-green);
            border-radius: 50%; animation: pulse 2s ease-in-out infinite;
        }}
        
        /* KPI Grid */
        .kpi-grid {{
            display: grid;
            grid-template-columns: repeat(5, 1fr);
            gap: 24px;
            margin-bottom: 80px;
        }}
        
        .kpi-card {{
            background: var(--bg-card);
            backdrop-filter: blur(20px);
            border: 1px solid var(--border);
            border-radius: 24px;
            padding: 32px;
            position: relative;
            overflow: hidden;
            transition: all 0.4s cubic-bezier(0.4, 0, 0.2, 1);
            animation: fadeInUp 0.8s ease-out backwards;
        }}
        
        .kpi-card:nth-child(1) {{ animation-delay: 0.1s; }}
        .kpi-card:nth-child(2) {{ animation-delay: 0.2s; }}
        .kpi-card:nth-child(3) {{ animation-delay: 0.3s; }}
        .kpi-card:nth-child(4) {{ animation-delay: 0.4s; }}
        .kpi-card:nth-child(5) {{ animation-delay: 0.5s; }}
        
        .kpi-card:hover {{
            transform: translateY(-8px);
            border-color: rgba(255,255,255,0.15);
            box-shadow: 0 30px 60px -20px rgba(0,0,0,0.5);
        }}
        
        .kpi-card::before {{
            content: '';
            position: absolute;
            top: 0; left: 0; right: 0;
            height: 3px;
            background: var(--accent);
            border-radius: 24px 24px 0 0;
        }}
        
        .kpi-card.red::before {{ background: linear-gradient(90deg, #ef4444, #f97316); }}
        .kpi-card.green::before {{ background: linear-gradient(90deg, #10b981, #34d399); }}
        .kpi-card.blue::before {{ background: linear-gradient(90deg, #3b82f6, #60a5fa); }}
        .kpi-card.purple::before {{ background: linear-gradient(90deg, #8b5cf6, #a78bfa); }}
        .kpi-card.amber::before {{ background: linear-gradient(90deg, #f59e0b, #fbbf24); }}
        
        .kpi-icon {{
            width: 48px; height: 48px;
            border-radius: 14px;
            display: flex; align-items: center; justify-content: center;
            font-size: 1.5rem;
            margin-bottom: 20px;
        }}
        
        .kpi-card.red .kpi-icon {{ background: rgba(239,68,68,0.15); }}
        .kpi-card.green .kpi-icon {{ background: rgba(16,185,129,0.15); }}
        .kpi-card.blue .kpi-icon {{ background: rgba(59,130,246,0.15); }}
        .kpi-card.purple .kpi-icon {{ background: rgba(139,92,246,0.15); }}
        .kpi-card.amber .kpi-icon {{ background: rgba(245,158,11,0.15); }}
        
        .kpi-value {{
            font-size: 2.8rem; font-weight: 700; color: var(--text-primary);
            line-height: 1; margin-bottom: 8px;
            animation: countUp 1s ease-out backwards;
        }}
        
        .kpi-card:nth-child(1) .kpi-value {{ animation-delay: 0.3s; }}
        .kpi-card:nth-child(2) .kpi-value {{ animation-delay: 0.4s; }}
        .kpi-card:nth-child(3) .kpi-value {{ animation-delay: 0.5s; }}
        .kpi-card:nth-child(4) .kpi-value {{ animation-delay: 0.6s; }}
        .kpi-card:nth-child(5) .kpi-value {{ animation-delay: 0.7s; }}
        
        .kpi-label {{
            font-size: 0.85rem; color: var(--text-muted);
            text-transform: uppercase; letter-spacing: 0.08em; font-weight: 500;
        }}
        
        .kpi-sub {{
            margin-top: 12px; padding-top: 12px;
            border-top: 1px solid var(--border);
            font-size: 0.8rem; color: var(--text-muted);
        }}
        
        /* Sections */
        .section {{
            margin-bottom: 80px;
            animation: fadeInUp 0.8s ease-out backwards;
        }}
        
        .section-header {{
            display: flex; align-items: center; gap: 16px;
            margin-bottom: 32px;
        }}
        
        .section-number {{
            width: 40px; height: 40px;
            background: linear-gradient(135deg, #ef4444, #b91c1c);
            border-radius: 12px;
            display: flex; align-items: center; justify-content: center;
            font-size: 1rem; font-weight: 700;
        }}
        
        .section-title {{
            font-size: 1.5rem; font-weight: 700; color: var(--text-primary);
        }}
        
        /* Cards */
        .card {{
            background: var(--bg-card);
            backdrop-filter: blur(20px);
            border: 1px solid var(--border);
            border-radius: 24px;
            padding: 32px;
            transition: all 0.3s ease;
        }}
        
        .card:hover {{
            border-color: rgba(255,255,255,0.1);
        }}
        
        .card-title {{
            font-size: 1rem; font-weight: 600; color: var(--text-secondary);
            margin-bottom: 24px;
        }}
        
        .two-col {{ display: grid; grid-template-columns: 1fr 1fr; gap: 32px; }}
        .three-col {{ display: grid; grid-template-columns: repeat(3, 1fr); gap: 24px; }}
        
        /* Entorno Cards */
        .entorno-card {{
            background: var(--bg-card);
            border: 1px solid var(--border);
            border-radius: 20px;
            padding: 28px;
            text-align: center;
            transition: all 0.3s ease;
        }}
        
        .entorno-card:hover {{
            transform: translateY(-4px);
            border-color: rgba(255,255,255,0.1);
        }}
        
        .entorno-icon {{
            width: 60px; height: 60px;
            margin: 0 auto 16px;
            border-radius: 16px;
            display: flex; align-items: center; justify-content: center;
            font-size: 1.5rem; font-weight: 700;
        }}
        
        .entorno-card.prod .entorno-icon {{ background: linear-gradient(135deg, rgba(239,68,68,0.2), rgba(239,68,68,0.1)); color: #ef4444; }}
        .entorno-card.hml .entorno-icon {{ background: linear-gradient(135deg, rgba(245,158,11,0.2), rgba(245,158,11,0.1)); color: #f59e0b; }}
        .entorno-card.dev .entorno-icon {{ background: linear-gradient(135deg, rgba(16,185,129,0.2), rgba(16,185,129,0.1)); color: #10b981; }}
        
        .entorno-value {{ font-size: 2.2rem; font-weight: 700; color: var(--text-primary); }}
        .entorno-label {{ font-size: 0.9rem; color: var(--text-muted); margin-top: 4px; }}
        
        /* Table */
        .table-wrapper {{ overflow-x: auto; }}
        
        table {{
            width: 100%;
            border-collapse: separate;
            border-spacing: 0;
        }}
        
        th {{
            padding: 16px 20px;
            text-align: left;
            font-size: 0.75rem;
            text-transform: uppercase;
            letter-spacing: 0.08em;
            color: var(--text-muted);
            font-weight: 600;
            border-bottom: 1px solid var(--border);
        }}
        
        td {{
            padding: 20px;
            border-bottom: 1px solid var(--border);
            font-size: 0.95rem;
        }}
        
        tr:hover td {{
            background: rgba(255,255,255,0.02);
        }}
        
        .badge {{
            display: inline-flex;
            align-items: center;
            padding: 6px 14px;
            border-radius: 100px;
            font-size: 0.8rem;
            font-weight: 600;
        }}
        
        .badge.green {{ background: rgba(16,185,129,0.15); color: #10b981; }}
        .badge.red {{ background: rgba(239,68,68,0.15); color: #ef4444; }}
        .badge.amber {{ background: rgba(245,158,11,0.15); color: #f59e0b; }}
        
        .progress-bar {{
            height: 8px;
            background: rgba(255,255,255,0.1);
            border-radius: 100px;
            overflow: hidden;
            width: 100px;
        }}
        
        .progress-fill {{
            height: 100%;
            background: linear-gradient(90deg, #10b981, #34d399);
            border-radius: 100px;
            transition: width 1s ease-out;
        }}
        
        /* Footer */
        .footer {{
            text-align: center;
            padding-top: 60px;
            border-top: 1px solid var(--border);
            color: var(--text-muted);
            font-size: 0.9rem;
        }}
        
        .footer-logo {{
            display: inline-flex;
            align-items: center;
            gap: 10px;
            margin-bottom: 16px;
        }}
        
        .footer-logo .icon {{
            width: 36px; height: 36px;
            background: linear-gradient(135deg, #ef4444, #b91c1c);
            border-radius: 10px;
            display: flex; align-items: center; justify-content: center;
            font-weight: 700; font-size: 1rem;
        }}
        
        @media (max-width: 1200px) {{
            .kpi-grid {{ grid-template-columns: repeat(3, 1fr); }}
        }}
        
        @media (max-width: 768px) {{
            .kpi-grid {{ grid-template-columns: 1fr; }}
            .two-col, .three-col {{ grid-template-columns: 1fr; }}
            .hero h1 {{ font-size: 2rem; }}
        }}
        
        @media print {{
            body {{ background: #0a0a0f !important; -webkit-print-color-adjust: exact; }}
            .bg-mesh {{ display: none; }}
        }}
    </style>
</head>
<body>
    <div class="bg-mesh"></div>
    
    <div class="container">
        <!-- Hero -->
        <header class="hero">
            <div class="logo">
                <div class="logo-icon">O</div>
                <span class="logo-text">ORAEX</span>
            </div>
            <h1>Relatório PSU Oracle 2025</h1>
            <p class="subtitle">Consolidação Anual de Atualizações - GetNet Infrastructure</p>
            <div class="period">
                <span class="dot"></span>
                <span>Fevereiro - Dezembro 2025</span>
                <span style="margin-left: 16px; color: var(--text-secondary);">Gerado em {datetime.now().strftime('%d/%m/%Y')}</span>
            </div>
        </header>
        
        <!-- KPIs -->
        <div class="kpi-grid">
            <div class="kpi-card red">
                <div class="kpi-icon">📊</div>
                <div class="kpi-value">{total_psu:,}</div>
                <div class="kpi-label">GMUDs PSU</div>
                <div class="kpi-sub">Total de mudanças Oracle</div>
            </div>
            <div class="kpi-card green">
                <div class="kpi-icon">✅</div>
                <div class="kpi-value">{sucesso_psu:,}</div>
                <div class="kpi-label">Sucesso</div>
                <div class="kpi-sub">{taxa_sucesso:.1f}% de taxa</div>
            </div>
            <div class="kpi-card purple">
                <div class="kpi-icon">🖥️</div>
                <div class="kpi-value">{unique_servers:,}</div>
                <div class="kpi-label">Servidores</div>
                <div class="kpi-sub">Hosts únicos atualizados</div>
            </div>
            <div class="kpi-card blue">
                <div class="kpi-icon">🔄</div>
                <div class="kpi-value">{total_updates:,}</div>
                <div class="kpi-label">Atualizações</div>
                <div class="kpi-sub">Total de intervenções</div>
            </div>
            <div class="kpi-card amber">
                <div class="kpi-icon">⏱️</div>
                <div class="kpi-value">{horas_totais:,}h</div>
                <div class="kpi-label">Esforço</div>
                <div class="kpi-sub">Horas estimadas (3h/srv)</div>
            </div>
        </div>
        
        <!-- Entornos -->
        <section class="section">
            <div class="section-header">
                <div class="section-number">01</div>
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
                <div class="section-number">02</div>
                <h2 class="section-title">Versões PSU Aplicadas</h2>
            </div>
            <div class="card">
                <div class="table-wrapper">
                    <table>
                        <thead>
                            <tr>
                                <th>Versão PSU</th>
                                <th>Total GMUDs</th>
                                <th>Concluídas</th>
                                <th>Taxa de Sucesso</th>
                            </tr>
                        </thead>
                        <tbody>
"""
    
    for _, row in versao_stats.iterrows():
        taxa = row['Sucesso'] / row['Total'] * 100 if row['Total'] > 0 else 0
        html += f"""
                            <tr>
                                <td><strong style="font-size: 1.1rem;">{row['Versao']}</strong></td>
                                <td>{int(row['Total'])}</td>
                                <td><span class="badge green">{int(row['Sucesso'])}</span></td>
                                <td>
                                    <div style="display: flex; align-items: center; gap: 12px;">
                                        <div class="progress-bar"><div class="progress-fill" style="width: {taxa}%;"></div></div>
                                        <span style="color: var(--text-secondary);">{taxa:.0f}%</span>
                                    </div>
                                </td>
                            </tr>
"""
    
    html += f"""
                        </tbody>
                    </table>
                </div>
            </div>
        </section>
        
        <!-- Evolução Mensal -->
        <section class="section">
            <div class="section-header">
                <div class="section-number">03</div>
                <h2 class="section-title">Evolução Mensal</h2>
            </div>
            <div class="card">
                <div id="chart-monthly"></div>
            </div>
        </section>
        
        <!-- Executores -->
        <section class="section">
            <div class="section-header">
                <div class="section-number">04</div>
                <h2 class="section-title">Performance da Equipe</h2>
            </div>
            <div class="two-col">
                <div class="card">
                    <div class="card-title">GMUDs por Executor</div>
                    <div id="chart-executor"></div>
                </div>
                <div class="card">
                    <div class="card-title">Ranking Detalhado</div>
                    <div class="table-wrapper">
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
                                    <td><span class="badge {'green' if row['Taxa'] >= 50 else 'amber'}">{row['Taxa']:.0f}%</span></td>
                                    <td style="color: var(--text-muted);">{int(row['Horas'])}h</td>
                                </tr>
"""
    
    html += f"""
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>
        </section>
        
        <!-- Footer -->
        <footer class="footer">
            <div class="footer-logo">
                <div class="icon">O</div>
                <span style="font-weight: 600; color: var(--text-secondary);">ORAEX Consulting</span>
            </div>
            <p>Relatório Confidencial | GetNet Infrastructure Operations</p>
            <p style="margin-top: 8px; font-size: 0.8rem;">© 2025 Todos os direitos reservados</p>
        </footer>
    </div>
    
    <script>
        // Charts
        Plotly.newPlot('chart-monthly', {fig_monthly.to_json()}.data, {fig_monthly.to_json()}.layout, {{responsive: true}});
        Plotly.newPlot('chart-executor', {fig_exec.to_json()}.data, {fig_exec.to_json()}.layout, {{responsive: true}});
        
        // Counter Animation (simple)
        document.querySelectorAll('.kpi-value').forEach(el => {{
            const text = el.textContent;
            const num = parseInt(text.replace(/[^0-9]/g, ''));
            if (!isNaN(num) && num > 0) {{
                let current = 0;
                const step = Math.ceil(num / 30);
                const suffix = text.includes('%') ? '%' : (text.includes('h') ? 'h' : '');
                const interval = setInterval(() => {{
                    current += step;
                    if (current >= num) {{
                        current = num;
                        clearInterval(interval);
                    }}
                    el.textContent = current.toLocaleString('pt-BR') + suffix;
                }}, 30);
            }}
        }});
    </script>
</body>
</html>
"""
    
    with open(OUTPUT_HTML, 'w', encoding='utf-8') as f:
        f.write(html)
    
    print(f"\n✅ Relatório V3 Premium gerado: {OUTPUT_HTML}")

if __name__ == "__main__":
    print("="*60)
    print("GERANDO RELATÓRIO PSU 2025 - V3 ULTRA PREMIUM")
    print("="*60)
    df = load_all_gmuds()
    if not df.empty:
        generate_premium_report(df)
