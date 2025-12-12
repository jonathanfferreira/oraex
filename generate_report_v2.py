import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import re

FILE_PATH = r"D:\antigravity\oraex\cmdb\ORAEX - Consolidação GetTech 2025 (1).xlsm"
OUTPUT_HTML = r"D:\antigravity\oraex\cmdb\relatorio_psu_2025_v2.html"

MONTHLY_SHEETS = [
    'FEVEREIRO-25', 'MARÇO-25', 'ABRIL-25', 'MAIO-25', 'JUNHO-25',
    'JULHO-25', 'AGOSTO-25', 'SETEMBRO-25', 'OUTUBRO-25', 'NOVEMBRO-25', 'DEZEMBRO-25'
]

MONTH_ORDER = ['FEVEREIRO', 'MARÇO', 'ABRIL', 'MAIO', 'JUNHO', 'JULHO', 
               'AGOSTO', 'SETEMBRO', 'OUTUBRO', 'NOVEMBRO', 'DEZEMBRO']

ENTORNO_MAP = {
    'P': 'Produção',
    'H': 'Homologação', 
    'D': 'Desenvolvimento',
    'T': 'Transacional'
}

def extract_hostnames(title):
    if pd.isna(title):
        return []
    pattern = r'(gncas[a-z0-9]+)'
    matches = re.findall(pattern, str(title).lower())
    return [m.upper() for m in matches]

def extract_psu_version(title):
    """Extrai versão PSU do título (ex: 19.25, 19.27, 19.28)"""
    if pd.isna(title):
        return None
    # Padrões: PSU 19.25, PSU 19.27, 19c, etc
    pattern = r'PSU\s*(19[.\d]+)'
    match = re.search(pattern, str(title), re.IGNORECASE)
    if match:
        return match.group(1)
    # Tentar padrão alternativo
    pattern2 = r'19\.(\d+)'
    match2 = re.search(pattern2, str(title))
    if match2:
        return f"19.{match2.group(1)}"
    return None

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
    elif 'ANDAMENTO' in status or 'EXECUÇÃO' in status or 'IMPLEMENT' in status:
        return 'EM ANDAMENTO'
    elif 'PROGRAM' in status or 'NOVO' in status or 'AUTORIZAR' in status or 'CAB' in status or 'AVALIAR' in status:
        return 'PENDENTE'
    else:
        return 'OUTROS'

def normalize_responsavel(resp):
    """Normaliza nomes de responsáveis"""
    if pd.isna(resp):
        return 'Não Atribuído'
    resp = str(resp).strip().title()
    # Unificar variações
    if 'Guilherme' in resp:
        return 'Guilherme Fonseca'
    if 'Bruno' in resp:
        return 'Bruno Ferreira'
    if 'Alcides' in resp:
        return 'Alcides Souto'
    if 'Kaue' in resp:
        return 'Kaue Santos'
    if 'Rafael' in resp:
        return 'Rafael Rabello'
    if 'Luca' in resp:
        return 'Luca Mozart'
    if 'Jonathan' in resp:
        return 'Jonathan Ferreira'
    return resp

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
        except Exception as e:
            print(f"Erro em {sheet}: {e}")
    
    return pd.concat(all_data, ignore_index=True) if all_data else pd.DataFrame()

def generate_report_v2(df):
    print("Processando dados...")
    
    # Enriquecer dados
    df['Status_Final'] = df['Status'].apply(normalize_status)
    df['Hostnames'] = df['Titulo'].apply(extract_hostnames)
    df['Num_Servers'] = df['Hostnames'].apply(len)
    df['Versao_PSU'] = df['Titulo'].apply(extract_psu_version)
    df['Responsavel_Norm'] = df['Responsavel'].apply(normalize_responsavel)
    df['Entorno_Nome'] = df['Entorno'].map(ENTORNO_MAP).fillna('Outros')
    df['Is_PSU'] = df['Titulo'].str.contains('PSU', case=False, na=False)
    
    # Filtrar apenas GMUDs de PSU
    df_psu = df[df['Is_PSU']].copy()
    
    # =============== MÉTRICAS GERAIS ===============
    total_gmuds = len(df)
    total_psu = len(df_psu)
    
    sucesso_total = len(df[df['Status_Final'] == 'SUCESSO'])
    sucesso_psu = len(df_psu[df_psu['Status_Final'] == 'SUCESSO'])
    
    # Servidores
    all_hosts = []
    for h in df_psu[df_psu['Status_Final'] == 'SUCESSO']['Hostnames']:
        all_hosts.extend(h)
    unique_servers_success = len(set(all_hosts))
    total_updates_success = len(all_hosts)
    
    # Horas trabalhadas (3h por servidor com sucesso)
    horas_totais = total_updates_success * 3
    
    # =============== POR VERSÃO PSU ===============
    versao_stats = df_psu.groupby('Versao_PSU').agg({
        'GMUD_ID': 'count',
        'Status_Final': lambda x: (x == 'SUCESSO').sum()
    }).reset_index()
    versao_stats.columns = ['Versao', 'Total', 'Sucesso']
    versao_stats = versao_stats.dropna(subset=['Versao'])
    versao_stats = versao_stats.sort_values('Versao')
    
    # =============== POR ENTORNO ===============
    entorno_stats = df_psu.groupby(['Entorno', 'Status_Final']).size().unstack(fill_value=0)
    
    # =============== POR EXECUTOR ===============
    executor_stats = df_psu.groupby('Responsavel_Norm').agg({
        'GMUD_ID': 'count',
        'Status_Final': lambda x: (x == 'SUCESSO').sum(),
        'Num_Servers': 'sum'
    }).reset_index()
    executor_stats.columns = ['Executor', 'Total_GMUDs', 'Sucesso', 'Servidores']
    executor_stats['Insucesso'] = executor_stats['Total_GMUDs'] - executor_stats['Sucesso']
    executor_stats['Taxa_Sucesso'] = (executor_stats['Sucesso'] / executor_stats['Total_GMUDs'] * 100).round(1)
    executor_stats['Horas_Estimadas'] = executor_stats['Servidores'] * 3
    executor_stats = executor_stats.sort_values('Total_GMUDs', ascending=False)
    executor_stats = executor_stats[executor_stats['Executor'] != 'Não Atribuído'].head(10)
    
    # =============== POR MÊS ===============
    monthly_psu = df_psu.groupby(['Mes', 'Status_Final']).size().unstack(fill_value=0)
    monthly_psu = monthly_psu.reindex(MONTH_ORDER)
    
    # =============== GRÁFICOS ===============
    # Chart 1: Versões PSU (Treemap)
    fig_versao = px.treemap(
        versao_stats,
        path=['Versao'],
        values='Total',
        color='Sucesso',
        color_continuous_scale=['#7f1d1d', '#22c55e']
    )
    fig_versao.update_layout(
        paper_bgcolor='rgba(0,0,0,0)', font=dict(color='white'), height=300,
        margin=dict(t=10, l=10, r=10, b=10)
    )
    fig_versao.update_coloraxes(showscale=False)
    
    # Chart 2: Entorno (Donut)
    entorno_totals = df_psu['Entorno'].value_counts()
    fig_entorno = go.Figure(data=[go.Pie(
        labels=[ENTORNO_MAP.get(e, e) for e in entorno_totals.index],
        values=entorno_totals.values,
        hole=0.6,
        marker_colors=['#ef4444', '#f59e0b', '#22c55e', '#3b82f6']
    )])
    fig_entorno.update_layout(
        paper_bgcolor='rgba(0,0,0,0)', font=dict(color='white'), height=300,
        legend=dict(orientation="h", y=-0.1)
    )
    
    # Chart 3: Evolução Mensal
    fig_monthly = go.Figure()
    colors = {'SUCESSO': '#22c55e', 'REPLANEJADA': '#f59e0b', 'CANCELADA': '#ef4444', 'INSUCESSO': '#dc2626'}
    for status in ['SUCESSO', 'REPLANEJADA', 'CANCELADA']:
        if status in monthly_psu.columns:
            fig_monthly.add_trace(go.Bar(
                name=status, x=monthly_psu.index, y=monthly_psu[status],
                marker_color=colors.get(status, '#888')
            ))
    fig_monthly.update_layout(
        barmode='stack',
        paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
        font=dict(color='white'), height=350,
        xaxis=dict(showgrid=False), yaxis=dict(showgrid=True, gridcolor='#333'),
        legend=dict(orientation="h", y=1.1)
    )
    
    # Chart 4: Top Executores (Horizontal Bar)
    fig_exec = go.Figure()
    fig_exec.add_trace(go.Bar(
        y=executor_stats['Executor'].head(8),
        x=executor_stats['Sucesso'].head(8),
        name='Sucesso',
        orientation='h',
        marker_color='#22c55e'
    ))
    fig_exec.add_trace(go.Bar(
        y=executor_stats['Executor'].head(8),
        x=executor_stats['Insucesso'].head(8),
        name='Insucesso/Outros',
        orientation='h',
        marker_color='#ef4444'
    ))
    fig_exec.update_layout(
        barmode='stack',
        paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
        font=dict(color='white'), height=350,
        xaxis=dict(showgrid=True, gridcolor='#333'),
        yaxis=dict(showgrid=False),
        legend=dict(orientation="h", y=1.1)
    )
    
    # =============== HTML ===============
    html = f"""
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Relatório PSU Oracle 2025 - Oraex (Detalhado)</title>
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        body {{
            font-family: 'Inter', sans-serif;
            background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%);
            color: #e2e8f0;
            min-height: 100vh;
            padding: 40px;
        }}
        .container {{ max-width: 1400px; margin: 0 auto; }}
        .header {{
            text-align: center;
            margin-bottom: 50px;
            padding-bottom: 30px;
            border-bottom: 2px solid #334155;
        }}
        .header h1 {{ font-size: 2.5rem; font-weight: 800; color: #fff; margin-bottom: 10px; }}
        .header .subtitle {{ font-size: 1.2rem; color: #94a3b8; }}
        .header .date {{ font-size: 0.9rem; color: #64748b; margin-top: 10px; }}
        .logo {{
            display: flex; align-items: center; justify-content: center;
            gap: 15px; margin-bottom: 20px;
        }}
        .logo-circle {{
            width: 60px; height: 60px;
            background: linear-gradient(135deg, #ec0000 0%, #b30000 100%);
            border-radius: 50%;
            display: flex; align-items: center; justify-content: center;
            font-weight: bold; font-size: 1.5rem; color: white;
        }}
        .kpi-grid {{
            display: grid;
            grid-template-columns: repeat(5, 1fr);
            gap: 20px;
            margin-bottom: 50px;
        }}
        .kpi-card {{
            background: rgba(255,255,255,0.05);
            border: 1px solid rgba(255,255,255,0.1);
            border-radius: 16px;
            padding: 24px;
            text-align: center;
            position: relative;
        }}
        .kpi-card.red {{ border-left: 4px solid #ec0000; }}
        .kpi-card.green {{ border-left: 4px solid #22c55e; }}
        .kpi-card.blue {{ border-left: 4px solid #3b82f6; }}
        .kpi-card.purple {{ border-left: 4px solid #8b5cf6; }}
        .kpi-card.yellow {{ border-left: 4px solid #f59e0b; }}
        .kpi-value {{ font-size: 2.2rem; font-weight: 700; color: #fff; }}
        .kpi-label {{ font-size: 0.8rem; color: #94a3b8; text-transform: uppercase; letter-spacing: 0.05em; margin-top: 5px; }}
        .kpi-sub {{ font-size: 0.7rem; color: #64748b; margin-top: 8px; }}
        .section {{ margin-bottom: 50px; }}
        .section-title {{
            font-size: 1.4rem; font-weight: 700; color: #fff;
            margin-bottom: 25px; display: flex; align-items: center; gap: 12px;
        }}
        .section-title::before {{
            content: ''; width: 4px; height: 28px; background: #ec0000; border-radius: 2px;
        }}
        .chart-container {{
            background: rgba(255,255,255,0.03);
            border: 1px solid rgba(255,255,255,0.05);
            border-radius: 16px;
            padding: 25px;
        }}
        .two-col {{ display: grid; grid-template-columns: 1fr 1fr; gap: 30px; }}
        .three-col {{ display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 30px; }}
        .table-container {{ overflow-x: auto; }}
        table {{ width: 100%; border-collapse: collapse; }}
        th, td {{ padding: 14px 16px; text-align: left; border-bottom: 1px solid #334155; }}
        th {{ background: rgba(255,255,255,0.05); font-weight: 600; color: #cbd5e1; font-size: 0.85rem; text-transform: uppercase; }}
        tr:hover {{ background: rgba(255,255,255,0.03); }}
        .badge {{
            display: inline-block; padding: 4px 12px; border-radius: 20px;
            font-size: 0.75rem; font-weight: 600;
        }}
        .badge.green {{ background: rgba(34,197,94,0.2); color: #22c55e; }}
        .badge.yellow {{ background: rgba(245,158,11,0.2); color: #f59e0b; }}
        .badge.red {{ background: rgba(239,68,68,0.2); color: #ef4444; }}
        .progress-bar {{
            height: 8px; background: #334155; border-radius: 4px; overflow: hidden;
        }}
        .progress-fill {{ height: 100%; background: #22c55e; border-radius: 4px; }}
        .footer {{
            text-align: center; padding-top: 30px; border-top: 1px solid #334155;
            color: #64748b; font-size: 0.85rem;
        }}
        .highlight-box {{
            background: linear-gradient(135deg, rgba(236,0,0,0.1) 0%, rgba(236,0,0,0.05) 100%);
            border: 1px solid rgba(236,0,0,0.3);
            border-radius: 12px;
            padding: 20px;
            margin-bottom: 30px;
        }}
        .highlight-box h3 {{ color: #ec0000; margin-bottom: 10px; }}
        @media print {{
            body {{ background: #0f172a; -webkit-print-color-adjust: exact; }}
        }}
    </style>
    <script src="https://cdn.plot.ly/plotly-2.27.0.min.js"></script>
</head>
<body>
    <div class="container">
        <div class="header">
            <div class="logo">
                <div class="logo-circle">O</div>
                <span style="font-size: 2rem; font-weight: 700;">ORAEX</span>
            </div>
            <h1>Relatório de Atualizações PSU Oracle 2025</h1>
            <div class="subtitle">Consolidação Anual Detalhada - GetNet Infrastructure</div>
            <div class="date">Período: Fevereiro a Dezembro 2025 | Gerado em: {datetime.now().strftime('%d/%m/%Y às %H:%M')}</div>
        </div>
        
        <!-- RESUMO EXECUTIVO -->
        <div class="highlight-box">
            <h3>📋 Resumo Executivo</h3>
            <p>Durante o ano de 2025, a equipe Oraex executou <strong>{total_psu:,} GMUDs de atualização PSU Oracle</strong>, 
            resultando em <strong>{unique_servers_success:,} servidores únicos atualizados com sucesso</strong>. 
            O esforço total estimado foi de <strong>{horas_totais:,} horas de trabalho técnico</strong>.</p>
        </div>
        
        <!-- KPIs PRINCIPAIS -->
        <div class="kpi-grid">
            <div class="kpi-card red">
                <div class="kpi-value">{total_psu:,}</div>
                <div class="kpi-label">GMUDs PSU</div>
                <div class="kpi-sub">Total de mudanças</div>
            </div>
            <div class="kpi-card green">
                <div class="kpi-value">{sucesso_psu:,}</div>
                <div class="kpi-label">Sucesso</div>
                <div class="kpi-sub">{sucesso_psu/total_psu*100:.1f}% do total</div>
            </div>
            <div class="kpi-card purple">
                <div class="kpi-value">{unique_servers_success:,}</div>
                <div class="kpi-label">Servidores Únicos</div>
                <div class="kpi-sub">Atualizados com sucesso</div>
            </div>
            <div class="kpi-card blue">
                <div class="kpi-value">{total_updates_success:,}</div>
                <div class="kpi-label">Atualizações</div>
                <div class="kpi-sub">Total de intervenções</div>
            </div>
            <div class="kpi-card yellow">
                <div class="kpi-value">{horas_totais:,}h</div>
                <div class="kpi-label">Horas Trabalhadas</div>
                <div class="kpi-sub">Estimativa (3h/servidor)</div>
            </div>
        </div>
        
        <!-- SEÇÃO: VERSÕES PSU -->
        <div class="section">
            <div class="section-title">Atualizações por Versão PSU (Quarter Oracle)</div>
            <div class="two-col">
                <div class="chart-container">
                    <h4 style="margin-bottom: 15px; color: #cbd5e1;">Distribuição de GMUDs por Versão</h4>
                    <div id="chart-versao"></div>
                </div>
                <div class="chart-container table-container">
                    <h4 style="margin-bottom: 15px; color: #cbd5e1;">Detalhamento por Versão</h4>
                    <table>
                        <thead>
                            <tr><th>Versão PSU</th><th>Total GMUDs</th><th>Sucesso</th><th>Taxa</th></tr>
                        </thead>
                        <tbody>
"""
    
    for _, row in versao_stats.iterrows():
        taxa = row['Sucesso'] / row['Total'] * 100 if row['Total'] > 0 else 0
        html += f"""
                            <tr>
                                <td><strong>{row['Versao']}</strong></td>
                                <td>{int(row['Total'])}</td>
                                <td><span class="badge green">{int(row['Sucesso'])}</span></td>
                                <td>
                                    <div style="display: flex; align-items: center; gap: 10px;">
                                        <div class="progress-bar" style="width: 100px;">
                                            <div class="progress-fill" style="width: {taxa}%;"></div>
                                        </div>
                                        <span>{taxa:.0f}%</span>
                                    </div>
                                </td>
                            </tr>
"""
    
    html += f"""
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
        
        <!-- SEÇÃO: POR ENTORNO -->
        <div class="section">
            <div class="section-title">Distribuição por Ambiente (Entorno)</div>
            <div class="two-col">
                <div class="chart-container">
                    <h4 style="margin-bottom: 15px; color: #cbd5e1;">GMUDs por Ambiente</h4>
                    <div id="chart-entorno"></div>
                </div>
                <div class="chart-container">
                    <h4 style="margin-bottom: 15px; color: #cbd5e1;">Legenda de Ambientes</h4>
                    <div style="display: grid; gap: 15px; margin-top: 20px;">
                        <div style="display: flex; align-items: center; gap: 15px;">
                            <div style="width: 40px; height: 40px; background: #ef4444; border-radius: 8px; display: flex; align-items: center; justify-content: center; font-weight: bold;">P</div>
                            <div>
                                <div style="font-weight: 600;">Produção</div>
                                <div style="color: #94a3b8; font-size: 0.85rem;">{len(df_psu[df_psu['Entorno'] == 'P'])} GMUDs - Ambiente crítico/transacional</div>
                            </div>
                        </div>
                        <div style="display: flex; align-items: center; gap: 15px;">
                            <div style="width: 40px; height: 40px; background: #f59e0b; border-radius: 8px; display: flex; align-items: center; justify-content: center; font-weight: bold;">H</div>
                            <div>
                                <div style="font-weight: 600;">Homologação</div>
                                <div style="color: #94a3b8; font-size: 0.85rem;">{len(df_psu[df_psu['Entorno'] == 'H'])} GMUDs - Ambiente de testes</div>
                            </div>
                        </div>
                        <div style="display: flex; align-items: center; gap: 15px;">
                            <div style="width: 40px; height: 40px; background: #22c55e; border-radius: 8px; display: flex; align-items: center; justify-content: center; font-weight: bold;">D</div>
                            <div>
                                <div style="font-weight: 600;">Desenvolvimento</div>
                                <div style="color: #94a3b8; font-size: 0.85rem;">{len(df_psu[df_psu['Entorno'] == 'D'])} GMUDs - Ambiente de dev</div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
        
        <!-- SEÇÃO: EVOLUÇÃO MENSAL -->
        <div class="section">
            <div class="section-title">Evolução Mensal de Execuções</div>
            <div class="chart-container">
                <div id="chart-monthly"></div>
            </div>
        </div>
        
        <!-- SEÇÃO: PERFORMANCE POR EXECUTOR -->
        <div class="section">
            <div class="section-title">Performance por Executor (Top 8)</div>
            <div class="two-col">
                <div class="chart-container">
                    <h4 style="margin-bottom: 15px; color: #cbd5e1;">GMUDs Executadas por Pessoa</h4>
                    <div id="chart-executor"></div>
                </div>
                <div class="chart-container table-container">
                    <h4 style="margin-bottom: 15px; color: #cbd5e1;">Ranking de Executores</h4>
                    <table>
                        <thead>
                            <tr><th>Executor</th><th>GMUDs</th><th>✅</th><th>❌</th><th>Taxa</th><th>Horas</th></tr>
                        </thead>
                        <tbody>
"""
    
    for _, row in executor_stats.head(8).iterrows():
        html += f"""
                            <tr>
                                <td><strong>{row['Executor']}</strong></td>
                                <td>{int(row['Total_GMUDs'])}</td>
                                <td><span class="badge green">{int(row['Sucesso'])}</span></td>
                                <td><span class="badge red">{int(row['Insucesso'])}</span></td>
                                <td>{row['Taxa_Sucesso']:.0f}%</td>
                                <td>{int(row['Horas_Estimadas'])}h</td>
                            </tr>
"""
    
    html += f"""
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
        
        <!-- FOOTER -->
        <div class="footer">
            <p><strong>© 2025 ORAEX Consulting</strong> | Relatório Confidencial</p>
            <p style="margin-top: 5px;">GetNet Infrastructure Operations - Database Team</p>
        </div>
    </div>
    
    <script>
        Plotly.newPlot('chart-versao', {fig_versao.to_json()}.data, {fig_versao.to_json()}.layout, {{responsive: true}});
        Plotly.newPlot('chart-entorno', {fig_entorno.to_json()}.data, {fig_entorno.to_json()}.layout, {{responsive: true}});
        Plotly.newPlot('chart-monthly', {fig_monthly.to_json()}.data, {fig_monthly.to_json()}.layout, {{responsive: true}});
        Plotly.newPlot('chart-executor', {fig_exec.to_json()}.data, {fig_exec.to_json()}.layout, {{responsive: true}});
    </script>
</body>
</html>
"""
    
    with open(OUTPUT_HTML, 'w', encoding='utf-8') as f:
        f.write(html)
    
    print(f"\n✅ Relatório V2 gerado: {OUTPUT_HTML}")
    return df_psu

if __name__ == "__main__":
    print("="*60)
    print("GERANDO RELATÓRIO PSU 2025 - VERSÃO DETALHADA")
    print("="*60)
    df = load_all_gmuds()
    if not df.empty:
        generate_report_v2(df)
