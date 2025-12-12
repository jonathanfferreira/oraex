import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import re

FILE_PATH = r"D:\antigravity\oraex\cmdb\ORAEX - Consolidação GetTech 2025 (1).xlsm"
OUTPUT_HTML = r"D:\antigravity\oraex\cmdb\relatorio_psu_2025.html"

MONTHLY_SHEETS = [
    'FEVEREIRO-25', 'MARÇO-25', 'ABRIL-25', 'MAIO-25', 'JUNHO-25',
    'JULHO-25', 'AGOSTO-25', 'SETEMBRO-25', 'OUTUBRO-25', 'NOVEMBRO-25', 'DEZEMBRO-25'
]

MONTH_ORDER = ['FEVEREIRO', 'MARÇO', 'ABRIL', 'MAIO', 'JUNHO', 'JULHO', 
               'AGOSTO', 'SETEMBRO', 'OUTUBRO', 'NOVEMBRO', 'DEZEMBRO']

def extract_hostnames(title):
    if pd.isna(title):
        return []
    pattern = r'(gncas[a-z0-9]+)'
    matches = re.findall(pattern, str(title).lower())
    return [m.upper() for m in matches]

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
                elif 'data' in col_lower and 'início' in col_lower:
                    col_mapping[col] = 'Data_Inicio'
                elif 'entorno' in col_lower:
                    col_mapping[col] = 'Entorno'
                elif 'designado' in col_lower:
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

def generate_executive_report(df):
    # Enrich data
    df['Status_Final'] = df['Status'].apply(normalize_status)
    df['Hostnames'] = df['Titulo'].apply(extract_hostnames)
    df['Num_Servers'] = df['Hostnames'].apply(len)
    
    # Metrics
    total_gmuds = len(df)
    sucesso = len(df[df['Status_Final'] == 'SUCESSO'])
    canceladas = len(df[df['Status_Final'] == 'CANCELADA'])
    replanejadas = len(df[df['Status_Final'] == 'REPLANEJADA'])
    insucesso = len(df[df['Status_Final'] == 'INSUCESSO'])
    pendentes = len(df[df['Status_Final'] == 'PENDENTE'])
    
    all_hosts = []
    for h in df['Hostnames']:
        all_hosts.extend(h)
    unique_servers = len(set(all_hosts))
    total_updates = len(all_hosts)
    
    success_rate = (sucesso / total_gmuds * 100) if total_gmuds > 0 else 0
    
    # Monthly chart data
    monthly_status = df.groupby(['Mes', 'Status_Final']).size().unstack(fill_value=0)
    monthly_status = monthly_status.reindex(MONTH_ORDER)
    
    # Create charts
    # Chart 1: Status Distribution (Donut)
    status_counts = df['Status_Final'].value_counts()
    colors_map = {
        'SUCESSO': '#22c55e', 'CANCELADA': '#ef4444', 'REPLANEJADA': '#f59e0b',
        'INSUCESSO': '#dc2626', 'PENDENTE': '#6b7280', 'EM ANDAMENTO': '#3b82f6', 'OUTROS': '#9ca3af'
    }
    fig_donut = go.Figure(data=[go.Pie(
        labels=status_counts.index,
        values=status_counts.values,
        hole=0.6,
        marker_colors=[colors_map.get(s, '#888') for s in status_counts.index]
    )])
    fig_donut.update_layout(
        paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
        font=dict(color='white'), showlegend=True, height=350,
        legend=dict(orientation="h", yanchor="bottom", y=-0.2)
    )
    
    # Chart 2: Monthly Evolution
    fig_monthly = go.Figure()
    if 'SUCESSO' in monthly_status.columns:
        fig_monthly.add_trace(go.Bar(name='Sucesso', x=monthly_status.index, y=monthly_status.get('SUCESSO', 0), marker_color='#22c55e'))
    if 'REPLANEJADA' in monthly_status.columns:
        fig_monthly.add_trace(go.Bar(name='Replanejada', x=monthly_status.index, y=monthly_status.get('REPLANEJADA', 0), marker_color='#f59e0b'))
    if 'CANCELADA' in monthly_status.columns:
        fig_monthly.add_trace(go.Bar(name='Cancelada', x=monthly_status.index, y=monthly_status.get('CANCELADA', 0), marker_color='#ef4444'))
    
    fig_monthly.update_layout(
        barmode='stack',
        paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
        font=dict(color='white'), height=350,
        xaxis=dict(showgrid=False), yaxis=dict(showgrid=True, gridcolor='#333'),
        legend=dict(orientation="h", yanchor="bottom", y=1.02)
    )
    
    # Generate HTML
    html = f"""
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Relatório PSU Oracle 2025 - Oraex</title>
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;700&display=swap');
        
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        
        body {{
            font-family: 'Inter', sans-serif;
            background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%);
            color: #e2e8f0;
            min-height: 100vh;
            padding: 40px;
        }}
        
        .container {{ max-width: 1200px; margin: 0 auto; }}
        
        .header {{
            text-align: center;
            margin-bottom: 50px;
            padding-bottom: 30px;
            border-bottom: 2px solid #334155;
        }}
        
        .header h1 {{
            font-size: 2.5rem;
            font-weight: 700;
            color: #ffffff;
            margin-bottom: 10px;
        }}
        
        .header .subtitle {{
            font-size: 1.2rem;
            color: #94a3b8;
        }}
        
        .header .date {{
            font-size: 0.9rem;
            color: #64748b;
            margin-top: 10px;
        }}
        
        .logo {{
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 15px;
            margin-bottom: 20px;
        }}
        
        .logo-circle {{
            width: 60px;
            height: 60px;
            background: linear-gradient(135deg, #ec0000 0%, #b30000 100%);
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            font-weight: bold;
            font-size: 1.5rem;
        }}
        
        .kpi-grid {{
            display: grid;
            grid-template-columns: repeat(4, 1fr);
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
            overflow: hidden;
        }}
        
        .kpi-card.success {{ border-left: 4px solid #22c55e; }}
        .kpi-card.warning {{ border-left: 4px solid #f59e0b; }}
        .kpi-card.danger {{ border-left: 4px solid #ef4444; }}
        .kpi-card.info {{ border-left: 4px solid #3b82f6; }}
        
        .kpi-value {{
            font-size: 2.5rem;
            font-weight: 700;
            color: #ffffff;
            margin-bottom: 5px;
        }}
        
        .kpi-label {{
            font-size: 0.85rem;
            color: #94a3b8;
            text-transform: uppercase;
            letter-spacing: 0.05em;
        }}
        
        .kpi-sub {{
            font-size: 0.75rem;
            color: #64748b;
            margin-top: 8px;
        }}
        
        .section {{
            margin-bottom: 50px;
        }}
        
        .section-title {{
            font-size: 1.5rem;
            font-weight: 600;
            color: #ffffff;
            margin-bottom: 20px;
            display: flex;
            align-items: center;
            gap: 10px;
        }}
        
        .section-title::before {{
            content: '';
            width: 4px;
            height: 24px;
            background: #ec0000;
            border-radius: 2px;
        }}
        
        .chart-container {{
            background: rgba(255,255,255,0.03);
            border: 1px solid rgba(255,255,255,0.05);
            border-radius: 16px;
            padding: 20px;
        }}
        
        .two-col {{
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 30px;
        }}
        
        .summary-table {{
            width: 100%;
            border-collapse: collapse;
            margin-top: 20px;
        }}
        
        .summary-table th, .summary-table td {{
            padding: 12px 16px;
            text-align: left;
            border-bottom: 1px solid #334155;
        }}
        
        .summary-table th {{
            background: rgba(255,255,255,0.05);
            font-weight: 600;
            color: #cbd5e1;
        }}
        
        .summary-table tr:hover {{
            background: rgba(255,255,255,0.03);
        }}
        
        .badge {{
            display: inline-block;
            padding: 4px 12px;
            border-radius: 20px;
            font-size: 0.75rem;
            font-weight: 600;
        }}
        
        .badge.success {{ background: rgba(34,197,94,0.2); color: #22c55e; }}
        .badge.warning {{ background: rgba(245,158,11,0.2); color: #f59e0b; }}
        .badge.danger {{ background: rgba(239,68,68,0.2); color: #ef4444; }}
        
        .footer {{
            text-align: center;
            padding-top: 30px;
            border-top: 1px solid #334155;
            color: #64748b;
            font-size: 0.85rem;
        }}
        
        @media print {{
            body {{ background: #0f172a; -webkit-print-color-adjust: exact; }}
            .container {{ max-width: 100%; }}
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
            <h1>Relatório de Atualizações PSU Oracle</h1>
            <div class="subtitle">Consolidação Anual 2025 - GetNet</div>
            <div class="date">Gerado em: {datetime.now().strftime('%d/%m/%Y às %H:%M')}</div>
        </div>
        
        <div class="kpi-grid">
            <div class="kpi-card info">
                <div class="kpi-value">{total_gmuds:,}</div>
                <div class="kpi-label">Total de GMUDs</div>
                <div class="kpi-sub">Fev - Dez 2025</div>
            </div>
            <div class="kpi-card success">
                <div class="kpi-value">{sucesso:,}</div>
                <div class="kpi-label">Concluídas com Sucesso</div>
                <div class="kpi-sub">{success_rate:.1f}% do total</div>
            </div>
            <div class="kpi-card warning">
                <div class="kpi-value">{replanejadas:,}</div>
                <div class="kpi-label">Replanejadas</div>
                <div class="kpi-sub">{replanejadas/total_gmuds*100:.1f}% do total</div>
            </div>
            <div class="kpi-card danger">
                <div class="kpi-value">{canceladas + insucesso:,}</div>
                <div class="kpi-label">Canceladas/Insucesso</div>
                <div class="kpi-sub">{(canceladas+insucesso)/total_gmuds*100:.1f}% do total</div>
            </div>
        </div>
        
        <div class="kpi-grid" style="grid-template-columns: repeat(2, 1fr);">
            <div class="kpi-card" style="border-left: 4px solid #8b5cf6;">
                <div class="kpi-value">{unique_servers:,}</div>
                <div class="kpi-label">Servidores Únicos Atualizados</div>
                <div class="kpi-sub">Hostnames distintos processados</div>
            </div>
            <div class="kpi-card" style="border-left: 4px solid #06b6d4;">
                <div class="kpi-value">{total_updates:,}</div>
                <div class="kpi-label">Total de Atualizações</div>
                <div class="kpi-sub">Incluindo múltiplas passagens</div>
            </div>
        </div>
        
        <div class="section">
            <div class="section-title">Análise Visual</div>
            <div class="two-col">
                <div class="chart-container">
                    <h3 style="margin-bottom: 15px; color: #cbd5e1;">Distribuição por Status</h3>
                    <div id="chart-donut"></div>
                </div>
                <div class="chart-container">
                    <h3 style="margin-bottom: 15px; color: #cbd5e1;">Evolução Mensal</h3>
                    <div id="chart-monthly"></div>
                </div>
            </div>
        </div>
        
        <div class="section">
            <div class="section-title">Resumo por Mês</div>
            <div class="chart-container">
                <table class="summary-table">
                    <thead>
                        <tr>
                            <th>Mês</th>
                            <th>Total GMUDs</th>
                            <th>Sucesso</th>
                            <th>Replanejadas</th>
                            <th>Canceladas</th>
                        </tr>
                    </thead>
                    <tbody>
"""
    
    # Add monthly rows
    monthly_summary = df.groupby('Mes')['Status_Final'].value_counts().unstack(fill_value=0)
    for mes in MONTH_ORDER:
        if mes in monthly_summary.index:
            row = monthly_summary.loc[mes]
            total_mes = row.sum()
            suc = row.get('SUCESSO', 0)
            rep = row.get('REPLANEJADA', 0)
            can = row.get('CANCELADA', 0) + row.get('INSUCESSO', 0)
            html += f"""
                        <tr>
                            <td>{mes}</td>
                            <td>{total_mes}</td>
                            <td><span class="badge success">{suc}</span></td>
                            <td><span class="badge warning">{rep}</span></td>
                            <td><span class="badge danger">{can}</span></td>
                        </tr>
"""
    
    html += f"""
                    </tbody>
                </table>
            </div>
        </div>
        
        <div class="footer">
            <p>© 2025 ORAEX Consulting | Relatório Confidencial</p>
            <p style="margin-top: 5px;">GetNet Infrastructure Operations</p>
        </div>
    </div>
    
    <script>
        var donutData = {fig_donut.to_json()};
        Plotly.newPlot('chart-donut', donutData.data, donutData.layout, {{responsive: true}});
        
        var monthlyData = {fig_monthly.to_json()};
        Plotly.newPlot('chart-monthly', monthlyData.data, monthlyData.layout, {{responsive: true}});
    </script>
</body>
</html>
"""
    
    with open(OUTPUT_HTML, 'w', encoding='utf-8') as f:
        f.write(html)
    
    print(f"✅ Relatório HTML gerado: {OUTPUT_HTML}")
    return df

if __name__ == "__main__":
    print("Carregando dados...")
    df = load_all_gmuds()
    if not df.empty:
        generate_executive_report(df)
        print("\n🎯 Para converter em PDF, abra o HTML no navegador e use Ctrl+P > Salvar como PDF")
    else:
        print("Erro: Nenhum dado encontrado!")
