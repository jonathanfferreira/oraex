import json
import pandas as pd
from datetime import datetime
import base64

# Carregar dados
with open('data_summary.json', 'r', encoding='utf-8') as f:
    data = json.load(f)

# Garantir tipos numéricos
data['total_registros'] = int(str(data['total_registros']).replace('.', '').replace(',', ''))
data['total_tabelas'] = int(str(data['total_tabelas']))

# Carregar logo (tenta ler do arquivo local se existir, senão usa placeholder ou string vazia)
try:
    with open(r'd:\antigravity\oraex\cmdb\oraex_logo.png', 'rb') as img_file:
        logo_b64 = base64.b64encode(img_file.read()).decode('utf-8')
        logo_src = f"data:image/png;base64,{logo_b64}"
except Exception as e:
    logo_src = "" # Fallback se não encontrar

# Preparar dados para gráficos
db_names = list(data['databases'].keys())
db_sizes = [] # Precisaríamos do tamanho total por DB, vamos estimar somando as tabelas (nao temos isso direto no resumo, mas ok)

# Vamos recalcular o tamanho total aproximado por DB baseado nas top tables ou usar contagem
# Como não temos tamanho total exato no JSON resumo, vamos focar nos dados que TEMOS: Top Tabelas.

top_tables = data['top20_tables']
top_labels = [t['Full Table Name'].split('.')[-1] for t in top_tables[:10]]
top_data_size = []
top_index_size = []

def parse_size(size_str):
    if not isinstance(size_str, str): return 0
    # Remove pontos de milhar e substitui virgula decimal por ponto
    clean_str = size_str.upper().replace('.', '').replace(',', '.')
    
    if 'GB' in clean_str:
        return float(clean_str.replace(' GB', '').strip()) * 1024
    elif 'MB' in clean_str:
        return float(clean_str.replace(' MB', '').strip())
    elif 'KB' in clean_str:
        return float(clean_str.replace(' KB', '').strip()) / 1024
    # Tenta converter apenas o numero se nao tiver unidade (bytes)
    try:
        return float(clean_str.replace(' BYTES', '').strip()) / (1024*1024)
    except:
        return 0

for t in top_tables[:10]:
    top_data_size.append(parse_size(t['Data']))
    top_index_size.append(parse_size(t['Indexes']))

# Gerar HTML
html_content = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Relatório Capacidade Oracle | ORAEX</title>
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
        
        .kpi-grid {{ display: grid; grid-template-columns: repeat(4, 1fr); gap: 20px; margin-bottom: 50px; }}
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
        
        table {{ width: 100%; border-collapse: collapse; }}
        th {{ padding: 12px 14px; text-align: left; font-size: 0.65rem; text-transform: uppercase; letter-spacing: 0.1em; color: var(--text-muted); font-weight: 600; border-bottom: 2px solid var(--border); background: var(--bg-light); }}
        td {{ padding: 14px; border-bottom: 1px solid var(--border); font-size: 0.85rem; }}
        tr:hover td {{ background: var(--bg-light); }}
        
        .badge {{ display: inline-flex; padding: 4px 10px; border-radius: 100px; font-size: 0.7rem; font-weight: 600; }}
        .badge.blue {{ background: var(--oraex-blue-bg); color: var(--oraex-blue); }}
        .badge.green {{ background: #D1FAE5; color: var(--success); }}
        .badge.yellow {{ background: #FEF3C7; color: var(--warning); }}
        .badge.red {{ background: #FEE2E2; color: var(--danger); }}
        
        .footer {{ text-align: center; padding: 40px 0; margin-top: 40px; border-top: 1px solid var(--border); background: var(--bg-white); }}
        .footer-logo img {{ height: 40px; }}
        .footer p {{ color: var(--text-muted); font-size: 0.85rem; margin-top: 12px; }}
        
        @media (max-width: 1100px) {{ .kpi-grid {{ grid-template-columns: repeat(2, 1fr); }} }}
        @media (max-width: 768px) {{ .kpi-grid, .two-col {{ grid-template-columns: 1fr; }} .hero h1 {{ font-size: 1.8rem; }} .container {{ padding: 30px 20px; }} }}
    </style>
</head>
<body>
    <header class="header">
        <div class="header-inner">
            <img src="{logo_src}" alt="ORAEX" class="logo-img">
            <div class="header-meta">Gerado em {datetime.now().strftime('%d/%m/%Y às %H:%M')}</div>
        </div>
    </header>
    
    <div class="container">
        <div class="hero">
            <div class="hero-badge"><span class="dot"></span><span>Análise de Capacidade</span></div>
            <h1>Relatório Storage {data['servidor']}</h1>
            <p class="subtitle">Visão consolidada de ocupação, tabelas e objetos de banco de dados</p>
        </div>
        
        <div class="kpi-grid">
            <div class="kpi-card primary">
                <div class="kpi-icon">💾</div>
                <div class="kpi-value">{data['total_databases']}</div>
                <div class="kpi-label">Databases</div>
                <div class="kpi-sub">Bancos ativos</div>
            </div>
            <div class="kpi-card success">
                <div class="kpi-icon">📋</div>
                <div class="kpi-value">{data['total_tabelas']}</div>
                <div class="kpi-label">Tabelas</div>
                <div class="kpi-sub">Objetos analisados</div>
            </div>
            <div class="kpi-card">
                <div class="kpi-icon">🔢</div>
                <div class="kpi-value">{(data['total_registros']/1000000000):.1f} Bi</div>
                <div class="kpi-label">Registros</div>
                <div class="kpi-sub">{data['total_registros']:,} linhas totais</div>
            </div>
            <div class="kpi-card warning">
                <div class="kpi-icon">🗑️</div>
                <div class="kpi-value">{data['advanced_metrics']['total_unused_mb']/1024:.1f} GB</div>
                <div class="kpi-label">Fragmentação</div>
                <div class="kpi-sub">Espaço desperdiçado</div>
            </div>
        </div>
        
        <div class="two-col">
            <div class="card">
                <div class="card-title">Distribuição por Schema (Top 10)</div>
                <div id="chart-schema"></div>
            </div>
            <div class="card">
                <div class="card-title">Top 5 Tabelas Fragmentadas (Unused Space)</div>
                <table style="font-size: 0.8rem;">
                    <thead>
                        <tr>
                            <th>Tabela</th>
                            <th>Unused</th>
                            <th>Total</th>
                        </tr>
                    </thead>
                    <tbody>
                        {"".join([f"<tr><td>{t['Full Table Name']}</td><td><span class='badge red'>{t['Unused']}</span></td><td>{t['Total Reserved Size']}</td></tr>" for t in data['advanced_metrics']['top_fragmented'][:5]])}
                    </tbody>
                </table>
            </div>
        </div>
        <br>

        <section class="section">
            <div class="section-header"><div class="section-line"></div><h2 class="section-title">Top 10 Tabelas - Distribuição de Espaço</h2></div>
            <div class="card"><div id="chart-tables"></div></div>
        </section>
        
        <section class="section">
            <div class="section-header" style="justify-content: space-between;">
                <div style="display: flex; align-items: center; gap: 14px;">
                    <div class="section-line"></div><h2 class="section-title">Detalhamento das Maiores Tabelas</h2>
                </div>
                <button onclick="exportTableToCSV('relatorio_storage.csv')" style="padding: 8px 16px; background: var(--oraex-blue); color: white; border: none; border-radius: 6px; cursor: pointer; font-size: 0.85rem;">📥 Exportar CSV</button>
            </div>
            <div class="card">
                <table id="mainTable">
                    <thead>
                        <tr>
                            <th>Database</th>
                            <th>Tabela</th>
                            <th>Registros</th>
                            <th>Total Size</th>
                            <th>Data</th>
                            <th>Index</th>
                            <th>% Index</th>
                        </tr>
                    </thead>
                    <tbody>"""

for t in top_tables:
    index_weight = float(t['% Index Weight'])
    badge_color = 'green'
    if index_weight > 0.5: badge_color = 'yellow'
    if index_weight > 0.7: badge_color = 'red'
    
    html_content += f"""
                        <tr>
                            <td><span class="badge blue">{t['Database']}</span></td>
                            <td><strong>{t['Full Table Name']}</strong></td>
                            <td>{t['Total Rows']:,}</td>
                            <td><strong>{t['Total Reserved Size']}</strong></td>
                            <td style="color: var(--text-gray);">{t['Data']}</td>
                            <td style="color: var(--text-gray);">{t['Indexes']}</td>
                            <td><span class="badge {badge_color}">{index_weight:.1%}</span></td>
                        </tr>"""

# Adicionar script de gráfico de schema e exportação
schema_labels = [s['label'] for s in data['advanced_metrics']['schemas']]
schema_values = [s['value'] for s in data['advanced_metrics']['schemas']]

html_content += f"""
                    </tbody>
                </table>
            </div>
        </section>
        
        <script>
            // Gráfico de Top Tabelas
            var trace1 = {{
              x: {json.dumps(top_labels)},
              y: {json.dumps(top_data_size)},
              name: 'Dados (MB)',
              type: 'bar',
              marker: {{ color: '#0000FF' }}
            }};
            
            var trace2 = {{
              x: {json.dumps(top_labels)},
              y: {json.dumps(top_index_size)},
              name: 'Índices (MB)',
              type: 'bar',
              marker: {{ color: '#93C5FD' }}
            }};
            
            var layoutBar = {{
                barmode: 'stack',
                font: {{ family: 'Inter', color: '#374151' }},
                paper_bgcolor: 'rgba(0,0,0,0)',
                plot_bgcolor: 'rgba(0,0,0,0)',
                hovermode: 'closest',
                showlegend: true,
                legend: {{ orientation: "h", y: 1.1, x: 0.5, xanchor: "center" }},
                margin: {{ t: 30, b: 40, l: 40, r: 20 }},
                height: 350
            }};
            
            Plotly.newPlot('chart-tables', [trace1, trace2], layoutBar, {{responsive: true, displayModeBar: false}});

            // Gráfico de Schemas
            var tracePie = {{
                labels: {json.dumps(schema_labels)},
                values: {json.dumps(schema_values)},
                type: 'pie',
                hole: .4,
                marker: {{ colors: ['#0000FF', '#4D4DFF', '#93C5FD', '#BFDBFE', '#E5E7EB'] }}
            }};
            
            var layoutPie = {{
                font: {{ family: 'Inter', color: '#374151' }},
                paper_bgcolor: 'rgba(0,0,0,0)',
                showlegend: true,
                legend: {{ orientation: "v", x: 1.1, y: 0.5 }},
                margin: {{ t: 20, b: 20, l: 20, r: 20 }},
                height: 300
            }};
            
            Plotly.newPlot('chart-schema', [tracePie], layoutPie, {{responsive: true, displayModeBar: false}});

            // Função de Exportação CSV
            function exportTableToCSV(filename) {{
                var csv = [];
                var rows = document.querySelectorAll("#mainTable tr");
                
                for (var i = 0; i < rows.length; i++) {{
                    var row = [], cols = rows[i].querySelectorAll("td, th");
                    for (var j = 0; j < cols.length; j++) 
                        row.push(cols[j].innerText);
                    csv.push(row.join(";"));        
                }}

                var csvFile = new Blob([csv.join("\\n")], {{type: "text/csv"}});
                var downloadLink = document.createElement("a");
                downloadLink.download = filename;
                downloadLink.href = window.URL.createObjectURL(csvFile);
                downloadLink.style.display = "none";
                document.body.appendChild(downloadLink);
                downloadLink.click();
            }}
        </script>
    </div>
    
    <footer class="footer">
        <div class="footer-logo"><img src="{logo_src}" alt="ORAEX"></div>
        <p>Relatório Confidencial • GetNet Infrastructure Analysis</p>
        <p>© 2025 ORAEX Cloud Consulting</p>
    </footer>
</body>
</html>
"""

# Salvar arquivo HTML
file_name = f"relatorio_espaco_GNCASNPW03527.html"
with open(file_name, 'w', encoding='utf-8') as f:
    f.write(html_content)

print(f"Relatório gerado em: {file_name}")
