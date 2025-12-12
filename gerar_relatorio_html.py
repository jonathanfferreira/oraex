import pandas as pd
import plotly.express as px
import plotly.io as pio
from jinja2 import Template
from datetime import datetime
import os
import base64

# Configuração
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ARQUIVO_PLANILHA = os.path.join(BASE_DIR, 'ORAEX_Planejamento_GetNet_2026.xlsx')
ARQUIVO_SAIDA = os.path.join(BASE_DIR, 'Relatorio_GMUDs_2026.html')
ARQUIVO_LOGO = os.path.join(BASE_DIR, 'oraex_logo.png')
ARQUIVO_TEMPLATE = os.path.join(BASE_DIR, 'template_relatorio.html')
ABA_INVENTARIO = 'INVENTÁRIO SERVIDORES'

MESES = [
    'JANEIRO-26', 'FEVEREIRO-26', 'MARÇO-26', 'ABRIL-26', 'MAIO-26', 'JUNHO-26',
    'JULHO-26', 'AGOSTO-26', 'SETEMBRO-26', 'OUTUBRO-26', 'NOVEMBRO-26', 'DEZEMBRO-26'
]

def load_template():
    with open(ARQUIVO_TEMPLATE, 'r', encoding='utf-8') as f:
        return f.read()

TEMPLATE_HTML = "" # Placeholder, will be loaded dynamically

def get_logo_b64():
    try:
        if os.path.exists(ARQUIVO_LOGO):
            with open(ARQUIVO_LOGO, "rb") as image_file:
                return base64.b64encode(image_file.read()).decode('utf-8')
    except: return ""

def carregar_gmuds():
    dfs = []
    for mes in MESES:
        try:
            df = pd.read_excel(ARQUIVO_PLANILHA, sheet_name=mes, header=2)
            df.columns = df.columns.astype(str).str.strip().str.upper()
            df['MES_REF'] = mes
            df = df.dropna(subset=['CLIENTE'])
            dfs.append(df)
        except: pass
    return pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()

def carregar_inventario():
    try:
        df = pd.read_excel(ARQUIVO_PLANILHA, sheet_name=ABA_INVENTARIO)
        # Assumindo colunas: HOSTNAME, AMBIENTE, IP, TIPO DB, GRID VERSION, DB VERSION, PSU ATUAL, STATUS PSU
        df.columns = df.columns.astype(str).str.strip().str.upper()
        return df
    except Exception as e:
        print(f"Erro ao ler inventário: {e}")
        return pd.DataFrame()

def gerar_relatorio():
    print("Gerando Relatório Completo...")
    
    # --- GMUDS ---
    df_gmud = carregar_gmuds()
    
    # Calcular KPIs GMUD
    if not df_gmud.empty:
        col_status = 'STATUS GMUD'
        if col_status not in df_gmud.columns: col_status = 'STATUS' # Fallback
        
        df_gmud[col_status] = df_gmud[col_status].fillna('NOVO').astype(str).str.upper().str.strip()
        
        total_gmuds = len(df_gmud)
        sucesso_count = len(df_gmud[df_gmud[col_status] == 'ENCERRADA'])
        falha_count = len(df_gmud[df_gmud[col_status].isin(['REPLANEJAR', 'CANCELADA', 'FALHA'])])
        base_calc = sucesso_count + falha_count
        taxa = (sucesso_count / base_calc * 100) if base_calc > 0 else 0
        
        # Plots GMUD
        color_map = {'ENCERRADA': '#198754', 'REPLANEJAR': '#dc3545', 'NOVO': '#e9ecef', 'PROGRAMADA': '#0d6efd', 'AUTORIZAR': '#ffc107', 'AVALIAR': '#fd7e14'}
        
        fig_m = px.bar(df_gmud.groupby(['MES_REF', col_status]).size().reset_index(name='QTD'), 
                      x='MES_REF', y='QTD', color=col_status, color_discrete_map=color_map)
        fig_m.update_layout(plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', font_family="Segoe UI", margin=dict(t=10,l=10,r=10,b=10))
        plot_mensal = pio.to_html(fig_m, full_html=False, include_plotlyjs='cdn', config={'displayModeBar': False})
        
        fig_p = px.pie(df_gmud, names=col_status, color=col_status, color_discrete_map=color_map, hole=0.5)
        fig_p.update_layout(plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', font_family="Segoe UI", margin=dict(t=10,l=10,r=10,b=10))
        plot_pizza = pio.to_html(fig_p, full_html=False, include_plotlyjs=False, config={'displayModeBar': False})
        
        # Tabela GMUD
        cols = ['DATA INICIO', 'CLIENTE', 'ENTORNO', 'GMUD', 'TÍTULO', col_status, 'DESIGNADO A']
        cols = [c for c in cols if c in df_gmud.columns]
        gmud_html = df_gmud[df_gmud[col_status]!='NOVO'].tail(50).to_html(classes='table table-hover table-sm small', index=False, border=0)

        # Timeline
        plot_timeline = ""
        try:
             range_cols = ['DATA INICIO', 'DATA FIM']
             if all(c in df_gmud.columns for c in range_cols):
                 df_gmud[range_cols[0]] = pd.to_datetime(df_gmud[range_cols[0]], errors='coerce')
                 df_gmud[range_cols[1]] = pd.to_datetime(df_gmud[range_cols[1]], errors='coerce')
                 df_tl = df_gmud.dropna(subset=range_cols)
                 df_tl = df_tl[~df_tl[col_status].isin(['NOVO', 'CANCELADA'])]
                 
                 if not df_tl.empty:
                     y_ax = 'AMBIENTE' if 'AMBIENTE' in df_tl.columns else 'CLIENTE'
                     fig_tl = px.timeline(df_tl, x_start=range_cols[0], x_end=range_cols[1], y=y_ax, color=col_status, 
                                          color_discrete_map=color_map, hover_name='TÍTULO')
                     fig_tl.update_yaxes(autorange="reversed")
                     fig_tl.update_layout(plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', font_family="Inter", 
                                          showlegend=True, margin=dict(t=10,l=10,r=10,b=10), height=350)
                     plot_timeline = pio.to_html(fig_tl, full_html=False, include_plotlyjs=False, config={'displayModeBar': True, 'responsive': True})
        except Exception as e: print(f"Erro Timeline: {e}")
    else:
        total_gmuds = 0; sucesso_count=0; falha_count=0; taxa=0; plot_mensal=""; plot_pizza=""; gmud_html="<p>Sem dados</p>"; plot_timeline=""

    # --- RETRO 2025 ---
    plot_2025_mensal = ""
    plot_2025_status = ""
    kpi_2025_total = 0
    kpi_2025_sucesso = "0%"
    try:
        file_2025 = os.path.join(DIR_BASE, 'consolidated_gmuds_2025.xlsx')
        if os.path.exists(file_2025):
            df_25 = pd.read_excel(file_2025)
            # Normalizar
            df_25.columns = df_25.columns.astype(str).str.strip().str.upper()
            col_dt_25 = next((c for c in df_25.columns if 'DATA' in c or 'INICIO' in c), None)
            col_st_25 = next((c for c in df_25.columns if 'STATUS' in c or 'SITU' in c), None)
            
            if col_dt_25 and col_st_25:
                df_25[col_dt_25] = pd.to_datetime(df_25[col_dt_25], errors='coerce')
                kpi_2025_total = len(df_25)
                
                # Sucesso Rate
                suc_25 = df_25[df_25[col_st_25].astype(str).str.contains('SUCESSO|CONCLU', case=False, na=False)]
                if kpi_2025_total > 0:
                    kpi_2025_sucesso = f"{(len(suc_25)/kpi_2025_total)*100:.1f}%"
                
                # Plot Mensal 2025
                df_25['Mes'] = df_25[col_dt_25].dt.to_period('M').astype(str)
                df_mes_25 = df_25.groupby('Mes').size().reset_index(name='Qtd')
                fig_m25 = px.bar(df_mes_25, x='Mes', y='Qtd', text='Qtd', title='', color_discrete_sequence=['#3b82f6'])
                fig_m25.update_layout(plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', font_family="Inter", margin=dict(t=10,l=10,r=10,b=20))
                plot_2025_mensal = pio.to_html(fig_m25, full_html=False, include_plotlyjs=False, config={'displayModeBar': False})
                
                # Plot Status 2025
                df_st_25 = df_25[col_st_25].value_counts().reset_index()
                df_st_25.columns = ['Status', 'Qtd']
                fig_s25 = px.pie(df_st_25, names='Status', values='Qtd', hole=0.7, color_discrete_sequence=px.colors.qualitative.Pastel)
                fig_s25.update_layout(plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', font_family="Inter", showlegend=True, margin=dict(t=0,l=0,r=0,b=0))
                plot_2025_status = pio.to_html(fig_s25, full_html=False, include_plotlyjs=False, config={'displayModeBar': False})
    except Exception as e:
        print(f"Erro 2025: {e}")

    # --- INVENTÁRIO (Lógica Nova: Primary + Standby) ---
    # Ler abas originais para garantir dados de Standby
    original_sheets = ['GetNet - Oracle Databases', 'PagoNxt - Databases']
    df_inv_list = []
    
    for sh in original_sheets:
        try:
            d = pd.read_excel(ARQUIVO_PLANILHA, sheet_name=sh)
            d.columns = d.columns.astype(str).str.strip().str.upper()
            
            # Normalizar nomes de colunas (as vezes muda um pouco)
            # Mapeamento basico
            rename_map = {}
            for c in d.columns:
                if 'ENV' in c: rename_map[c] = 'AMBIENTE'
                if 'PRIMARY' in c: rename_map[c] = 'HOSTNAME'
                if 'STANDBY' in c: rename_map[c] = 'STANDBY'
                if 'DB VERSION' in c: rename_map[c] = 'VERSION'
                if 'PSU' in c and 'VER' in c: rename_map[c] = 'PSU VERSION'
                if 'SITUA' in c: rename_map[c] = 'STATUS PSU' # Situação costuma ser o status
            
            d = d.rename(columns=rename_map)
            # Dedup columns (keep first)
            d = d.loc[:, ~d.columns.duplicated()]
            df_inv_list.append(d)
        except Exception as e:
            print(f"Erro ao ler aba {sh}: {e}")

    df_raw = pd.concat(df_inv_list, ignore_index=True) if df_inv_list else pd.DataFrame()
    
    inv_total = 0; inv_criticos = 0; inv_atualizados = 0
    plot_inv_env=""; plot_inv_ver=""; plot_inv_status=""; inv_html="<p>Sem dados</p>"

    if not df_raw.empty:
        # Lógica de Explosão (Unpivot) para Contagem Real
        # Criar DF de Primários
        df_p = df_raw.copy()
        df_p['TYPE'] = 'PRIMARY'
        
        # Criar DF de Standbys (onde houver)
        if 'STANDBY' in df_raw.columns:
            df_s = df_raw[df_raw['STANDBY'].notna()].copy()
            df_s['HOSTNAME'] = df_s['STANDBY'] # O hostname passa a ser o do standby
            df_s['TYPE'] = 'STANDBY'
            # Manter as outras colunas (Ambiente, Versão, etc assumimos iguais ao primary)
        else:
            df_s = pd.DataFrame()
        
        # Juntar tudo para estatísticas
        df_full_servers = pd.concat([df_p, df_s], ignore_index=True)
        
        # Filtrar hostnames vazios
        df_full_servers = df_full_servers.dropna(subset=['HOSTNAME'])
        
        inv_total = len(df_full_servers)
        
        # KPIs baseados em STATUS PSU (coluna SITUAÇÃO/STATUS)
        col_status_psu = 'STATUS PSU'
        if col_status_psu not in df_full_servers.columns:
            # Tentar achar
            poss = [c for c in df_full_servers.columns if 'SITUA' in c or 'STATUS' in c]
            if poss: col_status_psu = poss[0]
            
        if col_status_psu in df_full_servers.columns:
            inv_criticos = len(df_full_servers[df_full_servers[col_status_psu].astype(str).str.contains('Desatualizado|Atenção', case=False, na=False)])
            inv_atualizados = len(df_full_servers[df_full_servers[col_status_psu].astype(str).str.contains('Atualizado|Ok', case=False, na=False)])

        # Plots (AGORA REFINADOS)
        
        # Paleta Inspirada no Protótipo
        colors_proto = {'Atualizado': '#10b981', 'Ok': '#10b981', 'ENCERRADA': '#10b981', 
                        'Desatualizado': '#ef4444', 'FALHA': '#ef4444', 'CANCELADA': '#ef4444', 'REPLANEJAR': '#f59e0b',
                        'Atenção': '#f59e0b', 'NOVO': '#e5e7eb', 'DEV': '#3b82f6', 'HML': '#0ea5e9', 'PROD': '#1e3a8a'}

        # 1. Por Ambiente (Horizontal Bar)
        try:
            if 'AMBIENTE' in df_full_servers.columns:
                df_env = df_full_servers['AMBIENTE'].value_counts().reset_index()
                df_env.columns = ['Ambiente', 'Qtd']
                df_env = df_env.sort_values('Qtd', ascending=True)
                
                fig_env = px.bar(df_env, x='Qtd', y='Ambiente', text='Qtd', orientation='h', color='Ambiente',
                                 color_discrete_map=colors_proto)
                fig_env.update_traces(textposition='outside')
                fig_env.update_layout(plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', font_family="Inter", 
                                     showlegend=False, margin=dict(t=0,l=0,r=0,b=0))
                fig_env.update_xaxes(visible=False)
                plot_inv_env = pio.to_html(fig_env, full_html=False, include_plotlyjs=False, config={'displayModeBar': False, 'responsive': True})
        except Exception as e: 
            print(f"Erro Plot Env: {e}")
            plot_inv_env = ""

        # 2. Por Versão
        try:
            if 'VERSION' in df_full_servers.columns:
                df_full_servers['VerShort'] = df_full_servers['VERSION'].astype(str).str.extract(r'(\d+(?:\.\d+)?)')
                df_ver = df_full_servers['VerShort'].value_counts().reset_index()
                df_ver.columns = ['Versao', 'Qtd']
                fig_ver = px.bar(df_ver, x='Versao', y='Qtd', text='Qtd')
                fig_ver.update_traces(marker_color='#3b82f6', marker_cornerradius=5)
                fig_ver.update_layout(plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', font_family="Inter", margin=dict(t=20,l=0,r=0,b=20))
                fig_ver.update_yaxes(visible=False)
                plot_inv_ver = pio.to_html(fig_ver, full_html=False, include_plotlyjs=False, config={'displayModeBar': False, 'responsive': True})
        except Exception as e: 
            print(f"Erro Plot Version: {e}")
            plot_inv_ver = ""

        # 3. PSU Status Overview
        try:
            if col_status_psu in df_full_servers.columns:
                df_stat = df_full_servers[col_status_psu].value_counts().reset_index()
                df_stat.columns = ['Status', 'Qtd']
                
                fig_stat = px.pie(df_stat, names='Status', values='Qtd', hole=0.7, 
                                 color='Status', color_discrete_map=colors_proto)
                
                total_psu = df_stat['Qtd'].sum()
                fig_stat.update_layout(
                    plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', font_family="Inter", showlegend=True, 
                    legend=dict(orientation="h", yanchor="bottom", y=-0.2, xanchor="center", x=0.5),
                    margin=dict(t=0,l=0,r=0,b=20),
                    annotations=[dict(text=f"{total_psu}<br><span style='font-size:12px; color:gray'>Servidores</span>", x=0.5, y=0.5, font_size=24, showarrow=False)]
                )
                plot_inv_status = pio.to_html(fig_stat, full_html=False, include_plotlyjs=False, config={'displayModeBar': False, 'responsive': True})
        except Exception as e: 
            print(f"Erro Plot Status: {e}")
            plot_inv_status = ""

        # Tabela CRÍTICA (Filtro)
        try:
            cols_inv = ['HOSTNAME', 'TYPE', 'AMBIENTE', 'VERSION', col_status_psu]
            cols_to_show = [c for c in cols_inv if c in df_full_servers.columns]
            
            # Filtro de Criticidade
            mask_crit = df_full_servers[col_status_psu].astype(str).str.contains('Desatualizado|Atenção|Baixo|Crítico', case=False, na=False)
            df_critical = df_full_servers[mask_crit].copy()
            
            if df_critical.empty:
                inv_html = "<div class='p-4 text-green-600 bg-green-50 rounded-lg'>✅ Nenhum servidor crítico encontrado! Parabéns!</div>"
            else:
                inv_html = df_critical[cols_to_show].head(1000).to_html(classes='w-full text-sm text-left', index=False, border=0)
        except: 
            inv_html="<p>Sem dados</p>"

    # --- EXECUTORES ---
    try:
        col_resp = 'DESIGNADO A'
        if col_resp not in df_gmud.columns: col_resp = 'ABERTO POR'
        
        if col_resp in df_gmud.columns:
            top_exec = df_gmud[col_resp].value_counts().head(5).reset_index()
            top_exec.columns = ['Executor', 'Qtd']
            fig_exec = px.pie(top_exec, names='Executor', values='Qtd', hole=0.7, title='')
            fig_exec.update_traces(textinfo='percent')
            fig_exec.update_layout(
                 plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', font_family="Inter", showlegend=True,
                 legend=dict(orientation="v", yanchor="middle", y=0.5, xanchor="left", x=1.0),
                 margin=dict(t=0,l=0,r=0,b=0),
                 annotations=[dict(text=f"TOP 5", x=0.5, y=0.5, font_size=16, showarrow=False)]
            )
            plot_executores = pio.to_html(fig_exec, full_html=False, include_plotlyjs=False, config={'displayModeBar': False, 'responsive': True})
        else: plot_executores = ""
    except Exception as e: 
        print(f"Erro Plot Executores: {e}")
        plot_executores = ""

    # Render
    template_content = load_template()
    template = Template(template_content)
    logo_b64_data = get_logo_b64()
    html_out = template.render(
        data_geracao=datetime.now().strftime('%d/%m/%Y %H:%M'),
        logo_b64=logo_b64_data,
        # GMUD Data
        total_gmuds=total_gmuds, total_sucesso=sucesso_count, total_falhas=falha_count, taxa_sucesso=f"{taxa:.1f}",
        plot_mensal=plot_mensal, plot_pizza=plot_pizza, tabela_gmuds=gmud_html,
        plot_executores=plot_executores, plot_timeline=plot_timeline,
        # Inv Data
        inv_total=inv_total, inv_criticos=inv_criticos, inv_atualizados=inv_atualizados,
        plot_inv_env=plot_inv_env, plot_inv_ver=plot_inv_ver, plot_inv_status=plot_inv_status, tabela_inv=inv_html,
        # 2025 Data
        plot_2025_mensal=plot_2025_mensal, plot_2025_status=plot_2025_status, 
        kpi_2025_total=kpi_2025_total, kpi_2025_sucesso=kpi_2025_sucesso
    )

    with open(ARQUIVO_SAIDA, 'w', encoding='utf-8') as f:
        f.write(html_out)
    
    print(f"Relatório Completo gerado em: {ARQUIVO_SAIDA}")

if __name__ == '__main__':
    gerar_relatorio()
