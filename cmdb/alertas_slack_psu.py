"""
Script de Alertas Slack para Gestão PSU GetNet
===============================================
Este script envia alertas automáticos para o canal Slack da ORAEX.

CONFIGURAÇÃO:
1. Crie um Incoming Webhook no Slack (veja guia_slack_webhook.md)
2. Cole a URL do webhook na variável SLACK_WEBHOOK_URL abaixo
3. Execute o script com: python alertas_slack_psu.py

AUTOMAÇÃO:
- No Windows, use o Agendador de Tarefas (Task Scheduler)
- Configure para executar às 17:00 de segunda a sexta
"""

import json
import urllib.request
from datetime import datetime, timedelta
import openpyxl
import os

# ============ CONFIGURAÇÃO ============
# A URL do webhook deve ser configurada como variável de ambiente
# Para uso local, crie um arquivo .env com: SLACK_WEBHOOK_URL=sua_url_aqui
# Para GitHub Actions, configure em Settings -> Secrets -> Actions
import os
SLACK_WEBHOOK_URL = os.environ.get("SLACK_WEBHOOK_URL", "")

# Caminho para a planilha de planejamento
PLANILHA_PATH = r"D:\antigravity\oraex\cmdb\calendario_psu_2026.xlsx"

# ============ FUNÇÕES ============

def enviar_slack(mensagem: str, webhook_url: str = SLACK_WEBHOOK_URL) -> bool:
    """Envia uma mensagem para o Slack via webhook."""
    try:
        payload = {
            "text": mensagem,
            "username": "PSU Bot GetNet",
            "icon_emoji": ":robot_face:"
        }
        data = json.dumps(payload).encode('utf-8')
        req = urllib.request.Request(
            webhook_url, 
            data=data,
            headers={'Content-Type': 'application/json'}
        )
        urllib.request.urlopen(req)
        print(f"✅ Mensagem enviada para o Slack!")
        return True
    except Exception as e:
        print(f"❌ Erro ao enviar para Slack: {e}")
        return False


def carregar_servidores_criticos(planilha_path: str) -> list:
    """Carrega servidores com status 'Crítico' da planilha."""
    try:
        wb = openpyxl.load_workbook(planilha_path)
        ws = wb['Servidores']
        criticos = []
        for row in ws.iter_rows(min_row=4, values_only=True):
            if row[4] == 'Crítico':  # Coluna Status
                criticos.append({
                    'hostname': row[0],
                    'ambiente': row[1],
                    'psu_atual': row[2],
                    'ultima_atualizacao': row[5]
                })
        return criticos
    except Exception as e:
        print(f"Erro ao carregar planilha: {e}")
        return []


def lembrete_diario():
    """Envia lembrete diário às 17:00 com as atividades da noite."""
    hoje = datetime.now()
    dia_semana = hoje.strftime('%A')
    data_formatada = hoje.strftime('%d/%m/%Y')
    
    mensagem = f"""
:calendar: *LEMBRETE DIÁRIO PSU - {data_formatada}*

:clock6: *Janela de Execução Hoje:*

:small_blue_diamond: *DEV* (18:00 - 03:00): 2 GMUDs planejadas
:small_blue_diamond: *HML* (18:00 - 03:00): 3 GMUDs planejadas  
:small_blue_diamond: *PROD* (22:00 - 05:00): ~5 GMUDs planejadas

:warning: *Lembretes:*
• Verificar aprovação do plantonista antes de iniciar
• Validar acesso aos servidores
• Atualizar status na planilha após execução

:rocket: Bom trabalho, equipe!
"""
    return enviar_slack(mensagem)


def alerta_servidores_criticos():
    """Envia alerta sobre servidores em estado crítico."""
    criticos = carregar_servidores_criticos(PLANILHA_PATH)
    
    if not criticos:
        print("Nenhum servidor crítico encontrado.")
        return True
    
    lista_servidores = "\n".join([
        f"• `{s['hostname']}` - {s['ambiente']} - PSU {s['psu_atual']}"
        for s in criticos
    ])
    
    mensagem = f"""
:rotating_light: *ALERTA: SERVIDORES CRÍTICOS*

Os seguintes servidores estão com PSU desatualizado (3+ quarters):

{lista_servidores}

:point_right: *Ação necessária:* Priorizar atualização destes hosts.
"""
    return enviar_slack(mensagem)


def resumo_semanal():
    """Envia resumo semanal às segundas-feiras."""
    hoje = datetime.now()
    semana_passada = hoje - timedelta(days=7)
    
    mensagem = f"""
:bar_chart: *RESUMO SEMANAL PSU - Semana {semana_passada.strftime('%d/%m')} a {hoje.strftime('%d/%m/%Y')}*

:white_check_mark: *GMUDs Executadas:* [preencher]
:x: *GMUDs Canceladas:* [preencher]
:arrows_counterclockwise: *GMUDs Replanejadas:* [preencher]

:chart_with_upwards_trend: *Taxa de Sucesso:* [calcular]%

:calendar: *Próxima Semana:*
• [GMUDs planejadas]

:memo: Atualizar dados na planilha calendario_psu_2026.xlsx
"""
    return enviar_slack(mensagem)


def testar_conexao():
    """Testa a conexão com o Slack."""
    mensagem = ":wave: *Teste de conexão do Bot PSU GetNet!*\n\nSe você está vendo esta mensagem, a integração está funcionando! :white_check_mark:"
    return enviar_slack(mensagem)


# ============ EXECUÇÃO ============
if __name__ == "__main__":
    print("=" * 50)
    print("Sistema de Alertas PSU GetNet")
    print("=" * 50)
    
    print("\nEscolha uma opção:")
    print("1. Testar conexão com Slack")
    print("2. Enviar lembrete diário")
    print("3. Enviar alerta de servidores críticos")
    print("4. Enviar resumo semanal")
    print("5. Sair")
    
    opcao = input("\nOpção: ").strip()
    
    if opcao == "1":
        testar_conexao()
    elif opcao == "2":
        lembrete_diario()
    elif opcao == "3":
        alerta_servidores_criticos()
    elif opcao == "4":
        resumo_semanal()
    elif opcao == "5":
        print("Até mais!")
    else:
        print("Opção inválida.")
