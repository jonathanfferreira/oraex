# Alertas PSU GetNet - Automação via GitHub Actions

Este repositório contém scripts de automação para gerenciamento das atividades de PSU Oracle do cliente GetNet.

## 🚀 Funcionalidades

- 📅 **Lembrete Diário** (17:00 BRT): Aviso das GMUDs da noite
- 📊 **Resumo Semanal** (Segunda 09:00 BRT): Status geral das atividades
- 🚨 **Alertas Críticos**: Servidores com PSU desatualizado

## ⚙️ Configuração

### Secrets Necessários

No GitHub, vá em **Settings → Secrets and variables → Actions** e adicione:

| Nome | Valor |
|------|-------|
| `SLACK_WEBHOOK_URL` | URL do webhook do Slack |

## 📁 Estrutura

```
oraex/
├── cmdb/
│   ├── alertas_slack_psu.py      # Script principal
│   ├── calendario_psu_2026.xlsx  # Planilha de planejamento
│   └── guia_slack_webhook.md     # Documentação
└── .github/
    └── workflows/
        └── alertas-psu.yml       # Automação GitHub Actions
```

## 🔧 Execução Manual

```bash
python cmdb/alertas_slack_psu.py
```

---

*ORAEX Cloud Consulting © 2025*
