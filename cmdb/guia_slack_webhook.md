# Guia: Configurando Webhook do Slack para Alertas PSU

## Passo a Passo para Criar o Webhook

### 1. Acesse o Slack App Directory
1. Abra o Slack da ORAEX no navegador: https://app.slack.com
2. Clique no nome do workspace no canto superior esquerdo
3. Vá em **Settings & administration** → **Manage apps**

### 2. Criar um Incoming Webhook
1. Na página que abrir, procure por **"Incoming WebHooks"** na barra de busca
2. Clique em **"Add to Slack"**
3. Escolha o **canal** onde quer receber os alertas (ex: `#getnet-psu` ou `#alertas`)
4. Clique em **"Add Incoming WebHooks Integration"**

### 3. Copiar a URL do Webhook
Você verá uma URL parecida com:
```
https://hooks.slack.com/services/T00000000/B00000000/XXXXXXXXXXXXXXXXXXXXXXXX
```
**Copie essa URL** - ela será usada para enviar alertas automáticos.

### 4. Me Envie a URL
Após criar, cole a URL aqui ou me informe o nome do canal criado.

---

## Alternativa: Usando Slack Workflow Builder (Sem Código)

Se preferir uma solução sem código:
1. No Slack, clique em **Automações** (ícone de raio ⚡)
2. **Create Workflow** → **From a webhook**
3. Isso permite criar automações visuais sem programação

---

## Próximo Passo
Após configurar o webhook, vou criar um script Python simples que:
- Lê a planilha de GMUDs do dia
- Envia lembrete às 17:00 com as atividades da noite
- Envia resumo semanal às segundas
