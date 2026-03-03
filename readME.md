# Chatbot whatsapp calendar
## 🚀 Automação Integrada com Google Workspace

Este repositório contém pequena automação de escritório desenvolvida em **Google Apps Script**. O sistema transforma uma Planilha Google em um centro de comando capaz de gerenciar agendas(Google Calendar) via WhatsApp, automatizar envios de e-mail(Gmail) e consultar dados corporativos (CNPJ).


## 🌟 Funcionalidades

* **📅 Assistente de Agenda (WhatsApp):** Receba resumos diários ou semanais dos seus compromissos, feriados e eventos de dia inteiro diretamente no seu celular.
* **📧 Automação de E-mail:** Envio em massa de e-mails personalizados baseados em listas na planilha com controle de status.
* **🏢 Consulta de CNPJ:** Validação e preenchimento automático de dados cadastrais (Razão Social e Situação) via API.
* **🛠️ Instalador Automático:** Função que prepara toda a estrutura da planilha (abas, cores e cabeçalhos) com um clique.


## 🛠️ Tecnologias Utilizadas

* **Linguagem:** JavaScript (Google Apps Script)
* **APIs do Google:** Calendar API, Gmail API, Spreadsheet Service.
* **APIs Externas:**  [callmebot](https://www.callmebot.com/blog/free-api-whatsapp-messages/) (WhatsApp Gateway), [Brasil_API](https://brasilapi.com.br/) (Dados Públicos). ![image](https://www.callmebot.com/wp-content/uploads/2019/10/Logo-Negro_x2.png)

  <img src="https://github.com/ildem96/Chatbot-whatsapp-calendar/blob/main/snapshot.jpeg?raw=true" width="400">  <img src="https://github.com/ildem96/Chatbot-whatsapp-calendar/blob/main/snapshot02.jpeg?raw=true" width="400">  <img src="https://github.com/ildem96/Chatbot-whatsapp-calendar/blob/main/%7B8DA7334F-2923-42DF-AB5F-F39A02424655%7D.png?raw=true" width="350">



## 📖 Manual de Instalação

Siga os passos abaixo para configurar o bot

### 1. Preparação da Planilha

1. Crie uma nova [Planilha Google](https://sheet.new).
2. No menu superior, vá em **Extensões** > **Apps Script**.
3. Apague todo o código existente no editor e cole o conteúdo do arquivo `Código.gs` deste repositório.
4. Clique no ícone de disquete (Salvar) e dê o nome de "SmartOffice Bot".

### 2. Executando o Instalador

1. Atualize a página da sua planilha. Um novo menu chamado **🚀 Painel de Automação** aparecerá no topo.
2. Clique em **🚀 Painel de Automação** > **⚙️ Configurar Planilha (Instalador)**.
3. O Google pedirá permissões de segurança. Clique em *Avançado* e *Acessar SmartOffice Bot (não seguro)* para autorizar.
4. Após a execução, as abas `Config`, `Logs` e `Envios` serão criadas automaticamente.

### 3. Configuração do WhatsApp

1. Adicione o número do **CallMeBot** aos seus contatos: `+34 644 66 32 62`.
2. Envie a mensagem: `I allow callmebot to send me messages`.
3. Você receberá uma **ApiKey**.
4. Na sua planilha, vá na aba **Config** e preencha:
* **Telefone:** Seu número com DDD (ex: `5551999999999`).
* **ApiKey:** A chave recebida no passo anterior.



### 4. Ativando o Envio Automático (Trigger)

Para receber a agenda todo dia sem precisar clicar em nada:

1. No editor de Apps Script, clique no ícone de **Relógio** (Acionadores) na barra lateral esquerda.
2. Clique em **+ Adicionar Acionador**.
3. Selecione a função: `enviarAgendaDoDia`.
4. Origem do evento: `Baseado no tempo`.
5. Tipo de timer: `Contador diário`.
6. Horário: Selecione o período desejado (ex: `07:00 às 08:00`).

---

## 📋 Como Usar

* **Para a Agenda:** Use o menu superior ou o acionador automático. O bot filtrará seus eventos do Google Agenda e enviará formatado com emojis.
* **Para E-mails:** Preencha Nome e E-mail na aba `Envios` e use o menu **Ferramentas de E-mail**.
* **Para CNPJ:** Digite um CNPJ em qualquer célula, selecione-a e vá em **Serviços Externos** > **Consultar CNPJ**.
