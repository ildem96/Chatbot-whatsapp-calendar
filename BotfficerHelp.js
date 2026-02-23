//SISTEMA INTEGRADO DE AUTOMAÇÃO DE ESCRITÓRIO (MVP)
// Funcionalidades: Bot WhatsApp (Agenda), Consulta de CNPJ, Envio de E-mails e Logs.
// Desenvolvido para Google Apps Script.
// 
// --- 1. CONFIGURAÇÃO E INSTALAÇÃO ---
// Cria as abas necessárias e define os cabeçalhos.
//Execute esta função ao instalar o script pela primeira vez.
function configurarMVP() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ui = SpreadsheetApp.getUi();
    // Configuração da aba 'Config'
    var abaConfig = ss.getSheetByName("Config") || ss.insertSheet("Config");
    abaConfig.clear();
    abaConfig.getRange("A1:B1").setValues([["CHAVE DE CONFIGURAÇÃO", "VALORES"]]).setBackground("#4285F4").setFontColor("white").setFontWeight("bold");
    abaConfig.getRange("A2:A3").setValues([["Telefone (ex: 55519...)"], ["ApiKey CallMeBot"]]);
    abaConfig.setColumnWidth(1, 250);
    abaConfig.setColumnWidth(2, 250);
    // Configuração da aba 'Logs'
    var abaLogs = ss.getSheetByName("Logs") || ss.insertSheet("Logs");
    abaLogs.clear();
    abaLogs.getRange("A1:B1").setValues([["DATA/HORA", "RELATÓRIO DE ENVIO"]]).setBackground("#34A853").setFontColor("white").setFontWeight("bold");
    abaLogs.setColumnWidth(1, 150);
    abaLogs.setColumnWidth(2, 500);
    // Configuração da aba 'Envios' (Para E-mails)
    var abaEnvios = ss.getSheetByName("Envios") || ss.insertSheet("Envios");
    if (abaEnvios.getLastRow() === 0) {
        abaEnvios.getRange("A1:C1").setValues([["Nome", "Email", "Status"]]).setBackground("#EA4335").setFontColor("white").setFontWeight("bold");
    }
    ui.alert("✅ Sistema configurado com sucesso!");
}
//Cria o menu personalizado no topo da planilha.
function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('🚀 Painel de Automação').addItem('⚙️ Configurar Planilha (Instalador)', 'configurarMVP').addSeparator().addSubMenu(ui.createMenu('📅 Bot WhatsApp').addItem('📩 Enviar Agenda de Hoje', 'enviarAgendaDoDia').addItem('🗓️ Enviar Agenda da Semana', 'enviarAgendaSemana')).addSubMenu(ui.createMenu('📧 Ferramentas de E-mail').addItem('📤 Enviar E-mails Pendentes', 'enviarEmailsAutomaticos')).addSubMenu(ui.createMenu('🏢 Serviços Externos').addItem('🔍 Consultar CNPJ Selecionado', 'consultarCNPJ')).addSeparator().addItem('🗑️ Limpar Logs', 'limparLogs').addToUi();
}
// --- 2. BOT DE WHATSAPP (AGENDA) ---
function enviarAgendaDoDia() {
    var agora = new Date();
    var inicioDia = new Date(agora.getFullYear(), agora.getMonth(), agora.getDate(), 0, 0, 0);
    var fimDia = new Date(agora.getFullYear(), agora.getMonth(), agora.getDate(), 23, 59, 59);
    var calendario = CalendarApp.getDefaultCalendar();
    var eventos = calendario.getEvents(inicioDia, fimDia);
    var mensagem = "*\uD83D\uDCC5 Agenda de Hoje (" + formatarData(inicioDia) + ")*\n\n";
    if (eventos.length === 0) {
        mensagem += "\uD83D\uDE34 Nenhum compromisso para hoje.";
    }
    else {
        for (var i = 0; i < eventos.length; i++) {
            var evento = eventos[i];
            if (evento.isAllDayEvent()) {
                mensagem += "\uD83D\uDCCC *Dia Inteiro:* " + evento.getTitle() + "\n";
            }
            else {
                mensagem += "\uD83D\uDD52 " + formatarHora(evento.getStartTime()) + " - " + evento.getTitle() + "\n";
            }
            var local = evento.getLocation();
            if (local) mensagem += "   📍 _" + local + "_\n";
        }
    }
    enviarWhatsApp(mensagem);
}

function enviarAgendaSemana() {
    var inicio = new Date();
    var diasSemanaPT = ["Dom", "Seg", "Ter", "Qua", "Qui", "Sex", "Sáb"];
    var mensagem = "*\uD83D\uDCC5 AGENDA DOS PRÓXIMOS 7 DIAS*\n\n";
    var temEvento = false;
    for (var d = 0; d < 7; d++) {
        var dataFoco = new Date();
        dataFoco.setDate(inicio.getDate() + d);
        var evs = CalendarApp.getDefaultCalendar().getEventsForDay(dataFoco);
        if (evs.length > 0) {
            temEvento = true;
            mensagem += "*─── " + diasSemanaPT[dataFoco.getDay()] + ", " + formatarData(dataFoco).substring(0, 5) + " ───*\n";
            evs.forEach(function (e) {
                var icone = e.isAllDayEvent() ? "\uD83D\uDCCC " : "\uD83D\uDD52 " + formatarHora(e.getStartTime()) + " - ";
                mensagem += icone + e.getTitle() + "\n";
            });
            mensagem += "\n";
        }
    }
    if (!temEvento) mensagem += "Sem compromissos na semana.";
    enviarWhatsApp(mensagem);
}
// --- 3. INTEGRAÇÕES (WHATSAPP, EMAIL & CNPJ) ---
function enviarWhatsApp(texto) {
    var config = pegarConfiguracao();
    var url = "https://api.callmebot.com/whatsapp.php?phone=" + config.telefone + "&text=" + encodeURIComponent(texto) + "&apikey=" + config.apiKey;
    try {
        UrlFetchApp.fetch(url);
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Logs").appendRow([new Date(), texto]);
    }
    catch (e) {
        Logger.log("Erro WhatsApp: " + e.message);
    }
}

function enviarEmailsAutomaticos() {
    var aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Envios");
    var dados = aba.getRange(2, 1, aba.getLastRow() - 1, 3).getValues();
    dados.forEach(function (linha, i) {
        if (linha[1] !== "" && linha[2] !== "Enviado") {
            MailApp.sendEmail(linha[1], "Olá " + linha[0], "Segue resumo automático.");
            aba.getRange(i + 2, 3).setValue("Enviado");
        }
    });
}

function consultarCNPJ() {
    var aba = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var cnpj = aba.getActiveCell().getValue().toString().replace(/\D/g, '');
    if (cnpj.length !== 14) return;
    var res = UrlFetchApp.fetch("https://brasilapi.com.br/api/cnpj/v1/" + cnpj);
    var d = JSON.parse(res.getContentText());
    aba.getActiveCell().offset(0, 1).setValue(d.razao_social);
    aba.getActiveCell().offset(0, 2).setValue(d.descricao_situacao_cadastral);
}
// --- 4. UTILITÁRIOS ---
function pegarConfiguracao() {
    var aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Config");
    return {
        telefone: aba.getRange("B2").getValue().toString().replace(/\D/g, '')
        , apiKey: aba.getRange("B3").getValue().toString().trim()
    };
}

function formatarData(d) {
    return Utilities.formatDate(d, "GMT-3", "dd/MM/yyyy");
}

function formatarHora(d) {
    return Utilities.formatDate(d, "GMT-3", "HH:mm");
}

function limparLogs() {
    var aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Logs");
    if (aba.getLastRow() > 1) aba.getRange(2, 1, aba.getLastRow() - 1, 2).clearContent();
}