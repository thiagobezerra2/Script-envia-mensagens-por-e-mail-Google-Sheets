/***************************
 * EnviaEmail()
 * 
 * Dispara mensagens por e-mail quando tiver
 * Backlog ativo 
 *  
 * 
 *****************************************/



var nomeAba = "RawData"; 
var colunaMonitor = 26; 
var colunaOK = 27;
var linhaInicial = 4;
var email = "thiago.bezerra@uber.com";
var textoOK = 'Enviado';
 
function EnviaEmail(){
  var resultado,linha,data;
  var assunto = "Estáfuncionando a funçaõ";
  var mensagem = "Está funcionando a função e o Victor trabalha do fazendão";
  var aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nomeAba);
  var colunas = aba.getLastColumn();
  
  var timezone = Session.getScriptTimeZone();
  var hoje = Utilities.formatDate(new Date, timezone,'dd MM yyyy')
 
  for(var i=linhaInicial;i<=aba.getLastRow();i++){
    linha = aba.getRange(i,1,1,colunas).getValues();
    data = Utilities.formatDate(linha[0][colunaMonitor-1], timezone, 'dd MM yyyy')
        if(linha[0][colunaOK-1]=='' && data==hoje){
      try {
        MailApp.sendEmail({
          to:email,
          subject:assunto,
          htmlBody:mensagem
        });
        resultado = textoOK;
      } catch(erro) {
        resultado = erro;
      }
      aba.getRange(i,colunaOK).setValue(resultado);
      SpreadsheetApp.flush();
    }
  }
}