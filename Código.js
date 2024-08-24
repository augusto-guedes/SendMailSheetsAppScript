var mailApp = MailApp;
var app = SpreadsheetApp;
var spreadsheet = app.getActiveSpreadsheet();
var sheet = spreadsheet.getSheetByName("Relatório Geral");
var ui = app.getUi();


// Command for put a button in the menu
function onOpen(input) {
  ui.createMenu("Avisos prévio").addItem("Enviar Emails", "sendMails").addToUi();
}


// Command for Send mails.
function sendMails() {
  var data = new Date();
  var values = sheet.getDataRange().getValues();

  // for to compare every line until the last. 
  for (var row = 1; row < values.length; row++) {
    
    // if to check if have dados [row][column] 
    if (typeof values[row][23] === 'object') {
      
      // Valor to set a value in the sheets. [position start from 1. (row, column)] 
      let valor = sheet.getRange(row+1, 29);
      let timeNow = data.getTime();
      let timeCont = values[row][23].getTime();
      
      //Time comparator this number is 2 months in miliseconds.
      if(timeNow<timeCont){
        if ((timeCont-5256000000)<=timeNow) {
          if (values[row][28] == 'Email não enviado')  {
            mailApp.sendEmail(values[row][29],"assunto", "mensagem...");
            valor.setValue('Email enviado')
          }
        } else {
          valor.setValue('Email não enviado')
        }
      }  
      else {
        valor.setValue('Email não enviado')
      } 
    }
  }
}
