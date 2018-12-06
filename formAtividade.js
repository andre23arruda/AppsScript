function onOpen() {
  var plan = SpreadsheetApp.getActive();
  Browser.msgBox("Atenção!", "Não alterar formatação da planilha, nem adicionar ou remover colunas!", Browser.Buttons.OK);
  var menu = [{name:"Atualizar", functionName:"atualizar"}];
  plan.addMenu("ATALHOS", menu);
}

function atualizar(){
  var ss1 = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var codigo1 = ss1.getRange('B4').getValue();
  var ano1 = ss1.getRange('B5').getValue();
  var responsavel1 = ss1.getRange('C3').getValue();
  var status1 = ss1.getRange('C6').getValue();
  var solicitacao = ss1.getRange('C2').getValue();
  var obs = ss1.getRange('E4').getValue();
  
  codigo1 = parseInt(codigo1);
  ano1 = parseInt(ano1);
  
  var lastRow = ss1.getLastRow();
  var dataAtualizacao = ss1.getRange(lastRow, 1).getValue();
  if (dataAtualizacao == "Data"){
    dataAtualizacao = "";
  }
  
  var file = DriveApp.getFilesByName('NOVO Planilha de Acompanhamento de Atividades').next();
  var ss2 = SpreadsheetApp.open(file).getSheetByName('Solicitações');
  var lastRow2 = ss2.getLastRow();
  
  var i = -1;
  var codigo2;
  var ano2;
  var cod_e_ano = ss2.getRange(2,2,lastRow2+1,2).getValues();
  
  while (true){
    i++;
    codigo2 = cod_e_ano[i][0];
    ano2 = cod_e_ano[i][1];
    codigo2 = parseInt(codigo2);
    ano2 = parseInt(ano2);
    if (codigo1 == codigo2 && ano1 == ano2){
      ss2.getRange(i+2,22).setValue(dataAtualizacao);
      ss2.getRange(i+2, 7, 1, 4).setValues([[solicitacao,responsavel1,status1,obs]]);
      if (responsavel1 == "Amaro"){
        ss2.getRange(i+2, 1, 1, 24).setFontColor("#6aa84f");
      }
      else if(responsavel1 == "Josimar"){
        ss2.getRange(i+2, 1, 1, 24).setFontColor("#a64d79");
      }
      else if(responsavel1 == "Paulo"){
        ss2.getRange(i+2, 1, 1, 24).setFontColor("#1155cc");
      }
      else{
        ss2.getRange(i+2, 1, 1, 24).setFontColor("#000000");
      }
      if (status1 == "Concluído"){
        ss2.getRange(i+2, 17).setValue(dataAtualizacao);
      }
      break;
    }
  } 
}

