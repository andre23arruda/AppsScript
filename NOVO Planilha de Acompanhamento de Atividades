function onOpen() {
  var plan = SpreadsheetApp.getActive();
  if(plan.getActiveSheet().getFilter()!= null){
    plan.getActiveSheet().getFilter().remove();
  }
  Browser.msgBox("Atenção!", "Não renomear ou alterar formatação da planilha, nem adicionar ou remover colunas!", Browser.Buttons.OK);
  var menu = [{name:"Adicionar Linhas", functionName:"addLinhas"},{name:"Adicionar Tarefa", functionName:"addTarefa"},{name:"Adicionar subtarefa", functionName:"addSub"}, {name:"Criar pasta", functionName:"createFolder"},{name:"Criar URL", functionName:"doUrl"},{name:"Em andamento", functionName:"addFilter"},
              {name:"Localizar Solicitação", functionName:"procurar1"},{name:"Localizar Observação", functionName:"procurar2"},{name:"Ocultar/Mostrar colunas", functionName:"ocultarColunas"}];
  menu.push(null); // LINHA
  menu.push({name:"ABRIR DASHBOARD", functionName:"openDashboard"});
  plan.addMenu("ATALHOS", menu);
  ocultarColunas();
}

function openDashboard(){
  // URL específica para cada dashboard
  var urlDashboard = 'https://datastudio.google.com/open/1xQHd3jPFwCpBbFLKrumntvVjmjBBCAOb';
  // Output de um html com um script
  var html = HtmlService.createHtmlOutput('<html><script>'
  +'window.close = function(){window.setTimeout(function(){google.script.host.close()},9)};'
  +'var a = document.createElement("a"); a.href="'+urlDashboard+'"; a.target="_blank";'
  +'if(document.createEvent){'+'var event=document.createEvent("MouseEvents");'
  //  Condição para "ignorar" bloqueador de Pop-up (NOT REALLY... é gambs)
  +'if(navigator.userAgent.toLowerCase().indexOf("chrome")>-1||navigator.userAgent.toLowerCase().indexOf("firefox")>-1){window.document.body.append(a)}'
  +'event.initEvent("click",true,true); a.dispatchEvent(event);'+'}else{ a.click() }'+'close();'+'</script>'
  // URL como link clicável caso não abra automaticamente
  +'<body style="word-break:break-word;font-family:sans-serif;"><a href="'+urlDashboard+'" target="_blank" onclick="window.close()">Clique aqui para prosseguir.</a></body>'
  +'<script>google.script.host.setHeight(35);google.script.host.setWidth(280)</script>'+'</html>').setWidth(90).setHeight(1);
  SpreadsheetApp.getUi().showModalDialog(html,"Gerando link do Dashboard...");
}

function doUrl() {
  var plan = SpreadsheetApp.getActive();
  var p1 = plan.getSheetByName("Solicitações");
  var coluna = p1.getActiveRange().getColumn();
  if(coluna == 2){
    var id1 =  p1.getActiveRange().getValues();
    var ano = p1.getActiveRange().offset(0,1).getValues();
    var tamanho = id1.length;
    var name = [];
    var folder = DriveApp.getFoldersByName("SOLICITACOES").next(); 
    name[tamanho-1] = "";
    for(i=0;i<tamanho;i++){
      var folder2 = folder.getFoldersByName(ano[i].toString()).next();
      id1[i] =  "000" + id1[i].toString();
      name[i] = id1[i].slice(-3) +"." + ano[i].toString().slice(-2); 
      var folder3 = folder2.getFoldersByName(name[i]).next(); 
      var link = folder3.getUrl();
      name[i] = ['= HYPERLINK("'+ link + '";"'+ name[i]+ '")'];     
    }
    p1.getActiveRange().offset(0,-1).setFormulas(name);
  }
  else{
    Browser.msgBox('Atenção!', 'Por favor, selecione a coluna B para gerar a referência.', Browser.Buttons.OK);
  }
}


function addSub() {
  var plan = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (plan.getSheetName() == "Solicitações"){
    var linhaAtual = plan.getActiveCell().getRow();
    var nColunas = plan.getLastColumn();
    var valores = plan.getRange(linhaAtual,2,1,5).getValues();
    var ano = plan.getRange(linhaAtual,3).getValue();
    plan.insertRowAfter(linhaAtual);
    var cor1 = plan.getRange(linhaAtual, 1).getBackground();
    linhaAtual+=1;
    plan.getRange(linhaAtual,1).setValue("-");
    plan.getRange(linhaAtual, 2, 1, 5).setValues(valores);
    plan.getRange(linhaAtual,13).setValue("Subsequente");
    Logger.log(cor1);
    if (cor1 == "#dcdcdc"){
      plan.getRange(linhaAtual, 1, 1, nColunas).setBackground("#999999");
    }
    else{
      plan.getRange(linhaAtual, 1, 1, nColunas).setBackground("#dcdcdc");
    }    
  }
}


function addTarefa() {
  var plan = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (plan.getSheetName() == "Solicitações"){
    var linhaAtual = plan.getActiveCell().getRow();
    var nColunas = plan.getLastColumn();
    var valores = plan.getRange(linhaAtual,2,1,5).getValues();
    var ano = plan.getRange(linhaAtual,3).getValue();
    var formulaAtualizacao = plan.getRange(linhaAtual, 11).getFormula();
    plan.insertRowAfter(linhaAtual);
    linhaAtual+=1;
    plan.getRange(linhaAtual,11).setFormula(formulaAtualizacao);
    plan.getRange(linhaAtual,13).setValue("Principal");
    plan.getRange(linhaAtual, 1, 1, nColunas).setBackground(null);
  }
}


function addFilter() {
  var plan = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var ultimaLinha = plan.getLastRow();
  var range = plan.getRange(1,9,ultimaLinha);
  if(range.getFilter() != null){
    range.getFilter().remove();
  }
  var criterio = SpreadsheetApp.newFilterCriteria().setHiddenValues(["Concluído","Encerrado","Anulado","Cancelado"]).build();
  var filtro = range.createFilter().setColumnFilterCriteria(9,criterio);
}


function ocultarColunas() {
  var plan = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (plan.isColumnHiddenByUser(2)){
    plan.showColumns(3, 1);
    plan.showColumns(22, 2);
  }
  else{
    plan.hideColumns(3, 1);
    plan.hideColumns(22, 2);
  }
}


function procurar1() {
  var plan = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var ultimaLinha = plan.getLastRow();
  var range = plan.getRange(1,7,ultimaLinha);
  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt(
      'Localizar Solicitação',
      'Por favor, digite o termo que deseja procurar:',
      ui.ButtonSet.OK_CANCEL);
  var button = result.getSelectedButton();
  var text = result.getResponseText();
  if (text != ""){
    if(range.getFilter() != null){
    range.getFilter().remove();
  }
  var criterio = SpreadsheetApp.newFilterCriteria().whenTextContains(text).build();
  range.createFilter().setColumnFilterCriteria(7,criterio);  
  }
}

function addLinhas() {
  var sheetAtual = SpreadsheetApp.getActive().getActiveSheet();
  
  var ultimaLinha = sheetAtual.getLastRow();
  var ultimaColuna = sheetAtual.getLastColumn();

  sheetAtual.insertRowsAfter(ultimaLinha, 50);
  sheetAtual.getRange(1,1,ultimaLinha+50,ultimaColuna).setBorder(true ,true ,true ,true ,true,true);
}

function procurar2() {
  var plan = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var ultimaLinha = plan.getLastRow();
  var range = plan.getRange(1,10,ultimaLinha);
  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt(
      'Localizar Observação',
      'Por favor, digite o termo que deseja procurar:',
      ui.ButtonSet.OK_CANCEL);
  var button = result.getSelectedButton();
  var text = result.getResponseText();
  if (text != ""){
    if(range.getFilter() != null){
    range.getFilter().remove();
  }
  var criterio = SpreadsheetApp.newFilterCriteria().whenTextContains(text).build();
  range.createFilter().setColumnFilterCriteria(10,criterio);  
  }
}

function addLinhas() {
  var sheetAtual = SpreadsheetApp.getActive().getActiveSheet();
  
  var ultimaLinha = sheetAtual.getLastRow();
  var ultimaColuna = sheetAtual.getLastColumn();

  sheetAtual.insertRowsAfter(ultimaLinha, 50);
  sheetAtual.getRange(1,1,ultimaLinha+50,ultimaColuna).setBorder(true ,true ,true ,true ,true,true);
}


function createFolder(){
  var plan = SpreadsheetApp.getActive();
  var p1 = plan.getSheetByName("Solicitações");
  var coluna = p1.getActiveRange().getColumn();
  if(coluna == 2){
    var id1 =  p1.getActiveRange().getValues();
    var ano = p1.getActiveRange().offset(0,1).getValues();
    var id2 =  "000" + id1.toString();
    var name = id2.slice(-3) +"." + ano.toString().slice(-2); 
    var folder = DriveApp.getFoldersByName("SOLICITACOES").next().getFoldersByName(ano.toString()).next();
    if (folder.getFoldersByName(name).hasNext() == false){
      folder.createFolder(name);
      var files = DriveApp.getFoldersByName("SOLICITACOES").next().getFilesByName("FORMULÁRIO DE ACOMPANHAMENTO DAS ATIVIDADES - NOVO").next();
      files.makeCopy(folder.getFoldersByName(name).next());
      var fileCopy = folder.getFoldersByName(name).next().getFilesByName("Cópia de FORMULÁRIO DE ACOMPANHAMENTO DAS ATIVIDADES - NOVO").next();
      var planForm = SpreadsheetApp.open(fileCopy).getSheetByName('Acompanhamento');
      planForm.getRange('B4').setValue(id1);
      planForm.getRange('B5').setValue(ano);
      SpreadsheetApp.open(fileCopy).rename("FORMULÁRIO DE ACOMPANHAMENTO DAS ATIVIDADES " + name);
      doUrl();
      Browser.msgBox('Concluído!', 'Pasta criada com sucesso', Browser.Buttons.OK);
    }
    else{
      Browser.msgBox('Atenção!', 'Pasta já existente.', Browser.Buttons.OK);
    }
  }
}



