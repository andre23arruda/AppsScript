function onOpen() {
  var plan = SpreadsheetApp.getActive();
  if(plan.getActiveSheet().getFilter()!= null){
    plan.getActiveSheet().getFilter().remove();
  }
  var menu = [{name:"Adicionar Linhas", functionName:"addLinhas"},{name:"Adicionar Tarefa", functionName:"addTarefa"},{name:"Adicionar subtarefa", functionName:"addSub"}, {name:"Criar pasta", functionName:"createFolder"},{name:"Criar URL", functionName:"doUrl"},{name:"Em andamento", functionName:"addFilter"},
              {name:"Localizar Solicitação", functionName:"procurar1"},{name:"Localizar Observação", functionName:"procurar2"},{name:"Ocultar/Mostrar colunas", functionName:"ocultarColunas"}];
  menu.push(null); // LINHA
  menu.push({name:"ABRIR DASHBOARD", functionName:"openDashboard"});
  plan.addMenu("ATALHOS", menu);
  ocultarColunas();
  Browser.msgBox("Atenção!", "Não renomear ou alterar formatação da planilha, nem adicionar ou remover colunas!", Browser.Buttons.OK);
}


function openDashboard(){
  // URL específica para cada dashboard
  var urlDashboard = 'https://datastudio.google.com/open/1xQHd3jPFwCpBbFLKrumntvVjmjBBCAOb';
  // Output de um html com um script, para que seja executado dentro da janela de interface de usuário showModalDialog()
  var html = HtmlService.createHtmlOutput('<html><script>'
  +'window.close = function(){window.setTimeout(function(){google.script.host.close()},9)};'
  +'var a = document.createElement("a"); a.href="'+urlDashboard+'"; a.target="_blank";'
  +'if(document.createEvent){'+'var event=document.createEvent("MouseEvents");'
  // Gambiarra para forçar a janela de interface. Para abrir SEM precisar do link é necessário DESATIVAR o bloqueador de pop-up e RETIRAR o trecho "navigator.userAgent.toLowerCase().indexOf("chrome")>-1" do código abaixo.
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
      try{
        var folder2 = folder.getFoldersByName(ano[i].toString()).next();
        id1[i] =  "000" + id1[i].toString();
        name[i] = id1[i].slice(-3) +"." + ano[i].toString().slice(-2); 
        var folder3 = folder2.getFoldersByName(name[i]).next(); 
        var link = folder3.getUrl();
        name[i] = ['= HYPERLINK("'+ link + '";"'+ name[i]+ '")'];    
      }
      catch(err){
        Browser.msgBox('Atenção!', 'A pasta ' + id1[i].slice(-3) +"." + ano[i].toString().slice(-2)  + ' não existe.', Browser.Buttons.OK);
        name[i] = ['= HYPERLINK("'+ ' ' + '";"'+ id1[i].slice(-3) +"." + ano[i].toString().slice(-2)  + '")'];    
      }
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
    var formulaPedidoRespondido = plan.getRange(linhaAtual, 19).getFormula();
    var nColunas = plan.getLastColumn();
    var valores = plan.getRange(linhaAtual,2,1,5).getValues();
    var ano = plan.getRange(linhaAtual,3).getValue();
    plan.insertRowAfter(linhaAtual);
    var cor1 = plan.getRange(linhaAtual, 1).getBackground();
    linhaAtual+=1;
    plan.getRange(linhaAtual,1).setValue("-");
    plan.getRange(linhaAtual, 2, 1, 5).setValues(valores);
    plan.getRange(linhaAtual,13).setValue("Subsequente");
    plan.getRange(linhaAtual,19).setFormula(formulaPedidoRespondido);
    Logger.log(cor1);
    if (cor1 == "#dcdcdc"){
      plan.getRange(linhaAtual, 1, 1, nColunas).setBackground("#b7b7b7");
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
    var formulaPedidoRespondido = plan.getRange(linhaAtual, 19).getFormula();
    plan.insertRowAfter(linhaAtual);
    linhaAtual+=1;
    plan.getRange(linhaAtual,11).setFormula(formulaAtualizacao);
    plan.getRange(linhaAtual,13).setValue("Principal");
    plan.getRange(linhaAtual,19).setFormula(formulaPedidoRespondido);
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
    plan.showColumns(2, 2);
    plan.showColumns(22, 2);
  }
  else{
    plan.hideColumns(2, 2);
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

  sheetAtual.insertRowsAfter(ultimaLinha, 10);
  sheetAtual.getRange(ultimaLinha,1,ultimaLinha+10,ultimaColuna).setBorder(true ,true ,true ,true ,true,true);
  updateRange();
  var formulas =  [['' ,'' ,'' ,'' ,'' ,'' ,'' ,'' ,'' ,'' , '=IF(Link="","",(IF(OR(Status2="Concluído",Status2="Encerrado",Status2="Cancelado",Status2="Anulado"),"Finalizado",(IF(PrazoAtualizacao-Atualizacao<30,"Atualizado","Desatualizado")))))','' ,'' ,'' ,'' ,'' ,'' ,'' , '=IF(SAG<>"","Não","")','' ,'' ,'' ,'' ,'']];
  for (i=1;i<10;i++){
    formulas[i] = formulas[0];
  }
  sheetAtual.getRange(ultimaLinha+1,1,10,ultimaColuna).setFormulas(formulas);
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


function createFolder(){
  var plan = SpreadsheetApp.getActive();
  var p1 = plan.getSheetByName("Solicitações");
  var coluna = p1.getActiveRange().getColumn();
  var nLinhasSelecionadas = p1.getActiveRange().getNumRows();    
  if(coluna == 2 && nLinhasSelecionadas == 1){
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
  else{
    Browser.msgBox('Atenção!', 'Por favor, selecione apenas uma linha da coluna B para criar a pasta.', Browser.Buttons.OK);
  }
}


function updateRange(){
  var namedRanges = SpreadsheetApp.getActiveSpreadsheet().getNamedRanges();
  var plan = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var nRows = plan.getLastRow();
  for (var i = 0; i < namedRanges.length; i++) {
    if (namedRanges[i].getName() == "Atualizacao"){
      var intervalo = plan.getRange(2,22,nRows+10,1);
      namedRanges[i].setRange(intervalo);
    }
    else if (namedRanges[i].getName() == "Link"){
      var intervalo = plan.getRange(2,1,nRows+10,1);
      namedRanges[i].setRange(intervalo);
    }
    else if (namedRanges[i].getName() == "Status2"){
      var intervalo = plan.getRange(2,9,nRows+10,1);
      namedRanges[i].setRange(intervalo);
    }
    else if (namedRanges[i].getName() == "SAG"){
      var intervalo = plan.getRange(2,18,nRows+10,1);
      namedRanges[i].setRange(intervalo);
    }
  }
}


