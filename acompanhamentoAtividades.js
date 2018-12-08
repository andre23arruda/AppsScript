function onOpen() { //Função executada na abertura da planilha
  var plan = SpreadsheetApp.getActive(); // Carregando a planilha
  if(plan.getActiveSheet().getFilter()!= null){ // Se tiver algum filtro na planilha
    plan.getActiveSheet().getFilter().remove(); // Remove o filtro
  }
  var menu = [{name:"Adicionar Linhas", functionName:"addLinhas"},{name:"Adicionar Subtarefa", functionName:"addSub"}, {name:"Criar Atividade", functionName:"createFolder"},{name:"Criar URL", functionName:"doUrl"},{name:"Em andamento", functionName:"addFilter"},
              {name:"Localizar Solicitação", functionName:"procurar1"},{name:"Localizar Observação", functionName:"procurar2"},{name:"Ocultar/Mostrar colunas", functionName:"ocultarColunas"}]; // Criando um objeto com o nome das funções e as funções
  menu.push(null); // LINHA
  menu.push({name:"ABRIR DASHBOARD", functionName:"openDashboard"});
  plan.addMenu("ATALHOS", menu); // Adicionando o menu na barra de tarefas da planilha com o nome ATALHOS
  dimensao3colunas(); // Executando a função que dimensiona as colunas D, E e F com largura de 35 pixels
  ocultarColunas(); // Ocultando ou mostrandos colunas com informações irrelevantes (B, C, V e W)
  Browser.msgBox("Atenção!", "Não renomear ou alterar formatação da planilha, nem adicionar ou remover colunas!", Browser.Buttons.OK); // Mensagem de aviso
}


function openDashboard(){ // Função para gerar o link do dashboard e acessar com apenas um click
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


function doUrl() { // Função para gerar o link de acesso da pasta da solicitacao
  // Essa função só é executada se o elemento selecionada estiver na segunda coluna
  // É possível gerar o link de multiplas pastas
  if(planilhaCorreta()){ // Se for a planilha de Solicitações
    var p1 = SpreadsheetApp.getActive().getActiveSheet(); // Selecionando a planilha geral
    var coluna = p1.getActiveRange().getColumn(); // Verificando a coluna do elemento selecionado
    if(coluna == 2){ // Se a coluna selecionada for a B
      var id1 =  p1.getActiveRange().getValues(); // Pegando o código do elemento selecionado
      var ano = p1.getActiveRange().offset(0,1).getValues(); // Pegando o ano do elemento selecionado
      var tamanho = id1.length; // Tamanho do elemento selecionado
      var name = []; // Iniciando o objeto que receberá toda a formula com hyperlink
      var folder = DriveApp.getFoldersByName("SOLICITACOES").next(); // Acessando a pasta com o nome SOLICITACOES do DRIVE da CIEMA
      name[tamanho-1] = ""; // Dimensionando o objeto com o tamanho do elemento selecionado
      for(i=0;i<tamanho;i++){ // Indo do primeiro até o ultimo elemento selecionado
        try{ // Se der certo esse bloco
          var folder2 = folder.getFoldersByName(ano[i].toString()).next(); // Acessando a pasta com o nome do ano do elemento selecionado
          id1[i] =  "000" + id1[i].toString(); // Colocando 000 no elemento. EX: 1 -> 0001
          name[i] = id1[i].slice(-3) +"." + ano[i].toString().slice(-2);  // Transformando para o nome da pasta. EX: 0001 -> 001.18
          var folder3 = folder2.getFoldersByName(name[i]).next();  // Acessando a pasta da solicitação através do nome acima
          var link = folder3.getUrl(); // Obtendo o link de acesso à pasta
          name[i] = ['= HYPERLINK("'+ link + '";"'+ name[i]+ '")']; // Criando a formula para ser colocada na celula da planilha com link de acesso
        }
        catch(err){ // Se não der certo o bloco acima
          Browser.msgBox('Atenção!', 'A pasta ' + id1[i].slice(-3) +"." + ano[i].toString().slice(-2)  + ' não existe.', Browser.Buttons.OK); // Mensagem de erro
          name[i] = ['= HYPERLINK("'+ ' ' + '";"'+ id1[i].slice(-3) +"." + ano[i].toString().slice(-2)  + '")']; // // Criando a formula para ser colocada na celula da planilha sem link de acesso
        }
      }
      p1.getActiveRange().offset(0,-1).setFormulas(name); // Colocando os links dentro da célula à esquerda do elemento selecionado
      if (tamanho == 1){
        return link;
      }
    }
    else{ // Se a coluna 2 não for selecionada
      Browser.msgBox('Atenção!', 'Por favor, selecione a coluna B para gerar a referência.', Browser.Buttons.OK); // Mensagem de aviso
    }
  }
  else{
    Browser.msgBox('Atenção!', 'Selecione a planilha de Solicitações para executar a função.', Browser.Buttons.OK); // Mensagem de aviso
  }
}


function addSub() { // Função para adicionar subtarefa
  if(planilhaCorreta()){ // Se for a planilha de Solicitações
    var plan = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(); // Selecionando a planilha de Solicitações
    var linhaAtual = plan.getActiveCell().getRow(); // Verificando a linha da célula atual
    var formulaPedidoRespondido = plan.getRange(linhaAtual, 19).getFormula(); // Pegando a formula da coluna S ->   =SE(SAG<>"";"Não";"")
    var nColunas = plan.getLastColumn(); // Pegando o numero de colunas da planilha
    var valores = plan.getRange(linhaAtual,2,1,5).getValues(); // Pegando os valores das colunas B até F
    plan.insertRowAfter(linhaAtual); // Inserindo uma linha abaixo da linha do elemento selecionado
    var cor1 = plan.getRange(linhaAtual, 1).getBackground(); // Pegando a cor da célula do elemento selecionado
    linhaAtual+=1; // Linha atual agora é a linha que foi adicionada
    plan.getRange(linhaAtual,1).setValue("-"); // Como é uma subtarefa, o link é '-' pois é proveniente da tarefa mãe
    plan.getRange(linhaAtual, 2, 1, 5).setValues(valores); // Colocando nas colunas B até F, os mesmo valores contidos na Tarefa mãe
    plan.getRange(linhaAtual,13).setValue("Subsequente"); // Colocando na coluna M (Característica) o tipo Subsequente porque é uma subtarefa
    plan.getRange(linhaAtual,19).setFormula(formulaPedidoRespondido); // Colocando a fórmula da coluna S ->   =SE(SAG<>"";"Não";"")
    if (cor1 == "#dcdcdc"){ // Se a célula do elemento selecionado for cinza claro
      plan.getRange(linhaAtual, 1, 1, nColunas).setBackground("#b7b7b7"); // A linha adicionada será cinza escuro
    }
    else{ // Se não for cinza claro
      plan.getRange(linhaAtual, 1, 1, nColunas).setBackground("#dcdcdc"); // A linha adicionada será cinza claro
    }    
  }
  else{
    Browser.msgBox('Atenção!', 'Selecione a planilha de Solicitações para executar a função.', Browser.Buttons.OK); // Mensagem de aviso
  }
}


function addTarefa() { // Função para adicionar tarefa principal
  if(planilhaCorreta()){ // Se for a planilha de Solicitações
    var plan = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var linhaAtual = plan.getActiveCell().getRow();
    var nColunas = plan.getLastColumn();
    var valores = plan.getRange(linhaAtual,2,1,5).getValues();
    var ano = plan.getRange(linhaAtual,3).getValue();
    var formulaAtualizacao = plan.getRange(linhaAtual, 11).getFormula();
    var formulaPedidoRespondido = plan.getRange(linhaAtual, 19).getFormula(); // Pegando a formula da coluna S ->   =SE(SAG<>"";"Não";"")
    plan.insertRowAfter(linhaAtual);
    linhaAtual+=1;
    plan.getRange(linhaAtual,11).setFormula(formulaAtualizacao);
    plan.getRange(linhaAtual,13).setValue("Principal");
    plan.getRange(linhaAtual,19).setFormula(formulaPedidoRespondido); // Colocando a fórmula da coluna S ->   =SE(SAG<>"";"Não";"")
    plan.getRange(linhaAtual, 1, 1, nColunas).setBackground(null);
  }
  else{
    Browser.msgBox('Atenção!', 'Selecione a planilha de Solicitações para executar a função.', Browser.Buttons.OK); // Mensagem de aviso
  }
}


function addFilter() { // Função para criar filtro das atividades que estão em andamento
  if(planilhaCorreta()){ // Se for a planilha de Solicitações
    var plan = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var ultimaLinha = plan.getLastRow();
    var range = plan.getRange(1,9,ultimaLinha);
    if(range.getFilter() != null){
      range.getFilter().remove();
    }
    var criterio = SpreadsheetApp.newFilterCriteria().setHiddenValues(["Concluído","Encerrado","Anulado","Cancelado"]).build();
    var filtro = range.createFilter().setColumnFilterCriteria(9,criterio);
  }
  else{
    Browser.msgBox('Atenção!', 'Selecione a planilha de Solicitações para executar a função.', Browser.Buttons.OK); // Mensagem de aviso
  }
}


function ocultarColunas() { // Função para ocultar ou mostrar as colunas B, C, V e W
  if(planilhaCorreta()){ // Se for a planilha de Solicitações
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
  else{
    Browser.msgBox('Atenção!', 'Selecione a planilha de Solicitações para executar a função.', Browser.Buttons.OK); // Mensagem de aviso
  }
}


function procurar1() { // Função para filtrar as solicitações que possuem o termo pesquisado
  if(planilhaCorreta()){ // Se for a planilha de Solicitações
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
  else{
    Browser.msgBox('Atenção!', 'Selecione a planilha de Solicitações para executar a função.', Browser.Buttons.OK); // Mensagem de aviso
  }
}


function addLinhas() { // Função para adicionar linhas formatadas
  if(planilhaCorreta()){ // Se for a planilha de Solicitações
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
    sheetAtual.getRange(ultimaLinha+1,1,10,ultimaColuna).setBackground(null);
  }
  else{
    Browser.msgBox('Atenção!', 'Selecione a planilha de Solicitações para executar a função.', Browser.Buttons.OK); // Mensagem de aviso
  }
}


function procurar2() { // Função para filtrar as observações que possuem o termo pesquisado
  if(planilhaCorreta()){ // Se for a planilha de Solicitações
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
  else{
    Browser.msgBox('Atenção!', 'Selecione a planilha de Solicitações para executar a função.', Browser.Buttons.OK); // Mensagem de aviso
  }
}


function createFolder(){ // Função para criar uma nova atividade. Nessa função é criada uma nova pasta referente à nova atividade, além de seu formulário
  if(planilhaCorreta()){ // Se for a planilha de Solicitações
    var p1 = SpreadsheetApp.getActive().getActiveSheet(); // Selecionando a planilha atual
    //var p1 = plan.getSheetByName("Solicitações"); // Selecionando a planilha de Solicitacoes
    var coluna = p1.getActiveRange().getColumn(); // Verificando a coluna do elemento selelecionado
    var nLinhasSelecionadas = p1.getActiveRange().getHeight(); // Verficando o numero de linhas selecionadas
    var nColunasSelecionadas = p1.getActiveRange().getWidth();
    if(coluna == 2 && nLinhasSelecionadas == 1 && nColunasSelecionadas == 1){ // Se a coluna 2 foi selecionada e apenas um elemento
      var linhaAtual = p1.getActiveCell().getRow(); // Guardando a posição da linha do elemento selecionado
      var id1 =  p1.getActiveRange().getValues(); // Pegando o código da atividade
      var ano = p1.getActiveRange().offset(0,1).getValues(); // Pegando o ano da atividade
      if (codAnoCorretos(id1,ano)){ // Se o ano e o código estiverem preenchidos
        var id2 =  "000" + id1.toString();
        var name = id2.slice(-3) +"." + ano.toString().slice(-2); 
        var folder = DriveApp.getFoldersByName("SOLICITACOES").next().getFoldersByName(ano.toString()).next();
        if (folder.getFoldersByName(name).hasNext() == false){
          folder.createFolder(name);
          var files = DriveApp.getFoldersByName("SOLICITACOES").next().getFilesByName("FORMULÁRIO DE ACOMPANHAMENTO DAS ATIVIDADES - NOVO").next();
          files.makeCopy(folder.getFoldersByName(name).next());
          var fileCopy = folder.getFoldersByName(name).next().getFilesByName("Cópia de FORMULÁRIO DE ACOMPANHAMENTO DAS ATIVIDADES - NOVO").next();
          var planForm = SpreadsheetApp.open(fileCopy).getSheetByName('Acompanhamento');
          var date = Utilities.formatDate(new Date(),"GMT",  "yyyy-MM-dd");
          planForm.getRange('B4').setValue(id1);
          planForm.getRange('B5').setValue(ano);
          planForm.getRange('A8').setValue(date);
          planForm.getRange('A8').setNumberFormat("dd/mm/yyyy");
          planForm.getRange('E8').setValue('Início da Atividade');
          planForm.getRange('E8').setHorizontalAlignment('center');
          p1.getRange(linhaAtual,13).setValue("Principal");
          p1.getRange(linhaAtual,16).setValue(date);
          p1.getRange(linhaAtual,16).setNumberFormat("dd/mm/yyyy");
          p1.getRange(linhaAtual, 1, 1, 24).setFontColor("#000000");
          SpreadsheetApp.open(fileCopy).rename("FORMULÁRIO DE ACOMPANHAMENTO DAS ATIVIDADES " + name);
          var link = SpreadsheetApp.open(fileCopy).getUrl();
          doUrl(); // Criando o link da pasta criada
          abrirAtividade(link,name);
        }
        else{
          Browser.msgBox('Atenção!', 'Atividade já existente.', Browser.Buttons.OK);
        }
      }
      else{
        Browser.msgBox('Atenção!', 'Para criar uma nova atividade é necessário que seu código e ano estejam preenchidos.', Browser.Buttons.OK);
      }
    }
    else{
      Browser.msgBox('Atenção!', 'Por favor, selecione apenas uma linha da coluna B para criar a Atividade.', Browser.Buttons.OK);
    }
  }
  else{
    Browser.msgBox('Atenção!', 'Selecione a planilha de Solicitações para executar a função.', Browser.Buttons.OK); // Mensagem de aviso
  }
}


function updateRange(){ // Função para atualizar o intervalo nomeado. Necessário no momento de adicionar linhas formatadas
  var namedRanges = SpreadsheetApp.getActiveSpreadsheet().getNamedRanges(); // Pegando os nomes dos intervalos nomeados
  var plan = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(); // Selecionando a planilha atual
  var nRows = plan.getLastRow(); // Pegando o número de linhas utilizadas na planilha
  for (var i = 0; i < namedRanges.length; i++) { // Indo do primeiro intervalo ao último
    if (namedRanges[i].getName() == "Atualizacao"){ // Se o nome do intervalo for Atualizacao
      var intervalo = plan.getRange(2,22,nRows+10,1); // Aumenta o numero de linhas desse intervalo em 10
      namedRanges[i].setRange(intervalo); // Decretando as novas linhas no intervalo nomeado
    }
    else if (namedRanges[i].getName() == "Link"){ // Se o nome do intervalo for Link
      var intervalo = plan.getRange(2,1,nRows+10,1); // Aumenta o numero de linhas desse intervalo em 10
      namedRanges[i].setRange(intervalo); // Decretando as novas linhas no intervalo nomeado
    }
    else if (namedRanges[i].getName() == "Status2"){ // Se o nome do intervalo for Status2
      var intervalo = plan.getRange(2,9,nRows+10,1); // Aumenta o numero de linhas desse intervalo em 10
      namedRanges[i].setRange(intervalo); // Decretando as novas linhas no intervalo nomeado
    }
    else if (namedRanges[i].getName() == "SAG"){ // Se o nome do intervalo for SAG
      var intervalo = plan.getRange(2,18,nRows+10,1); // Aumenta o numero de linhas desse intervalo em 10
      namedRanges[i].setRange(intervalo); // Decretando as novas linhas no intervalo nomeado
    }
  }
}


function dimensao3colunas(){ // Função que dimensiona as colunas D, E e F com largura de 35 pixels
  var plan = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(); // Selecionando a planilha atual
  plan.setColumnWidths(4, 3, 35) // Colocando 35 pixels nas colunas D, E e F
}


function planilhaCorreta(){ // Função para verificar se a planilha atual é de Solicitações
  var nomePlanilha = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName(); // Pegando o nome da planilha atual
  if (nomePlanilha == 'Solicitações'){ // Se o nome da planilha for Solicitações
    return true; // retorna verdadeiro
  }
  else{ // se o nome da planilha não for Solicitações
    return false;  // retorna falso
  }
}


function codAnoCorretos(cod,ano){ // Função para verificar se o código da atividade e o ano foram preenchidos
  if (cod!="" && ano != ""){ // Se o código e o ano estiverem preenchidos
    return true; // retorna verdadeiro
  }
  else{ // se um dos dois não estiverem preenchidos
    return false; // retorna falso
  }
}


function abrirAtividade(link, nome){ // Função para gerar o link da atividade criada e acessar com apenas um click
  var html = HtmlService.createHtmlOutput('<html><script>'
  +'window.close = function(){window.setTimeout(function(){google.script.host.close()},9)};'
  +'var a = document.createElement("a"); a.href="'+link+'"; a.target="_blank";'
  +'if(document.createEvent){'+'var event=document.createEvent("MouseEvents");'
  // Gambiarra para forçar a janela de interface. Para abrir SEM precisar do link é necessário DESATIVAR o bloqueador de pop-up e RETIRAR o trecho "navigator.userAgent.toLowerCase().indexOf("chrome")>-1" do código abaixo.
  +'if(navigator.userAgent.toLowerCase().indexOf("chrome")>-1||navigator.userAgent.toLowerCase().indexOf("firefox")>-1){window.document.body.append(a)}'
  +'event.initEvent("click",true,true); a.dispatchEvent(event);'+'}else{ a.click() }'+'close();'+'</script>'
  // URL como link clicável caso não abra automaticamente
  +'<body style="word-break:break-word;font-family:sans-serif;"> Ao abrir, aguardar alguns segundos para carregamento da Planilha. <br> Não esqueça de preencher as informações necessárias. <br><br><a href="'+link+'" target="_blank" onclick="window.close()">Clique aqui para prosseguir.</a></body>'
  +'<script>google.script.host.setHeight(150);google.script.host.setWidth(600)</script>'+'</html>').setWidth(150).setHeight(100);
  SpreadsheetApp.getUi().showModalDialog(html,"Gerando link para Formulário da Atividade \n" + nome + " ... ");
}

