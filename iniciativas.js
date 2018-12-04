function onOpen() {  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mapa = ss.getSheetByName("Mapa das Iniciativas");
  var n_iniciativas = mapa.getRange("B5").getValue() -1;
  Browser.msgBox("Atenção!", "Ao atualizar uma informação, esperar alguns segundos por favor.", Browser.Buttons.OK);
  for(var i = 6; i<=n_iniciativas+6; i++){
    if (mapa.getRange(i,15).getValue() == "Sim" || mapa.getRange(i,10).getValue() == "CONCLUÍDA"){
      mapa.hideRows(i);
    }
  }
  var n_planilhas = ss.getSheets().length;
  for(var i = 0; i<n_planilhas; i++){
    var planilha = ss.getSheets()[i];
    if(planilha.getSheetName()!="Mapa das Iniciativas" && planilha.getSheetName().split(" ")[0] != "Iniciativa"){
      planilha.hideSheet();
    }
 }
  var menu = [{name:"Adicionar iniciativa", functionName:"addIniciativa"}];
  ss.addMenu("ADICIONAR INICIATIVA", menu);
  
  var menu2 = [{name:"Adicionar tarefa", functionName:"addTarefa"}];
  ss.addMenu("ADICIONAR TAREFA", menu2);
  
}


function onEdit(e){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mapa = ss.getSheetByName("Mapa das Iniciativas");
  
  var d = ss.getActiveSheet(); 
  

  Logger.log((d.getSheetId()).toFixed());
  Logger.log(ss.getUrl()+"#gid="+(d.getSheetId()).toFixed());

  var intervalo_teste = d.getActiveRange();
  var x = intervalo_teste.getColumn();
  var y = intervalo_teste.getRow();
  intervalo_teste = d.getRange(y,x).getA1Notation();
  
  var valores = d.getActiveRange().getValues();
    
  if (intervalo_teste != "B17" && intervalo_teste != "B1" ){ 
    var masterSheet = ss.getActiveSheet();
    var name_masterSheet = masterSheet.getName();
    var masterRange = masterSheet.getActiveRange();
    var intervalo = masterRange.getA1Notation();
    var n_colunas = masterRange.getNumColumns();
    var n_linhas = masterRange.getNumRows();
    var referencia = masterSheet.getRange(intervalo).offset(0,1).getValue();
  
    if (name_masterSheet.split(" ")[0] != "Cópia"){
      var helperSheet = ss.getSheetByName("Cópia de "+ name_masterSheet);
      var coluna = masterRange.getColumn();
      var linha = masterRange.getRow();
      var permissao = masterSheet.getRange("A41").getValue();

      if(permissao == "Não"){
        edicao(ss,mapa,masterSheet,helperSheet,coluna,linha,intervalo,n_colunas,n_linhas,referencia);
        masterSheet.getRange("B17").copyTo(helperSheet.getRange("B17"), {contentsOnly:false});
      }
      else{
        masterSheet.getRange(intervalo).copyTo(helperSheet.getRange(intervalo), {contentsOnly:false});
        masterSheet.getRange("B17").copyTo(helperSheet.getRange("B17"), {contentsOnly:false});
      }
    }
    else{
        var helperSheet = ss.getSheetByName(name_masterSheet.slice(9,name_masterSheet.length));
        helperSheet.getRange(intervalo).copyTo(masterSheet.getRange(intervalo), {contentsOnly:false});
        Browser.msgBox("Atenção!", "Você não pode desfazer ou editar a Planilha de Cópia.", Browser.Buttons.OK);
        masterSheet.hideSheet();
    }
  }
}


function edicao(ss,mapa,masterSheet,helperSheet,coluna,linha,intervalo,n_colunas,n_linhas,referencia){
  var n_tarefas = masterSheet.getRange("A6").getValue()+1;
  Logger.log(n_tarefas)

   if(n_colunas + n_linhas == 2 && linha != 1 && coluna == 5 && masterSheet.getSheetName() != "Mapa das Iniciativas" && referencia != "-" && linha<=n_tarefas){
      var conclusao = masterSheet.getActiveRange().getValue();
      if (typeof conclusao != "number"){
        helperSheet.getRange(intervalo).copyTo(masterSheet.getRange(intervalo),{contentsOnly:false});
        Browser.msgBox("Atenção!", "Por favor, insira um valor válido.", Browser.Buttons.OK);
      }
      else{
        if(conclusao >= 1){
          masterSheet.getActiveRange().setValue(1);
          var data_termino = Utilities.formatDate(new Date, "GMT", "yyyy-MM-dd");
          var celula_termino = masterSheet.getRange(linha,10);
          if (celula_termino.getValue == "-" || celula_termino.isBlank() == true){
            celula_termino.setValue(data_termino);
            celula_termino.setNumberFormat("dd/mm/yyyy");
            var celula_termino = helperSheet.getRange(linha,10);
            celula_termino.setValue(data_termino);
            celula_termino.setNumberFormat("dd/mm/yyyy");
          }
        } 
        var r = get_used_rows(masterSheet, 12);
        var data = masterSheet.getRange(r,12); 
        var tempo_anterior = data.getValue();
        if(r != 1){
          var tempo_anterior = Utilities.formatDate(tempo_anterior, "GMT", "yyyy-MM");
        }
        var tempo_atual = Utilities.formatDate(new Date, "GMT", "yyyy-MM");
        
        
        var celula_teste = masterSheet.getRange(1,15);
        celula_teste.setFontColor("white");
        
        var intervalo_conclusao = "E2:E";
        var intervalo_duracao = "G2:G";
        
        var formula = '=(SUMIF(ARRAYFORMULA('+intervalo_conclusao+'*'+intervalo_duracao+');">=0"))/SUM('+intervalo_duracao+')';
        
        celula_teste.setFormula(formula);
        
        var concluida_total = masterSheet.getRange(1,15).getValue();
        celula_teste.clear();
        
        if(tempo_anterior != tempo_atual){
          var data = masterSheet.getRange(r+1,12);
          data.setValue(tempo_atual+"-15");
          data.setNumberFormat("mm/yyyy");
          var c_total = masterSheet.getRange(r+1,14);
          c_total.setValue(concluida_total);
          c_total.offset(0,-1).setValue("-");
          c_total.offset(0,-1).setHorizontalAlignment("center");
          c_total.offset(0,-1).setVerticalAlignment("middle");
          r = r+1;     
        }
        else{
          var data = masterSheet.getRange(r,12);
          data.setValue(tempo_atual+"-15");
          data.setNumberFormat("mm/yyyy");
          var c_total = masterSheet.getRange(r,14);
          c_total.setValue(concluida_total);
        }
        c_total.setNumberFormat("00.00%");
        masterSheet.getRange(linha,coluna).copyTo(helperSheet.getRange(linha,coluna), {contentsOnly:false});
        masterSheet.getRange(r,12,1,3).copyTo(helperSheet.getRange(r,12,1,3), {contentsOnly:false});
              
        var id = Number(masterSheet.getName().split(" ").pop());
        var cell_conclusao = mapa.getRange(id+5,6);
        cell_conclusao.setNumberFormat("00.00%");
        cell_conclusao.setValue(concluida_total);
        var data_atualizacao = Utilities.formatDate(new Date, "GMT", "yyyy-MM-dd");
        cell_conclusao.offset(0, 8).setValue(data_atualizacao);
        
        var mapa_copia = ss.getSheetByName("Cópia de Mapa das Iniciativas");
        mapa.getRange(id+5,6,1,9).copyTo(mapa_copia.getRange(id+5,6,1,9), {contentsOnly:false});
        mapa_copia.getRange("B17").copyTo(mapa.getRange("B17"), {contentsOnly:false});
      }
    }
    else if(coluna == 13 && linha != 1 && masterSheet.getSheetName() != "Mapa das Iniciativas"){
      var n_data_atualizacao = get_used_rows(masterSheet, 12);
      if (linha!=n_data_atualizacao){
        helperSheet.getRange(intervalo).copyTo(masterSheet.getRange(intervalo),{contentsOnly:false});
        Browser.msgBox("Atenção!", "Não é possível alterar essa análise crítica.", Browser.Buttons.OK);
      }
      else{
        var data_atualizacao = Utilities.formatDate(new Date, "GMT", "yyyy-MM-dd");
        
        var mapa = ss.getSheetByName("Mapa das Iniciativas");
        var id = Number(masterSheet.getName().split(" ").pop());
        var cell_atualizacao = mapa.getRange(id+5,14);
        cell_atualizacao.setValue(data_atualizacao);
        
        var mapa = ss.getSheetByName("Cópia de Mapa das Iniciativas");
        var cell_atualizacao = mapa.getRange(id+5,14);
        cell_atualizacao.setValue(data_atualizacao);
      }
    }
    else if(coluna == 1 && linha == 41){
    } 
    else{
      helperSheet.getRange(intervalo).copyTo(masterSheet.getRange(intervalo),{contentsOnly:false});
      Browser.msgBox("Atenção!", "Não é possível editar este campo.", Browser.Buttons.OK);
      masterSheet.getRange("B17").copyTo(helperSheet.getRange("B17"), {contentsOnly:false});
      if (masterSheet.getRange("A42").getValue() == 1){
        hue1();
      }
    }
  }


function get_used_rows(sheet, column_index){
      for (var r = sheet.getLastRow()+1; r--; r > 1) {
        var valor = sheet.getRange(r,column_index).getValue(); 
        if (valor != ""){
          return r;
          break;
        }
      }
}

function hue1(){ 
//  var ss = SpreadsheetApp.getActiveSpreadsheet();
//  var masterSheet = ss.getActiveSheet();
//  var nome_anterior = masterSheet.getName();
//  ss.getSheetByName("ZUERA").activate();
//  var sheet_zuera = ss.getActiveSheet();
//  var range = sheet_zuera.getRange(1,1);
//  range.setValue(nome_anterior);
}

function voltar(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet_zuera = ss.getActiveSheet();
  var nome_anterior = sheet_zuera.getRange(1,1).getValue();
  ss.getSheetByName("Mapa das Iniciativas").activate();
  sheet_zuera.hideSheet();
}

  
function addIniciativa(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var sheetName = sheet.getSheetName();
  var ui = SpreadsheetApp.getUi();
  if (sheetName == 'Mapa das Iniciativas' && sheet.getRange("A41").getValue() == "Sim"){
    var botao = ui.prompt('Qual iniciativa deseja incluir?',ui.ButtonSet.YES_NO);
    var nomeIniciativa = botao.getResponseText();
    var iniciativa = ss.getSheetByName(nomeIniciativa);
    if (botao.getSelectedButton() == ui.Button.YES){
      var nUltimaLinha = get_used_rows(sheet, 3);
      var ultimaLinha = sheet.getRange(nUltimaLinha, 3, 1, 14);
      var novaLinha = sheet.getRange(nUltimaLinha+1, 3, 1, 14);
      var tarefas = iniciativa.getRange("E2:E").getValues();
      var l = 1;
      while (true){
        l = l+1;
        if(tarefas != '-%'){
          break
        }
      }      
      var inicioIniciativa = iniciativa.getRange(l,6).getValue();
      var prazoIniciativa = iniciativa.getRange(get_used_rows(iniciativa, 3),8).getValue();
      
      ultimaLinha.copyTo(novaLinha);
      novaLinha.getCell(1,1).setValue(1+Number(novaLinha.getCell(1,1).getValue()));
      novaLinha.getCell(1,2).setValue(iniciativa.getRange("A4").getValue());
      //Logger.log('=HIPERLINK("'+ ss.getUrl()+"#gid="+(sheet.getSheetId()).toFixed()+'");"'+(1+Number(novaLinha.getCell(1,1).getValue())) +'")')
      novaLinha.getCell(1,3).setValue(iniciativa.getRange("A1").getValue());
      novaLinha.getCell(1,4).setValue(iniciativa.getRange(get_used_rows(iniciativa, 14),14).getValue());
      novaLinha.getCell(1,5).setValue(inicioIniciativa);
      novaLinha.getCell(1,6).setValue(prazoIniciativa);
    }
  }
  else{
    Browser.msgBox("Atenção!", "Não é possível inserir iniciativa. Entrar em contato com o Sr. Diego.", Browser.Buttons.OK);
  }
}


function addTarefa(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var sheetName = sheet.getSheetName().split();
  if (sheetName[0] = "Iniciativa" && sheet.getRange("A41").getValue() == "Sim"){
    Browser.msgBox("Atenção!", "Para adicionar uma tarefa, preencha apenas os campos: Nome da Tarefa, Início e Duração. Ao término, entrar em contato com o Sr. Diego.", Browser.Buttons.OK);
   var nTarefas = sheet.getRange("A6").getValue();
   var ultimaTarefa = sheet.getRange(nTarefas+1, 3,1,8);
   var novaTarefa =  sheet.getRange(nTarefas+2, 3,1,8);
   ultimaTarefa.copyTo(novaTarefa, {contentsOnly:false});
   var valores = [[nTarefas+1,'-','-','-']];
   novaTarefa =  sheet.getRange(nTarefas+2, 3,1,4);
   novaTarefa.setValues(valores);
   sheet.getRange(nTarefas+2,5).setNumberFormat("00.00%");
   sheet.getRange(nTarefas+2,7).setValue(""); 
   sheet.getRange(nTarefas+2,7).setNumberFormat("00");
   sheet.getRange(nTarefas+2,10).setValue(""); 
    
    
   novaTarefa =  sheet.getRange(nTarefas+2, 3,1,8);
   var planCopia = ss.getSheetByName("Cópia de " + sheet.getSheetName());
   var novaTarefaCopia = planCopia.getRange(nTarefas+2, 3,1,8);
   novaTarefa.copyTo(novaTarefaCopia, {contentsOnly:false}); 
  }
  else{
    Browser.msgBox("Atenção!", "Não é possível inserir tarefa. Entrar em contato com o Sr. Diego.", Browser.Buttons.OK);
  }
  
}
  
