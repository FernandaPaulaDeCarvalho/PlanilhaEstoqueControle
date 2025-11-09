/**
 * Nome do Projeto: ScriptBotaoAmareloGravar
 * Descrição: Este script se refere a automação dos botões GravarEntrada e GravarSaida, e complementa a planilha EstoqueControle. 
 *            
 * Repositório: [https://github.com/FernandaPaulaDeCarvalho/PlanilhaEstoqueControle]
 * 
 * Copyright (c) 2025 Fernanda Paula de Carvalho
 *
 * Este código está licenciado sob a Licença MIT.
 * Consulte o arquivo LICENSE no repositório para obter mais informações.
 * 
 * IMPORTANTE:
 * 1. Este script contém as funções gravarEntrada e gravarSaida, que são necessárias para o 
 *    funcionamento da planilha EstoqueControle.
 * 2. Este script deve ser copiado e colado no editor de script da planilha EstoqueControle recém criada. 
 * 3. Este script requer permissões para acessar planilhas no Google Sheets.
 *    Ao executar o script pela primeira vez, você será solicitado a conceder essas permissões.
 * 4. Este script depende do Google App Script e do GitHub e pode parar de funcionar se houver
 *    mudanças nas plataformas.
 *
 * ISENÇÃO DE RESPONSABILIDADE:
 * A autora deste script não se responsabiliza por quaisquer problemas ou danos que possam
 * ocorrer devido ao uso deste script.
*/

/** ATUALIZAR SEMPRE QUE HOUVER MODIFICAÇÃO */
// Definir versão atual, data de modificação, link do repositório do script
const versaoAtual = "1.0.0"; 
const dataModificacao = "2025-11-08";


function gravarEntrada() {

  //Abrir a aba ENTRADA e procurar se há algum PENDENTE
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var entrada = ss.getSheetByName('ENTRADA');
  var procurarPendente = entrada.getRange('K3:K').createTextFinder('PENDENTE').findAll();

 
  //Se encontrar algum PENDENTE, exibir alerta e interromper script 
  if (procurarPendente.length > 0) {
    var ui = SpreadsheetApp.getUi();
    var alertaPendente = ui.alert("PENDENTE(s)! Verifique os dados e tente gravar novamente!",ui.ButtonSet.OK);
    if (alertaPendente == ui.Button.OK) { return }

  }


  //Se não encontrar PENDENTE, perguntar "Deseja retirar fórmulas e gravar ENTRADA?" 
  else if (procurarPendente.length == 0) {
    var ui = SpreadsheetApp.getUi();
    var perguntaGravar = ui.alert("Deseja retirar fórmulas e gravar ENTRADA?",ui.ButtonSet.YES_NO);

    //Se for não, interromper script
    if (perguntaGravar == ui.Button.NO) { return } 

    //Se for sim, gravação das entradas: procurar linhas contendo <<GRAVAR>> 
    else if (perguntaGravar == ui.Button.YES) { 
  
      var entrada = ss.getSheetByName('ENTRADA');
      var gravar = entrada.getRange('K3:K').createTextFinder('<<GRAVAR>>');
      var procurarGravar = gravar.findAll(); 

      // Processar as linhas encontradas em blocos para otimizar o desempenho
      const blockSize = 100; // Tamanho do bloco
      for (let i = 0; i < procurarGravar.length; i += blockSize) {
        const block = procurarGravar.slice(i, Math.min(i + blockSize, procurarGravar.length));

        // Processar o bloco de linhas
        block.forEach(cell => {
          // Listar intervalos K*:M* contendo <<GRAVAR>>
          var intervaloGravar = "K"+cell.getRow()+":M"+cell.getRow();
          var listaIntervaloGravar = entrada.getRange(intervaloGravar);

          // Remover fórmulas dos intervalos <<GRAVAR>>
          listaIntervaloGravar.copyTo(listaIntervaloGravar,{contentsOnly:true});

          // Substituir <<GRAVAR>> por INSERIDO
          gravar.replaceAllWith('INSERIDO');
        });

        // Executar todas as alterações do loop antes de passar para o próximo bloco
        SpreadsheetApp.flush();

      }

    }

    //Mensagem final "ENTRADA concluída!"
    var mensagemFinal = ui.alert("ENTRADA concluída!",ui.ButtonSet.OK);
    if (mensagemFinal == ui.Button.OK) { return };
  
  }
    
}

function gravarSaida() {

  //Abrir a aba SAÍDA e procurar se há algum PENDENTE
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var saida = ss.getSheetByName('SAÍDA');
  var procurarPendente = saida.getRange('I3:I').createTextFinder('PENDENTE').findAll();

 
  //Se encontrar algum PENDENTE, exibir alerta e interromper script 
  if (procurarPendente.length > 0) {
    var ui = SpreadsheetApp.getUi();
    var alertaPendente = ui.alert("PENDENTE(s)! Verifique os dados e tente gravar novamente!",ui.ButtonSet.OK);
    if (alertaPendente == ui.Button.OK) { return }

  }


  //Se não encontrar PENDENTE, perguntar "Deseja retirar fórmulas e gravar SAÍDA?" 
  else if (procurarPendente.length == 0) {
    var ui = SpreadsheetApp.getUi();
    var perguntaGravar = ui.alert("Deseja retirar fórmulas e gravar SAÍDA?",ui.ButtonSet.YES_NO);

    //Se for não, interromper script
    if (perguntaGravar == ui.Button.NO) { return } 
    //Se for sim, gravação das retiradas: procurar linhas contendo <<GRAVAR>> 
    else if (perguntaGravar == ui.Button.YES) { 
  
      var saida = ss.getSheetByName('SAÍDA');
      var gravar = saida.getRange('I3:I').createTextFinder('<<GRAVAR>>');
      var procurarGravar = gravar.findAll(); 

      // Processar as linhas encontradas em blocos para otimizar o desempenho
      const blockSize = 100; // Tamanho do bloco
      for (let i = 0; i < procurarGravar.length; i += blockSize) {
        const block = procurarGravar.slice(i, Math.min(i + blockSize, procurarGravar.length));

        // Processar o bloco de linhas
        block.forEach(cell => {
          // Listar intervalos I*:K* contendo <<GRAVAR>>
          var intervaloGravar = "I"+cell.getRow()+":K"+cell.getRow();
          var listaIntervaloGravar = saida.getRange(intervaloGravar);

          // Remover fórmulas dos intervalos <<GRAVAR>>
          listaIntervaloGravar.copyTo(listaIntervaloGravar,{contentsOnly:true});

          // Substituir <<GRAVAR>> por RETIRADO
          gravar.replaceAllWith('RETIRADO');
        });

        // Executar todas as alterações do loop antes de passar para o próximo bloco
        SpreadsheetApp.flush();

      }

    }

    //Mensagem final "SAÍDA concluída!"
    var mensagemFinal = ui.alert("SAÍDA concluída!",ui.ButtonSet.OK);
    if (mensagemFinal == ui.Button.OK) { return };
  
  }
    
}




