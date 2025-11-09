/**
 * Nome do Projeto: ScriptPlanilhaEstoqueControle
 *
 * Descrição: Este script cria uma planilha Google Sheets chamada EstoqueControle que organiza a 
 * entrada e saída de produtos, e gera tabelas de estoque.
 *            
 * Repositório: [https://github.com/FernandaPaulaDeCarvalho/PlanilhaEstoqueControle]
 * 
 * Copyright (c) 2025 Fernanda Paula de Carvalho
 *
 * Este código está licenciado sob a Licença MIT.
 * Consulte o arquivo LICENSE no repositório para obter mais informações.
 * 
 * IMPORTANTE:
 * 1. Este script está dividido em três etapas distintas que devem ser executadas em sequência  
 *    dentro do Google App Script.
 * 2. Este script requer permissões para criar e acessar planilhas no Google Sheets.
 *    Ao executar o script pela primeira vez, você será solicitado a conceder essas permissões.
 * 3. Este script depende do Google App Script e do GitHub e pode parar de funcionar se houver
 *    mudanças nas plataformas.
 *
 * ISENÇÃO DE RESPONSABILIDADE:
 * A autora deste script não se responsabiliza por quaisquer problemas ou danos que possam
 * ocorrer devido ao uso deste script.
*/

/** ATUALIZAR SEMPRE QUE HOUVER MODIFICAÇÃO */
// Definir versão atual, data de modificação
const versaoAtual = "1.0.0"; 
const dataModificacao = "2025-11-05";

// Definir links importantes
const repositGitHub ='https://github.com/FernandaPaulaDeCarvalho/PlanilhaEstoqueControle'; // README do github

const imageUrl = 'https://raw.githubusercontent.com/FernandaPaulaDeCarvalho/PlanilhaEstoqueControle/56232b96d3deb5aad49205876340fe63bb4d4d64/ImagemBotaoAmareloGravar.png'; // Imagem BotaoAmareloGravar 


// PRIMEIRA ETAPA: Criar as planilhas e formatar o layout
function primeiraEtapa() {
  
  // Definir o ano atual
  const anoAtual = new Date().getFullYear();
  const anoSeguinte = anoAtual + 1;

  // Definir versão e data atualização do script
  const versaoData = versaoAtual+" ("+dataModificacao+")";

  // Definir cores
  const amarelo = "#ffff00";
  const amareloEscuro3 = "#7f6000";
  const azulCentaureaClaro3 = "#c9daf8";
  const azulEscuro1 = "#3d85c6";
  const branco = "#ffffff";
  const ciano = "#00ffff";
  const cianoEscuro3 = "#0c343d";  
  const cinzaEscuro2 = "#999999";
  const cinzaEscuro3 = "#666666";
  const laranjaClaro3 = "#fce5cd";
  const laranjaEscuro2 = "#b45f06";
  const laranjaEscuro3 = "#783f04";
  const laranjaEscuro15 = "#ff6d01";
  const laranjaEscuro16 = "#522b02";  
  const preto = "#000000";
  const roxo = "#8989eb";
  const roxoClaro3 = "#e8e7fc";
  const roxoEscuro2 = "#351c75";
  const verde = "#00ff00";
  const verdeClaro1 = "#93c47d";
  const verdeClaro3 = "#d9ead3";
  const verdeEscuro2 = "#38761d";
  const verdeEscuro3 = "#274e13";
  const vermelho = "#ff0000";
  const vermelhoCerejaClaro1 = "#cc4125";
  const vermelhoCerejaEscuro2 = "#85200c";
  const vermelhoClaro3 = "#f4cccc";
  const vermelhoEscuro2 = "#990000";

  // Definir regras de formatação condicional:
  // SpreadsheetApp.newConditionalFormatRule() precisa ser chamado para cada regra 
  const regra1 = SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("PENDENTE").setFontColor(vermelhoClaro3).setBackground(vermelhoEscuro2).setBold(true); // status entradas/saídas
  const regra1A = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=B3:B="PENDENTE"').setFontColor(vermelhoClaro3).setBackground(vermelhoEscuro2).setBold(true); // coluna alerta entradas/saídas
  const regra3 = SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("INSERIDO").setFontColor(verdeClaro3).setBackground(verdeEscuro2).setBold(true); // status entradas
  const regra3A = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=B3:B="INSERIDO"').setFontColor(verdeClaro3).setBackground(verdeEscuro2).setBold(true); // coluna alerta entradas
  const regra4 = SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("RETIRADO").setFontColor(verdeClaro3).setBackground(verdeEscuro2).setBold(true); // status saídas
  const regra4A = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=B3:B="RETIRADO"').setFontColor(verdeClaro3).setBackground(verdeEscuro2).setBold(true); // coluna alerta saídas  
  const regra5 = SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("<<GRAVAR>>").setFontColor(laranjaEscuro3).setBackground(amarelo).setBold(true); // status entradas/saídas
  const regra5A = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=B3:B="<<GRAVAR>>"').setFontColor(laranjaEscuro3).setBackground(amarelo).setBold(true); // coluna alerta entradas/saídas 
  const regra2 = SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("erro").setFontColor(roxoClaro3).setBackground(roxo).setBold(true); // status entradas/saídas
  const regra6 = SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("FAZER REVISÃO").setFontColor(laranjaClaro3).setBold(true).setBackground(laranjaEscuro2); // tabela estoque
  const regra7 = SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("FAZER DOAÇÃO").setFontColor(azulCentaureaClaro3).setBackground(azulEscuro1).setBold(true); // tabela estoque
  const regra8 = SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("EM FALTA").setFontColor(vermelhoClaro3).setBackground(vermelhoEscuro2).setBold(true); // tabela estoque
  const regra9 = SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("CRÍTICO").setFontColor(amareloEscuro3).setBackground(amarelo).setBold(true); // tabela estoque
  const regra10 = SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("OK").setFontColor(verdeClaro3).setBackground(verdeEscuro2).setBold(true); // tabela estoque
  const regra11 = SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("revisar quant.atual").setFontColor(cinzaEscuro2).setItalic(true).setBold(true); // tabela lotes
  const regra12 = SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("Falta menos de 60 dias").setFontColor(laranjaEscuro15).setItalic(true).setBold(true); // tabela lotes
  const regra13 = SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("VENCIDO").setFontColor(vermelho).setUnderline(true).setBold(true); // tabela lotes



  // *** PLANILHA ESTOQUE CONTROLE ***
  const nomePlanilha = "EstoqueControle" + anoSeguinte; // Definir nome da planilha
  var ss = SpreadsheetApp.create(nomePlanilha); // Criar a planilha
  const id = ss.getId();// Obter o ID da planilha 
  ss.setSpreadsheetLocale("pt_BR"); // Definir o local para Português do Brasil

  // Adicionar as abas
  const nomesAbas = ["DETALHADO","ENTRADA","SAÍDA","Leia-me","invent","config"]; // Definir nomes das abas
  for (var i = 0; i < nomesAbas.length; i++) { // Loop para criar cada aba
    var nomeAba = nomesAbas[i];
    var sheet;

    // Se for a primeira aba, obtém a aba ativa e renomeia
    if (i === 0) {
      sheet = ss.getActiveSheet();
      sheet.setName(nomeAba); // Renomeia a primeira aba para "DETALHADO"
    }
    
    // Caso contrário, cria uma nova aba
    else {
      sheet = ss.insertSheet(nomeAba);
    }
  }



  // *** Formatação da aba "invent" ***
  var invent = ss.getSheetByName("invent"); // Acessar a aba
  const numColunasInvent = 4; // Definir o número de colunas
  invent.deleteColumns(numColunasInvent + 1, invent.getMaxColumns() - numColunasInvent); // Deletar colunas extras
  invent.insertRowsAfter(1, 4000); // Inserir 4000 linhas, após a 1ªlinha
  invent.setFrozenRows(2); // Congelar na segunda linha
  invent.getRange('A:D').setHorizontalAlignment("left"); // Ajustar texto a esquerda 
  invent.setTabColor(cinzaEscuro3); // Colorir a tag da aba

  const cabecaInvent = ["CÓD.INTERNO", "E-FISCO", "PRODUTOS", "QUANT.MÍN"]; // Definir cabeçalhos
  invent.getRange('A1:D2').setBackground(cinzaEscuro2); // Colorir cabeçalho
  invent.getRange(2, 1, 1, numColunasInvent).setValues([cabecaInvent]).setFontWeight("bold"); // Adicionar cabeçalhos em negrito na segunda linha

  // Aplicar os valores de largura nas colunas
  const larguraInvent = ["100", "200", "800", "100"]; // Definir valores de largura das colunas 
  for (var i = 0; i < numColunasInvent; i++) {
    invent.setColumnWidth(i + 1, larguraInvent[i]); // i + 1 porque as colunas começam em 1
  }



  // *** Formatação da aba "config" ***
  var config = ss.getSheetByName("config"); // Acessar a aba
  const numColunasConfig = 7; // Definir o número de colunas
  config.deleteColumns(numColunasConfig + 1, config.getMaxColumns() - numColunasConfig); //  Deletar colunas extras
  config.insertRowsAfter(1, 4000); // Inserir 4000 linhas, após a 1ªlinha
  config.setFrozenRows(2); // Congelar segunda linha
  config.setTabColor(cinzaEscuro3); // Colorir a tag da aba
  config.getRange('A:G').setHorizontalAlignment("left"); // Ajustar texto a esquerda
  config.getRangeList(["A1:B2","C:G"]).setBackground(cinzaEscuro2); // Colorir as colunas com as fórmulas
  
  const cabecaConfig = [ // Definir os nomes dos cabeçalhos
    ["", "", "produtos c/ quantidade em estoque maior que zero", "",	"",	"",	""],
    ["ANO ATUAL", "SETORES (MÁXIMO 10)", "produto", "fabricante",	"validade",	"lote",	"quant"]
  ]; 
  config.getRange(1, 1, 2, numColunasConfig).setValues(cabecaConfig).setFontWeight("bold"); // Inserir cabeçalhos
  
  // Aplicar os valores de largura nas colunas e mesclar
  const larguraConfig = ["100", "200", "100", "100", "100", "100", "100"]; // Definir valores de largura das colunas 
  for (var i = 0; i < numColunasConfig; i++) { // Loop
    config.setColumnWidth(i + 1, larguraConfig[i]); // i + 1 porque as colunas começam em 1
  }

  config.getRange('C1:G1').merge().setHorizontalAlignment("center"); // Meclar cabeçalho



  // *** Formatação da aba "ENTRADA" ***
  var entrada = ss.getSheetByName("ENTRADA"); // Acessar a aba
  const numColunasEntrada = 14; // Definir o número de colunas
  entrada.deleteColumns(numColunasEntrada + 1, entrada.getMaxColumns() - numColunasEntrada); // Deletar colunas extras
  entrada.insertRowsAfter(1, 4002); // Inserir 4002 linhas, após a 1ªlinha
  entrada.setFrozenRows(2); // Congelar segunda linha
  entrada.hideColumns(10, 5); // Ocultar intervalo cálculos J:N

  const cabecaEntrada = [ // Definir cabeçalhos na segunda linha
    ["", "", "", "", "", "", "", "", "" , "cálculos", "cálculos", "cálculos", "cálculos", "cálculos"],
    ["", "GRAVAR", "DATA ENTRADA", "PRODUTO", "FABRICANTE", "VALIDADE dd/mm/aaaa", "LOTE", "QUANT.", "OBSERVAÇÕES (OPCIONAL)", "ALERTA DIVERGÊNCIA", "STATUS#", "", "ABA#", "PRODUTO @FABRICANTE @VALIDADE @LOTE#"]
  ];
  entrada.getRange(1, 1, 2, numColunasEntrada).setValues(cabecaEntrada); // Adicionar cabeçalhos

  // Aplicar os valores de largura nas colunas, antes de mesclar
  const larguraEntrada = ["50", "92","150", "1000", "150", "120", "180", "100", "300", "100", "100", "130", "100", "350"]; // Definir valores de largura das colunas 
  for (var i = 0; i < numColunasEntrada; i++) { // Loop
    entrada.setColumnWidth(i + 1, larguraEntrada[i]); // i + 1 porque as colunas começam em 1
  }

  // Obter largura das colunas AeB para aplicar o botão amarelo
  var cellWidthEntrada = Number(larguraEntrada[0]) + Number(larguraEntrada[1]); // A(0) e B(1)

  // Aplicar cores e ajustes
  entrada.setTabColor(verde); // colorir a tag da aba
  entrada.getRange('J:N').setFontColor(branco).setBackground(preto); // colorir intervalo cálculos
  entrada.getRange('A3:C').setFontColor(verdeEscuro3).setBackground(verdeClaro1).setHorizontalAlignment("center"); // colorir colunas A:C
  entrada.getRange('A1:I1').setFontColor(verdeClaro1).setFontWeight("bold").setFontSize(15).setBackground(verdeEscuro3).setHorizontalAlignment("center"); // colorir cabeçalho primeira linha
  entrada.getRange('A2:I2').setFontColor(verdeEscuro3).setFontWeight("bold").setBackground(verdeClaro1).setFontSize(11).setHorizontalAlignment("center").setVerticalAlignment("middle").setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP); // colorir cabeçalho segunda linha e ajustar texto dentro da célula
  entrada.getRange('E3:I').setHorizontalAlignment("center"); // centralizar
  entrada.getRange('A2:B2').setFontColor(laranjaEscuro3).setBackground(amarelo).merge(); // Colorir botão amarelo

  // Definir regras de formatação condicional na coluna divergência
  var regraPendente = regra1A.setRanges([entrada.getRange("A3:A")]).build(); // Regra status "PENDENTE"
  var regraDivergente = regra2.setRanges([entrada.getRange("A3:A")]).build(); // Regra status "DIVERGENTE"
  var regraGravar = regra5A.setRanges([entrada.getRange("A3:A")]).build(); // Regra status "<<GRAVAR>>"
  var regraInserido = regra3A.setRanges([entrada.getRange("A3:A")]).build(); // Regra status "INSERIDO"
  
  // Aplicar as regras de formatação condicional na coluna divergência
  var regrasEntrada = entrada.getConditionalFormatRules();
  regrasEntrada.push(regraDivergente, regraPendente, regraGravar, regraInserido); 
  entrada.setConditionalFormatRules(regrasEntrada);

  // Definir regras de formatação condicional na coluna status
  var regraPendente = regra1.setRanges([entrada.getRange("B3:B")]).build(); // Regra status "PENDENTE"
  var regraGravar = regra5.setRanges([entrada.getRange("B3:B")]).build(); // Regra status "<<GRAVAR>>"
  var regraInserido = regra3.setRanges([entrada.getRange("B3:B")]).build(); // Regra status "INSERIDO"

  // Aplicar as regras de formatação condicional em status
  var regrasEntrada = entrada.getConditionalFormatRules();
  regrasEntrada.push(regraPendente, regraGravar, regraInserido); 
  entrada.setConditionalFormatRules(regrasEntrada);

  


  // *** Formatação da aba "SAÍDA" ***
  var saida = ss.getSheetByName("SAÍDA"); // Acessar a aba
  const numColunasSaida = 11; // Definir o número de colunas
  saida.deleteColumns(numColunasSaida + 1, saida.getMaxColumns() - numColunasSaida); // Deletar colunas extras
  saida.insertRowsAfter(1, 4002); // Inserir 4002 linhas, após a 1ªlinha
  saida.hideColumns(8, 4); // Ocultar colunas cálculos H:K
  saida.setFrozenRows(2); // Congelar segunda linha

  const cabecaSaida = [ // Definir cabeçalhos na segunda linha
    ["", "", "", "", "", "", "", "cálculos", "cálculos", "cálculos", "cálculos"],
    ["", "GRAVAR", "DATA SAÍDA", "SETOR", "LOTE @VALIDADE @PRODUTO @FABRICANTE", "QUANT. RETIRADA", "OBSERVAÇÕES (OPCIONAL)", "ALERTA DIVERGÊNCIA", "STATUS#", "", "ABA#"]
  ];
  saida.getRange(1, 1, 2, numColunasSaida).setValues(cabecaSaida); // Adicionar cabeçalhos na segunda linha
  
  // Aplicar os valores de largura nas colunas, antes de mesclar
  const larguraSaida = ["50", "92", "150", "120", "1000", "100", "300", "100", "100", "130", "100"]; 
  for (var i = 0; i < numColunasSaida; i++) {
    saida.setColumnWidth(i + 1, larguraSaida[i]); // i + 1 porque as colunas começam em 1
  }

  // Obter largura das colunas AeB para aplicar o botão amarelo
  var cellWidthSaida = Number(larguraSaida[0]) + Number(larguraSaida[1]); // A(0) e B(1)

  // Aplicar cores e ajustes
  saida.setTabColor(vermelhoCerejaClaro1); // colorir a tag da aba
  saida.getRange("H:K").setFontColor(branco).setBackground(preto); // colorir colunas colunas cálculos
  saida.getRange('A3:C').setFontColor(vermelhoCerejaEscuro2).setBackground(vermelhoClaro3).setHorizontalAlignment("center"); // colorir colunas A:C
  saida.getRange('A1:G1').setFontColor(vermelhoCerejaEscuro2).setFontWeight("bold").setFontSize(15).setBackground(vermelhoClaro3).setHorizontalAlignment("center"); // colorir cabeçalho primeira linha
  saida.getRange('A2:G2').setFontColor(vermelhoClaro3).setFontWeight("bold").setBackground(vermelhoCerejaEscuro2).setFontSize(11).setHorizontalAlignment("center").setVerticalAlignment("middle").setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP); // colorir e ajustar cabeçalho segunda linha
  saida.getRangeList(['D3:D', 'F3:G']).setHorizontalAlignment("center"); // centralizar
  saida.getRange('A2:B2').setFontColor(laranjaEscuro3).setBackground(amarelo).merge(); // Colorir botão amarelo

  // Definir regras de formatação condicional na coluna divergência
  var regraPendente = regra1A.setRanges([saida.getRange("A3:A")]).build(); // Regra status "PENDENTE"
  var regraDivergente = regra2.setRanges([saida.getRange("A3:A")]).build(); // Regra status "DIVERGENTE"
  var regraGravar = regra5A.setRanges([saida.getRange("A3:A")]).build(); // Regra status "<<GRAVAR>>"
  var regraRetirado = regra4A.setRanges([saida.getRange("A3:A")]).build(); // Regra status "RETIRADO"

  // Aplicar as regras de formatação condicional na coluna divergência
  var regrasSaida = saida.getConditionalFormatRules();
  regrasSaida.push(regraDivergente, regraPendente, regraGravar, regraRetirado); 
  saida.setConditionalFormatRules(regrasSaida);


  // Definir regras de formatação condicional na coluna status
  var regraPendente = regra1.setRanges([saida.getRange("B3:B")]).build(); // Regra status "PENDENTE"
  var regraGravar = regra5.setRanges([saida.getRange("B3:B")]).build(); // Regra status "<<GRAVAR>>"
  var regraRetirado = regra4.setRanges([saida.getRange("B3:B")]).build(); // Regra status "RETIRADO"

  // Aplicar as regras de formatação condicional na coluna status
  var regrasSaida = saida.getConditionalFormatRules();
  regrasSaida.push(regraPendente, regraGravar, regraRetirado); 
  saida.setConditionalFormatRules(regrasSaida);

  
  

  // *** Formatação da aba "DETALHADO" ***
  var detalh = ss.getSheetByName("DETALHADO"); // Acessar a aba
  const numColunasDetalh = 62; // Definir o número de colunas
  detalh.insertRowsAfter(1, 4002); // Inserir 4002 linhas, após a 1ªlinha
  detalh.setFrozenRows(2); // Congelar segunda linha
  detalh.getRange('E3:E').setFontWeight("bold").setFontSize(11); // aumentar tamanho da fonte e negrito

  // Definir os nomes dos cabeçalhos
  const cabecaDetalh = ["CÓD.INTERNO", "E-FISCO", "", "QUANT. MÍN", "QUANT. ATUAL", "STATUS", "", "CÓD.INTERNO", "E-FISCO", "", "", "","","","","","","","","","","","CÓD.INTERNO", "E-FISCO", "", "JAN", "FEV", "MAR", "ABR", "MAI", "JUN", "JUL", "AGO", "SET", "OUT", "NOV", "DEZ", "", "", "CÓD.INTERNO", "E-FISCO", "", "", "JAN", "FEV", "MAR", "ABR", "MAI", "JUN", "JUL", "AGO", "SET", "OUT", "NOV", "DEZ", "", "", "", "QUANT. ATUAL", "ALERTA", "", ""]; 
  detalh.getRange(2, 1, 1, numColunasDetalh).setValues([cabecaDetalh]); // Adicionar cabeçalhos (colunas extras serão adicionadas automaticamente)

  // Aplicar os valores de largura nas colunas
  const larguraDetalhado = ["120", "120", "800", "120", "120", "120", "30", "120", "120","800", "120", "120","120", "120","120", "120","120", "120","120","120", "70", "30","120","120", "800", "40", "40", "40", "40", "40", "40","40", "40", "40", "40", "40", "40", "60", "30","120", "120", "800", "60", "40", "40", "40", "40","40", "40", "40", "40", "40", "40", "40", "40","70", "30", "1000", "120", "135", "30", "1000"];
  for (var i = 0; i < numColunasDetalh; i++) {
    detalh.setColumnWidth(i + 1, larguraDetalhado[i]); // i + 1 porque as colunas começam em 1
  }

  // Definir regras de formatação condicional
  const regrasEstoque = [ // tabela ESTOQUE
    regra6.setRanges([detalh.getRange("F3:F")]).build(), // Regra para estoque "FAZER REVISÃO"
    regra7.setRanges([detalh.getRange("F3:F")]).build(), // Regra para estoque "FAZER DOAÇÃO"
    regra8.setRanges([detalh.getRange("F3:F")]).build(), // Regra para estoque "EM FALTA"
    regra9.setRanges([detalh.getRange("F3:F")]).build(), // Regra para estoque "CRÍTICO"
    regra10.setRanges([detalh.getRange("F3:F")]).build() // Regra para estoque "OK"
  ];

  const regrasLote = [ // tabela LOTE
    regra11.setRanges([detalh.getRange("BH3:BH")]).build(), // Regra para alerta "revisar quant.atual"
    regra12.setRanges([detalh.getRange("BH3:BH")]).build(), // Regra para alerta "Falta menos de 60 dias"
    regra13.setRanges([detalh.getRange("BH3:BH")]).build() // Regra para alerta "VENCIDO"  
  ];

  // Aplicar regras de formatação condicional
  var todasAsRegras = regrasEstoque.concat(regrasLote); // Juntar todas regras
  detalh.setConditionalFormatRules([]); // Remove todas as regras existentes na aba
  detalh.setConditionalFormatRules(todasAsRegras); // Adicionar regras

  // Aplicar cores e ajustes gerais
  detalh.setTabColor(ciano); // Colorir a tag da aba
  detalh.getRangeList(['A2:B', 'D2:I', 'K2:X', 'Z2:AO', 'AQ2:BE', 'BG2:BI']).setHorizontalAlignment("center"); // centralizar (as demais colunas ficarão "justificado a esquerda")
  detalh.getRangeList(['A1:F1', 'H1:J1', 'W1:Y1','AN1:AP1', 'BF1:BF1', 'BJ1:BJ1']).setFontColor(cinzaEscuro3).setFontStyle("italic"); // cabeçalho primeira linha em italico cinza, antes de mesclar
  detalh.getRangeList(['A2:F2', 'H2:U2', 'W2:AL2','AN2:BD2', 'BF2:BH2', 'BJ2:BJ2']).setFontWeight("bold") // cabeçalho segunda linha em Negrito
  detalh.getRangeList(['G:G', 'V:V', 'AM:AM', 'BE:BE', 'BI:BI']).setBorder(true, false, true, false, false, true, branco, SpreadsheetApp.BorderStyle.SOLID); // colorir as bordas com branco
  detalh.getRange('A1:BJ1').setBorder(true, true, true, true, true, true, branco, SpreadsheetApp.BorderStyle.SOLID); // colorir as bordas com branco
  detalh.getRange('BJ2:BJ2').setBackground(roxo); // tabela LOTE @VALIDADE @PRODUTO @FABRICANTE   
  detalh.getRange('BJ3:BJ').setBackground(roxoClaro3); // tabela LOTE @VALIDADE @PRODUTO @FABRICANTE 

  // Definir cores da fonte no corpo da tabelas, usando array de objetos
  var tabelas = [
    { range: "A2:F", fontColor: cianoEscuro3 }, // ESTOQUE
    { range: "H2:U", fontColor: laranjaEscuro16 }, // CONSUMO SETORIAL
    { range: "W2:AL", fontColor: laranjaEscuro3 }, // CONSUMO MENSAL
    { range: "AN2:BD", fontColor: amareloEscuro3 }, // ABASTECIMENTO
    { range: "BF2:BH", fontColor: verdeEscuro3 }, // PRODUTO @FABRICANTE @VALIDADE @LOTE
    { range: "BJ2:BJ", fontColor: roxoEscuro2 }, // LOTE @VALIDADE @PRODUTO @FABRICANTE    
  ];

  // Aplicar cores e ajustes dos cabeçalhos na segunda linha, usando função
  tabelas.forEach(function(tabela) {
    detalh.getRange(tabela.range)
    .setFontColor(tabela.fontColor); // Colorir fonte
  });

  // Aplicar cores alternadas no preenchimento das tabelas 
  var maxLinhas = detalh.getMaxRows(); // Definir o maximo de linhas
  detalh.getRange(2, 1, maxLinhas, 6).applyRowBanding(SpreadsheetApp.BandingTheme.CYAN, true, false); // ESTOQUE
  detalh.getRange(2, 8, maxLinhas, 14).applyRowBanding(SpreadsheetApp.BandingTheme.BROWN, true, false); // CONSUMO SETORIAL
  detalh.getRange(2, 23, maxLinhas, 16).applyRowBanding(SpreadsheetApp.BandingTheme.ORANGE, true, false); // CONSUMO MENSAL
  detalh.getRange(2, 40, maxLinhas, 17).applyRowBanding(SpreadsheetApp.BandingTheme.YELLOW, true, false); // ABASTECIMENTO
  detalh.getRange(2, 58, maxLinhas, 3).applyRowBanding(SpreadsheetApp.BandingTheme.GREEN, true, false); // PRODUTO @FABRICANTE @VALIDADE @LOTE 




  // *** Formatação da aba "Leia-me" ***
  var leia = ss.getSheetByName("Leia-me"); // Acessar a aba
  const numColunasLeia = 2; // Definir o número de colunas
  leia.deleteColumns(numColunasLeia + 1, leia.getMaxColumns() - numColunasLeia); // Deletar colunas extras
  leia.setTabColor(roxo); // colorir tag da aba
  leia.getRange('A:B').setFontColor(roxoEscuro2).setFontSize(13).setBackground(roxoClaro3).setBorder(true, true, true, true, true, true, roxoClaro3, SpreadsheetApp.BorderStyle.SOLID); // colorir todo intervalo incluindo as bordas

  // Aplicar os valores de largura nas colunas
  const larguraLeia = ["900", "900"]; // Definir valores
  for (var i = 0; i < numColunasLeia; i++) {
    leia.setColumnWidth(i + 1, larguraLeia[i]); // i + 1 porque as colunas começam em 1
  }



  

  // *** FINALIZAR PRIMEIRA ETAPA ***
  // Forçar a atualização da planilha
  SpreadsheetApp.flush();

  // Acessar "propriedades do script" para salvar informações para a próxima etapa
  PropertiesService.getScriptProperties().setProperty("versaoData",versaoData);
  PropertiesService.getScriptProperties().setProperty("repositGitHub",repositGitHub);
  PropertiesService.getScriptProperties().setProperty("id",id);
  PropertiesService.getScriptProperties().setProperty("nomePlanilha",nomePlanilha);
  PropertiesService.getScriptProperties().setProperty("imageUrl",imageUrl);
  PropertiesService.getScriptProperties().setProperty("cellWidthEntrada",cellWidthEntrada);
  PropertiesService.getScriptProperties().setProperty("cellWidthSaida",cellWidthSaida);
  PropertiesService.getScriptProperties().setProperty("anoSeguinte",anoSeguinte);
  
  // Mensagem de conclusão
  Logger.log(
   "PRIMEIRA ETAPA finalizada com sucesso!\n" +
   "Planilha " + nomePlanilha + " ID: " + id + "\n" +
   "Pronta para prosseguir para a SEGUNDA ETAPA."
  );

}






// SEGUNDA ETAPA: Adicionar fórmulas e textos à planilha
function segundaEtapa() {

  // Acessar "propriedades do script" para recuperar as informações
  const versaoData = PropertiesService.getScriptProperties().getProperty("versaoData");
  const repositGitHub = PropertiesService.getScriptProperties().getProperty("repositGitHub");
  const id = PropertiesService.getScriptProperties().getProperty("id");
  const nomePlanilha = PropertiesService.getScriptProperties().getProperty("nomePlanilha");
  const imageUrl = PropertiesService.getScriptProperties().getProperty("imageUrl");
  const cellWidthEntradaString = PropertiesService.getScriptProperties().getProperty("cellWidthEntrada");
  const cellWidthSaidaString = PropertiesService.getScriptProperties().getProperty("cellWidthSaida");
  const anoSeguinteString = PropertiesService.getScriptProperties().getProperty("anoSeguinte");

  // Converter para inteiro (base 10) 
  const anoSeguinte = parseInt(anoSeguinteString, 10);   
  const cellWidthEntrada = parseInt(cellWidthEntradaString, 10);   
  const cellWidthSaida = parseInt(cellWidthSaidaString, 10);  
 
  // Definir fórmulas pontuais que ocupam apenas uma célula
  const formula1 = '=SPLIT(SORT(UNIQUE(ARRAYFORMULA(IF(DETALHADO!$BF3:$BF="";"";IF(DETALHADO!$BG3:$BG>0;INDEX(SPLIT(DETALHADO!$BF3:$BF;" @";FALSE;FALSE);0;4)&" @"&TEXT(INDEX(SPLIT(DETALHADO!$BF3:$BF;" @";FALSE;FALSE);0;3);"dd/mm/yyyy")&" @"&INDEX(SPLIT(DETALHADO!$BF3:$BF;" @";FALSE;FALSE);0;1)&" @"&INDEX(SPLIT(DETALHADO!$BF3:$BF;" @";FALSE;FALSE);0;2);""))));1;FALSE);" @";FALSE;FALSE)'; // Criar array na aba config contendo todos lotes que apresentam quantidade > zero, concatenando LOTE" @"VALIDADE" @"PRODUTO" @"FABRICANTE. Será usada para exportar os dados para a planilha estoque para o ano seguinte.
  const formula2 = '="ENTRADA ("&config!$A$3&")"'; // cabeçalho na primeira linha
  const formula3 = '="SAÍDA ("&config!$A$3&")"'; // cabeçalho na primeira linha
  const formula4 = '="ABASTECIMENTO MENSAL ("&config!$A$3&")"'; // cabeçalho da tabela
  const formula5 = '="CONSUMO SETORIAL ("&config!$A$3&")"'; // cabeçalho da tabela 
  const formula6 = '="CONSUMO MENSAL ("&config!$A$3&")"'; // cabeçalho da tabela
  const formula7 = '="ESTOQUE DE PRODUTOS ("&config!$A$3&")"'; // cabeçalho da tabela
  const formula8 = '="PRODUTO @FABRICANTE @VALIDADE @LOTE ("&config!$A$3&")"'; // cabeçalho da tabela
  const formula9 = '="LOTE @VALIDADE @PRODUTO @FABRICANTE ("&config!$A$3&")"'; // cabeçalho da tabela
  const formula10 = '=ARRAYFORMULA($J$3:$L)'; // Visualizar dados de entrada das colunas cálculos
  const formula11 = '=ARRAYFORMULA($H$3:$J)'; // Visualizar dados de saida das colunas cálculos
  const formula12 = '=TEXT(NOW();"dd/mm/yyyy HH:MM")'; // Obter data hora
  const formula13 = '="Fonte: planilha '+nomePlanilha+'" & " (atualizado em: " &TEXT(NOW();"dd/mm/yyyy HH:MM")& ")"'; // Cabeçalho na primeira linha com data e hora de edição
  const formula14 = '=$A$1'; // Copiar cabeçalho na primeira linha com data e hora de edição
  const formula15 = '=ARRAYFORMULA(invent!$A$3:$C)'; // Adicionar a lista com todos os produtos de invent
  const formula16 = '=ARRAYFORMULA(invent!$A$3:$D)'; // Adicionar a lista com todos os produtos de invent+quant min
  const formula17 = '=config!$A$3-1'; // Adicionar cabeçalho ano anterior na segunda linha
  const formula18 = '=config!$A$3'; // Adicionar  cabeçalho ano atual na segunda linha
  const formula19 = '=TRANSPOSE(ARRAYFORMULA(config!$B$3:$B$12))'; // Adicionar os setores na aba DETALHADO
  const formula20 = '=MAP($AP3:$AP;LAMBDA(celula_ap;IF(celula_ap="";"";SUMIFS(ENTRADA!$H$3:$H;ENTRADA!D$3:$D;celula_ap;ARRAYFORMULA(INT(ENTRADA!$L$3:$L));"<"&DATE(config!$A$3;1;1)))))'; // Sumifs para calcular quantitativos que entraram no ano anterior
  const formula21 = '=LET(ano;config!$A$3;intervalo_ap;$AP$3:$AP;meses;SEQUENCE(1;6;1);formula_mes;LAMBDA(mes_num;LET(data_inicio;DATE(ano;mes_num;1);proximo_mes;EOMONTH(data_inicio;0)+1;MAP(intervalo_ap;LAMBDA(celula_ap;IF(celula_ap="";"";SUMIFS(ENTRADA!$H$3:$H;ENTRADA!D$3:$D;celula_ap;ARRAYFORMULA(INT(ENTRADA!$L$3:$L));">="&data_inicio;ARRAYFORMULA(INT(ENTRADA!$L$3:$L));"<"&proximo_mes))))));MAP(meses;formula_mes))'; // Sumifs para calcular quantitativos que entraram no 1ºsemestre do ano atual
  const formula22 = '=LET(ano;config!$A$3;intervalo_ap;$AP$3:$AP;meses;SEQUENCE(1;6;7);formula_mes;LAMBDA(mes_num;LET(data_inicio;DATE(ano;mes_num;1);proximo_mes;EOMONTH(data_inicio;0)+1;MAP(intervalo_ap;LAMBDA(celula_ap;IF(celula_ap="";"";SUMIFS(ENTRADA!$H$3:$H;ENTRADA!D$3:$D;celula_ap;ARRAYFORMULA(INT(ENTRADA!$L$3:$L));">="&data_inicio;ARRAYFORMULA(INT(ENTRADA!$L$3:$L));"<"&proximo_mes))))));MAP(meses;formula_mes))'; // Sumifs para calcular quantitativos que entraram no 2ºsemestre do ano atual
  const formula23 = '=ARRAYFORMULA(IF($AP$3:$AP="";"";$AQ$3:$AQ+$AR$3:$AR+$AS$3:$AS+$AT$3:$AT+$AU$3:$AU+$AV$3:$AV+$AW$3:$AW+$AX$3:$AX+$AY$3:$AY+$AZ$3:$AZ+$BA$3:$BA+$BB$3:$BB+$BC$3:$BC))'; // Calcular somatório que entrou no ano atual (somando com os produtos que sobraram do ano anterior)
  const formula24 = '=LET(intervalo_j;$J$3:$J;linhas;SEQUENCE(ROWS(intervalo_j));setores;{$K$2\\$L$2\\$M$2\\$N$2\\$O$2\\$P$2\\$Q$2\\$R$2\\$S$2\\$T$2};resultado;MAP(setores;LAMBDA(setor;MAP(linhas;LAMBDA(linha;LET(celula_j;INDEX(intervalo_j;linha);IF(OR(celula_j="";ARRAYFORMULA(SAÍDA!$E$3:$E)="");"";SUMIFS(SAÍDA!$F$3:$F;ARRAYFORMULA(INDEX(SPLIT(SAÍDA!$E$3:$E;" @";FALSE;FALSE);0;3));celula_j;SAÍDA!$I$3:$I;"RETIRADO";SAÍDA!$D$3:$D;setor)))))));resultado)'; // Sumifs para calcular quantitativos retirados por setor
  const formula25 = '=ARRAYFORMULA(IF($J$3:$J="";"";$K$3:K+$L$3:$L+$M$3:$M+$N$3:$N+$O$3:$O+$P$3:$P+$Q$3:$Q+$R$3:$R+$S$3:$S+$T$3:$T))'; // Calcular somatório que foi retirado no ano atual
  const formula26 = '=LET(ano;config!$A$3;intervalo_y;$Y$3:$Y;linhas;SEQUENCE(ROWS(intervalo_y));meses;SEQUENCE(1;6;1);formula_mes;LAMBDA(mes;LET(data_inicio;DATE(ano;mes;1);proximo_mes;EOMONTH(data_inicio;0)+1;MAP(linhas;LAMBDA(linha;LET(celula_y;INDEX(intervalo_y;linha);IF(OR(celula_y="";ARRAYFORMULA(SAÍDA!$E$3:$E)="");"";SUMIFS(SAÍDA!$F$3:$F;ARRAYFORMULA(INDEX(SPLIT(SAÍDA!$E$3:$E;" @";FALSE;FALSE);0;3));celula_y;SAÍDA!$I$3:$I;"RETIRADO";ARRAYFORMULA(INT(SAÍDA!$J$3:$J));">="&data_inicio;ARRAYFORMULA(INT(SAÍDA!$J$3:$J));"<"&proximo_mes)))))));MAP(meses;formula_mes))'; // Sumifs para calcular quantitativos retirados no 1ºsemestre do ano atual
  const formula27 = '=LET(ano;config!$A$3;intervalo_y;$Y$3:$Y;linhas;SEQUENCE(ROWS(intervalo_y));meses;SEQUENCE(1;6;7);formula_mes;LAMBDA(mes;LET(data_inicio;DATE(ano;mes;1);proximo_mes;EOMONTH(data_inicio;0)+1;MAP(linhas;LAMBDA(linha;LET(celula_y;INDEX(intervalo_y;linha);IF(OR(celula_y="";ARRAYFORMULA(SAÍDA!$E$3:$E)="");"";SUMIFS(SAÍDA!$F$3:$F;ARRAYFORMULA(INDEX(SPLIT(SAÍDA!$E$3:$E;" @";FALSE;FALSE);0;3));celula_y;SAÍDA!$I$3:$I;"RETIRADO";ARRAYFORMULA(INT(SAÍDA!$J$3:$J));">="&data_inicio;ARRAYFORMULA(INT(SAÍDA!$J$3:$J));"<"&proximo_mes)))))));MAP(meses;formula_mes))'; // Sumifs para calcular quantitativos retirados no 2ºsemestre do ano atual
  const formula28 = '=ARRAYFORMULA(IF($Y$3:$Y="";"";$Z$3:$Z+$AA$3:$AA+$AB$3:$AB+$AC$3:$AC+$AD$3:$AD+$AE$3:$AE+$AF$3:$AF+$AG$3:$AG+$AH$3:$AH+$AI$3:$AI+$AJ$3:$AJ+$AK$3:$AK))'; // Calcular somatório que foi retirado no ano atual
  const formula29 = '=ARRAYFORMULA(IF($C$3:$C="";"";$BD$3:$BD-$AL$3:$AL))'; // Calcular Quant.Atual a partir das tabelas CONSUMO MENSAL e ABASTECIMENTO
  const formula30 = '=ARRAYFORMULA(IF($C$3:$C="";"";IFS($E$3:$E<0;"FAZER REVISÃO";($D$3:$D=0)*($E$3:$E=0);"OK";($D$3:$D=0)*($E$3:$E>0);"FAZER DOAÇÃO";($D$3:$D>0)*($E$3:$E=0);"EM FALTA";$E$3:$E<$D3:$D;"CRÍTICO";$E$3:$E>=$D$3:$D;"OK")))'; // Classificar o Status do estoque dos produtos de acordo com a Quant.Mín
  const formula31 = '=SORT(UNIQUE(QUERY(ENTRADA!N$3:$N;"select Col1 where Col1 is not null"));1;TRUE)'; // Adicionar na aba DETALHADO a lista com todos os lotes que entraram, ordenando de acordo com o produto 
  const formula32 = '=MAP($BF$3:$BF;LAMBDA(celula_bf;IF(celula_bf="";"";SUMIFS(ENTRADA!$H$3:$H;ENTRADA!N$3:$N;celula_bf)-SUMIFS(SAÍDA!$F$3:$F;ARRAYFORMULA(INDEX(SPLIT(SAÍDA!$E$3:$E;" @";FALSE;FALSE);0;3)&" @"&INDEX(SPLIT(SAÍDA!$E$3:$E;" @";FALSE;FALSE);0;4)&" @"&TEXT(INDEX(SPLIT(SAÍDA!$E$3:$E;" @";FALSE;FALSE);0;2);"dd/mm/yyyy")&" @"&INDEX(SPLIT(SAÍDA!$E$3:$E;" @";FALSE;FALSE);0;1));celula_bf;SAÍDA!$I$3:$I;"RETIRADO"))))'; // Calcular Quant.Atual dos lotes
  const formula33 = '=ARRAYFORMULA(IF($BF3:$BF="";"";IFS(($BG3:$BG=0)+((TEXT(IFERROR(INDEX(SPLIT($BF3:$BF;" @";FALSE;FALSE);0;3);"");"dd/mm/yyyy")= ".")*($BG3:$BG>=0));"-";$BG3:$BG<0;"revisar quant.atual";($BG3:$BG>0)*(TEXT(IFERROR(INDEX(SPLIT($BF3:$BF;" @";FALSE;FALSE);0;3);"");"dd/mm/yyyy")-TODAY()<= 0);"VENCIDO";($BG3:$BG>= 0)*(TEXT(IFERROR(INDEX(SPLIT($BF3:$BF;" @";FALSE;FALSE);0;3);"");"dd/mm/yyyy")-TODAY()>60);"-";($BG3:$BG>=0)*(TEXT(IFERROR(INDEX(SPLIT($BF3:$BF;" @";FALSE;FALSE);0;3);""); "dd/mm/yyyy")-TODAY()<= 60)*(TEXT(IFERROR(INDEX(SPLIT($BF3:$BF;" @";FALSE;FALSE);0;3);"");"dd/mm/yyyy")-TODAY()>0);"Falta menos de 60 dias")))'; // Classificar Status do lote quanto a validade e quantitativo
  const formula34 = '=SORT(UNIQUE(ARRAYFORMULA(IF(DETALHADO!$BF3:$BF="";"";INDEX(SPLIT(DETALHADO!$BF3:$BF;" @";FALSE;FALSE);0;4)&" @"&TEXT(INDEX(SPLIT(DETALHADO!$BF3:$BF;" @";FALSE;FALSE);0;3);"dd/mm/yyyy")&" @"&INDEX(SPLIT(DETALHADO!$BF3:$BF;" @";FALSE;FALSE);0;1)&" @"&INDEX(SPLIT(DETALHADO!$BF3:$BF;" @";FALSE;FALSE);0;2))));1;FALSE)'; // Adicionar a lista com todos os lotes que entraram, ordenando de acordo com o lote
  const formula35 = '=UNIQUE(invent!$C$3:$C)'; // Gera lista com os produtos de invent alimentando o inventario da aba "Leia-me"
  const formula36 = '=ARRAYFORMULA(IF($D$3:$D="";"";IF(COUNTIFS(invent!$C$3:$C;$D$3:$D)>0;"";"erro")))'; // Calcula de a alguma entrada é divergente em relação ao invent após a gravação
  const formula37 = '=ARRAYFORMULA(IF($E$3:$E="";"";IF(COUNTIFS(DETALHADO!$BJ3:$BJ;$E$3:$E)>0;"";"erro")))'; // Calcula se alguma saida é divergente em relação a entrada após a gravação



  // Definir fórmulas para serem arrastadas completando toda a coluna
  // INDIRECT(ADDRESS(ROW()-1;COLUMN();4))<>"" serve para evitar gravar deixando linha em branco
  const formula51 = '=LET(gravar;IF(AND($D{linha}<>"";$E{linha}<>"";$G{linha}<>"";ISNUMBER($H{linha})=TRUE;$H{linha}>0;INDIRECT(ADDRESS(ROW()-1;COLUMN();4))<>"";OR(ISDATE($F{linha})=TRUE;$F{linha}="."));"<<GRAVAR>>";"PENDENTE");agora;TO_DATE($L$2);IF(AND($D{linha}="";$E{linha}="");"";{gravar\\agora\\{setor}}))'; // Calcular aba ENTRADA coluna K:M. Classificar o status de acordo com o preenchimento das entradas e preencher data atual e entrada
  const formula52 = '=LET(gravar;IF(AND($D{linha}<>"";$E{linha}<>"";ISNUMBER($F{linha})=TRUE;INDIRECT(ADDRESS(ROW()-1;COLUMN();4))<>"");"<<GRAVAR>>";"PENDENTE");agora;TO_DATE($J$2);IF(AND($D{linha}="";$E{linha}="";$F{linha}="";$G{linha}="");"";{gravar\\agora\\{setor}}))'; // Calcular aba SAÍDA coluna H. Classificar o status de acordo com o preenchimento das saidas e preencher data atual e saida
  const formula53 = '=IF($K{linha}<>"";$D{linha}&" @"&REGEXREPLACE(REGEXREPLACE(REGEXREPLACE(REGEXREPLACE(REGEXREPLACE(REGEXREPLACE(REGEXREPLACE(REGEXREPLACE(REGEXREPLACE(REGEXREPLACE(REGEXREPLACE(REGEXREPLACE(UPPER($E{linha});"[ùúüû]";"u");"[èéêë]";"e");"[àáãâäå]";"a");"[íìîï]";"i");"[óòöôõ]";"o");"ç";"c");"[ÙÚÜÛ]";"U");"[ÈÉÊË]";"E");"[ÀÁÃÂÄÅ]";"A");"[ÍÌÎÏ]";"I");"[ÓÒÖÔÕ]";"O");"Ç";"C")&" @" &TEXT($F{linha};"dd/mm/yyyy")&" @"&UPPER($G{linha});"")'; // Calcular coluna N aba ENTRADA. Concatenar dados das entradas em produto@fabricante@validade@lote, removendo acentuação e transformando maiúsculo
  


  // Acessar a planilha
  const ss = SpreadsheetApp.openById(id); 


  // *** Formatação da aba "config" ***
  var config = ss.getSheetByName("config"); // Acessar a aba
  var dataConfig = [[anoSeguinte, "", ""], ["", "", ""]]; // Arrays com dados
  config.getRange(3, 1, dataConfig.length, dataConfig[0].length).setValues(dataConfig); // Adicionar dados

  // Adicionar fórmulas pontuais
  config.getRange(3, 3).setFormula(formula1); // calcular "produtos c/ quantidade em estoque maior que zero"



  // *** Formatação da aba "ENTRADA" ***
  var entrada = ss.getSheetByName("ENTRADA"); // Acessar a aba
  const regra1 = SpreadsheetApp.newDataValidation().requireValueInRange(ss.getSheetByName("invent").getRange("$C$3:$C"),true).setAllowInvalid(false).build(); // Criar menu suspenso de produtos para a aba ENTRADA usando o intervalo invent!C3:C  
  entrada.getRange("D3:D").setDataValidation(regra1); // Adicionar validação de dados (menu suspenso) 

  // Adicionar fórmulas pontuais
  entrada.getRange('A1:I1').merge().setFormula(formula2); // Mesclar cabeçalho na primeira linha
  entrada.getRange('A3').setFormula(formula10); // Visualização de dados ocultos
  entrada.getRange('J3').setFormula(formula36); // Inserir alerta de divergência "erro"
  entrada.getRange('L2').setFormula(formula12); // DATA HORA ENTRADA

  // Adicionar fórmulas para serem arrastadas
  var numLinhas = entrada.getMaxRows() - 2; // Definir número de linhas para preencher começando na linha 3

  // Inserir nome do setor na formula
  var setor = entrada.getName(); // Obter nome da aba ativa
  var formulaOriginal = formula51.replace(/{setor}/,"\"" + setor + "\""); // Fórmula original, substituindo o marcador {setor} pelo string do setor, mas sem a referência à linha, que calcula colunas K:M

  // Loop para preencher o array com as fórmulas
  var arrayFormulas = []; // Criar um array bidimensional para armazenar as fórmulas
  for (var i = 0; i < numLinhas; i++) {
    var linha = i + 3; // Definir o número da linha atual começando na linha 3 até o final
    var formulaAtual = formulaOriginal.replace(/{linha}/g, linha); // Substituir o marcador {linha} pelo número da linha atual na fórmula
    arrayFormulas.push([formulaAtual]); // Adicionar as fórmulas ao array
  }

  // Definir o intervalo onde o array (com as fórmulas) será inserido
  var intervalo = entrada.getRange(3, 11, numLinhas, 1); // Coluna K (10), a partir da linha 3
  intervalo.setFormulas(arrayFormulas); // Adicionar o array ao intervalo

  // Adicionar formula53 com loop
  var arrayFormulas = []; // Criar um array bidimensional para armazenar as fórmulas
  for (var i = 0; i < numLinhas; i++) {
    var linha = i + 3; // Definir o número da linha atual começando na linha 3 até o final
    var formulaAtual = formula53.replace(/{linha}/g, linha); // Substituir o marcador {linha} pelo número da linha atual na fórmula
    arrayFormulas.push([formulaAtual]); // Adicionar as fórmulas ao array
  }  

  // Definir o intervalo onde o array (com as fórmulas) será inserido
  var intervalo = entrada.getRange(3, 14, numLinhas, 1); // Coluna N (14), a partir da linha 3
  intervalo.setFormulas(arrayFormulas); // Adicionar o array ao intervalo



 


  // *** Formatação da aba "SAÍDA" ***
  var saida = ss.getSheetByName("SAÍDA"); // Acessar a aba
  const regra2 = SpreadsheetApp.newDataValidation().requireValueInRange(ss.getSheetByName("DETALHADO").getRange("$BJ$3:$BJ"),true).setAllowInvalid(false).build(); // Criar menu suspenso de produtos usando o intervalo DETALHADO!BJ3:BJ
  const regra3 = SpreadsheetApp.newDataValidation().requireValueInRange(ss.getSheetByName("config").getRange("$B$3:$B$12"),true).setAllowInvalid(false).build(); // Criar menu suspenso de setores usando o intervalo config!B3:B12
  saida.getRange("D3:D").setDataValidation(regra3); // Adicionar menu suspenso com setores
  saida.getRange("E3:E").setDataValidation(regra2); // Adicionar menu suspenso lista produtos para retirada
  
  // Adicionar fórmulas pontuais
  saida.getRange('A1:G1').merge().setFormula(formula3); // Mesclar cabeçalho na primeira linha
  saida.getRange('A3').setFormula(formula11); // Visualização de dados ocultos
  saida.getRange('H3').setFormula(formula37); // Inserir alerta de divergência "erro"
  saida.getRange('J2').setFormula(formula12); // DATA HORA SAIDA

  // Adicionar fórmula para ser arrastada
  var numLinhas = saida.getMaxRows() - 2; // Definir número de linhas para preencher começando na linha 3

  // Inserir nome do setor na formula
  var setor = saida.getName(); // Obter nome da aba ativa
  var formulaOriginal = formula52.replace(/{setor}/,"\"" + setor + "\""); // Fórmula original, substituindo o marcador {setor} pelo string do setor, mas sem a referência à linha, que calcula colunas I:K

  // Loop para preencher o array com as fórmulas
  var arrayFormulas = []; // Criar um array bidimensional para armazenar as fórmulas
  for (var i = 0; i < numLinhas; i++) {
    var linha = i + 3; // Definir o número da linha atual começando na linha 3 até o final
    var formulaAtual = formulaOriginal.replace(/{linha}/g, linha); // Substituir o marcador {linha} pelo número da linha atual na fórmula
    arrayFormulas.push([formulaAtual]); // Adicionar as fórmulas ao array
  }

  // Definir o intervalo onde o array (com as fórmulas) será inserido
  var intervalo = saida.getRange(3, 9, numLinhas, 1); // Coluna I (9), a partir da linha 3
  intervalo.setFormulas(arrayFormulas); // Adicionar o array ao intervalo




  // *** Formatação da aba "DETALHADO" ***
  var detalh = ss.getSheetByName("DETALHADO"); // Acessar a aba

  // Adicionar fórmulas pontuais na tabela ABASTECIMENTO
  detalh.getRange('AN1:AP1').merge().setFormula(formula14); // Adicionar cabeçalho nome planilha, data e hora na primeira linha
  detalh.getRange('AP2').setFormula(formula4); // Adicionar cabeçalho nome da tabela na segunda linha
  detalh.getRange('AQ2').setFormula(formula17); // Adicionar cabeçalho ano anterior na segunda linha
  detalh.getRange('BD2').setFormula(formula18); // Adicionar cabeçalho ano atual na segunda linha
  detalh.getRange('AN3').setFormula(formula15); // Adicionar a lista com todos os produtos de invent
  detalh.getRange('AQ3').setFormula(formula20); // Calcular quantitativos que entraram no ano anterior
  detalh.getRange('AR3').setFormula(formula21); // Calcular quantitativos que entraram no 1ºsemestre do ano atual
  detalh.getRange('AX3').setFormula(formula22); // Calcular quantitativos que entraram no 2ºsemestre do ano atual
  detalh.getRange('BD3').setFormula(formula23); // Calcular somatório que entrou no ano atual (somando com os produtos que sobraram do ano anterior)

  // Adicionar fórmulas pontuais na tabela CONSUMO SETORIAL
  detalh.getRange('H1:J1').merge().setFormula(formula14); // Adicionar cabeçalho nome planilha, data e hora na primeira linha
  detalh.getRange('J2').setFormula(formula5); // Adicionar cabeçalho nome da tabela na segunda linha
  detalh.getRange('U2').setFormula(formula18); // Adicionar cabeçalho ano atual na segunda linha
  detalh.getRange('K2').setFormula(formula19); // Adicionar setores a partir da lista em config
  detalh.getRange('H3').setFormula(formula15); // Adicionar a lista com todos os produtos de invent
  detalh.getRange('K3').setFormula(formula24); // Calcular quantitativos retirados por setor  
  detalh.getRange('U3').setFormula(formula25); // Calcular somatório que foi retirado no ano atual 

  // Adicionar fórmulas pontuais na tabela CONSUMO MENSAL
  detalh.getRange('W1:Y1').merge().setFormula(formula14); // Adicionar cabeçalho nome planilha, data e hora na primeira linha
  detalh.getRange('Y2').setFormula(formula6); // Adicionar cabeçalho nome da tabela na segunda linha
  detalh.getRange('AL2').setFormula(formula18); // Adicionar cabeçalho ano atual na segunda linha
  detalh.getRange('W3').setFormula(formula15); // Adicionar a lista com todos os produtos de invent
  detalh.getRange('Z3').setFormula(formula26); // Calcular quantitativos retirados no 1ºsemestre do ano atual  
  detalh.getRange('AF3').setFormula(formula27); // Calcular quantitativos retirados no 2ºsemestre do ano atual 
  detalh.getRange('AL3').setFormula(formula28); // Calcular somatório que foi retirado no ano atual

  // Adicionar fórmulas pontuais na tabela ESTOQUE
  detalh.getRange('A1:C1').merge().setFormula(formula13); // Adicionar cabeçalho nome planilha, data e hora na primeira linha
  detalh.getRange('C2').setFormula(formula7); // Adicionar cabeçalho nome da tabela na segunda linha
  detalh.getRange('A3').setFormula(formula16); // Adicionar a lista com todos os produtos de invent
  detalh.getRange('E3').setFormula(formula29); // Calcular Quant.Atual a partir das tabelas CONSUMO e ABASTECIMENTO
  detalh.getRange('F3').setFormula(formula30); // Classificar o Status dos produtos de acordo com a Quant.Mín

  // Adicionar fórmulas pontuais na tabela PRODUTOS
  detalh.getRange('BF1').setFormula(formula14); // Adicionar cabeçalho nome planilha, data e hora na primeira linha
  detalh.getRange('BF2').setFormula(formula8); // Adicionar cabeçalho nome da tabela na segunda linha
  detalh.getRange('BF3').setFormula(formula31); // Adicionar a lista com todos os lotes que entraram, ordenando de acordo com o produto
  detalh.getRange('BG3').setFormula(formula32); // Calcular Quant.Atual dos lotes
  detalh.getRange('BH3').setFormula(formula33); // Classificar Status do lote quanto a validade e quantitativo

  // Adicionar fórmulas pontuais na tabela LOTES
  detalh.getRange('BJ1').setFormula(formula14); // Adicionar cabeçalho nome planilha, data e hora na primeira linha
  detalh.getRange('BJ2').setFormula(formula9); // Adicionar cabeçalho nome da tabela na segunda linha
  detalh.getRange('BJ3').setFormula(formula34); // Adicionar a lista com todos os lotes que entraram, ordenando de acordo com o lote




  // *** Formatação da aba "Leia-me" ***

  // Definir as instruções
  var instrucoes = [
    ["INSTRUÇÕES PARA REGISTRO DE ENTRADAS E SAÍDAS DE PRODUTOS DO ESTOQUE VERSÃO: " + versaoData, 'bold', 'underline', null],
    ["", null, null, null], // Linha em branco
    ["IMPORTANTE:", 'bold', 'underline', null],
    ["*A PLANILHA ESTOQUE CONTROLE CONTABILIZA PRODUTOS PARA 1 (UM) ANO", 'bold', null, null],
    ["*COM ISSO, É NESSESSÁRIO GERAR UMA NOVA PLANILHA AO FINAL DE CADA ANO", 'bold', null, null],
    ["*A PLANILHA ESTOQUE CONTROLE TEM CAPACIDADE PARA 5MIL ENTRADAS E 5MIL SAÍDAS", 'bold', null, null],
    ["*AS CÉLULAS COLORIDAS CONTÉM FÓRMULAS. DIGITE APENAS NAS CÉLULAS BRANCAS", 'bold', null, null],
    ["*USE Ctrl+C (COPIAR) E Ctrl+Shift+V (COLAR SOMENTE VALORES)", 'bold', null, null],
    ["*NUNCA CLIQUE COM BOTÃO DIREITO", 'bold', null, null],
    ["*NUNCA CLIQUE NA CANETA (EDITAR)", 'bold', null, null],
    ["", null, null, null], // Linha em branco
    ["A) CRIAR A PLANILHA (PROPRIETÁRIO)", 'bold', 'underline', null],
    ["1- Ao final de cada ano, acessar o github: ", null, null, null],
    [repositGitHub, 'bold', null, 'italic'],
    ["2- Procurar os arquivos ScriptPlanilhaEstoqueControle e ScriptBotaoAmareloGravar com versão mais recente;", null, null, null],
    ["3- Copiar o código ScriptPlanilhaEstoqueControle;", null, null, null],
    ["4- Colar no editor Google App Script da sua conta Google;", null, null, null],
    ["5- Executar a primeiraEtapa, segundaEtapa e terceiraEtapa;", null, null, null],
    ["ATENÇÃO: A execução pela primeira vez requer permissões.", null, null, 'italic'],
    ["ATENÇÃO: Aguarde a mensagem de conclusão para executar a etapa seguinte.", null, null, 'italic'],
    ["6- Ao concluir a terceiraEtapa, abrir o editor de script da planilha recém criada (Extensões > App Script);", null, null, null],
    ["7- Copiar e colar o ScriptBotaoAmareloGravar.", null, null, null],
    ["", null, null, null], // Linha em branco
    ["B) CONFIGURAÇÃO INICIAL DA PLANILHA (PROPRIETÁRIO)", 'bold', 'underline', null],
    ["8- Ir para a aba INVENT da planilha vigente e copiar a lista de inventário invent!A3:D;", null, null, null],
    ["9- Ir para a aba INVENT da nova planilha e colar somente valores (Ctrl+Shift+V);", null, null, null],  
    ["10- Ir para a aba CONFIG da planilha vigente e copiar a lista de estoque config!C3:G;", null, null, null], 
    ["11- Ir para a aba ENTRADA da nova planilha, colar somente valores (Ctrl+Shift+V) e clicar em GRAVAR;", null, null, null],
    ["ATENÇÃO: O ScriptPlanilhaEstoqueControle clona (cria automaticamente) uma planilha para o ano seguinte.", null, null, 'italic'],
    ["ATENÇÃO: Contudo, é nessessário copiar manualmente o inventário e os produtos do estoque do ano anterior e colar na nova planilha.", null, null, 'italic'], 
    ["12- Ir para a aba CONFIG da nova planilha;", null, null, null],
    ["13- Conferir o ano na coluna \"AnoAtual\";", null, null, null],
    ["14- Digitar os setores na coluna \"Setores\";", null, null, null],
    ["ATENÇÃO: Utilize uma linha para cada setor, até o máximo de 10 setores.", null, null, 'italic'],
    ["15- Digitar código interno, e-fisco, produto e quantidade mínima para novos itens que entrarão no inventário;", null, null, null],
    ["ATENÇÃO: Nunca cadastre no inventário produto com \"espaço\" e/ou \"arroba\".", null, null, 'italic'],
    ["ATENÇÃO: Verifique periodicamente o banco de dados para incluir novos produtos, fazer correções e eliminar itens redundantes.", null, null, 'italic'],
    ["16- Ocultar a aba CONFIG (opcional);", null, null, null],
    ["17- Compartilhar a planilha com outros usuários (opcional);", null, null, null],
    ["ATENÇÃO: As fórmulas estão desprotegidas para o proprietário e protegidas para outros usuários.", null, null, 'italic'],
    ["", null, null, null], // Linha em branco
    ["C) ENTRADA DE PRODUTOS", 'bold', 'underline', null],
    ["18- Verificar o item desejado na aba DETALHADO;", null, null, null],
    ["19- Abrir a aba ENTRADA;", null, null, null],
    ["20- Ir na coluna \"Produto\" digitar as primeiras letras do produto e selecionar o item desejado;", null, null, null],
    ["21- Preencher os campos obrigatórios: fabricante, lote, data de vencimento e quantidade. Se necessário, digitar alguma observação;", null, null, null],
    ["ATENÇÃO: A data de vencimento deve ser escrita no formato dd/mm/aaaa.", null, null, 'italic'],
    ["ATENÇÃO: Os campos que não tiverem dados devem ser preenchidos com ponto.", null, null, 'italic'],
    ["ATENÇÃO: Nunca cadastre produto usando \"arroba\".", null, null, 'italic'],
    ["22- Notar que enquanto estiver preenchendo os campos, o status fica vermelho PENDENTE;", null, null, null],
    ["23- Quando concluir o preenchimento, o status torna-se amarelo GRAVAR;", null, null, null],
    ["24- Em seguida, clicar no botão amarelo GRAVAR para executar o script;", null, null, null],
    ["25- Confirmar a mensagem \"Deseja retirar fórmulas e gravar ENTRADA?\";", null, null, null],
    ["26- Não mexer, nem clicar em nada enquanto o script estiver em execução;", null, null, null],
    ["27- Se o script for encerrado por algum motivo, ou aparecer alerta \"tempo de execução excedido\", clicar novamente no botão GRAVAR;", null, null, null],
    ["28- Aguardar a mensagem de conclusão \"ENTRADA concluída!\";", null, null, null],
    ["29- Notar que as entradas que estavam amarelo GRAVAR tornam-se verdes INSERIDO.", null, null, null],
    ["ATENÇÃO: Após gravar, ainda é permitido fazer alterações. Basta deletar o dado incorreto e digitar novamente.", null, null, 'italic'],
    ["ATENÇÃO: Quando o status torna-se roxo ERRO, indica que o produto foi alterado na aba INVENT. Corrija o produto na aba ENTRADA para que fique igual ao inventário.", null, null, 'italic'],
    ["", null, null, null], // Linha em branco
    ["D) RETIRADA DE PRODUTOS", 'bold', 'underline', null],
    ["30- Abrir a aba SAÍDA", null, null, null],
    ["31- Ir na coluna \"LOTE @ VALIDADE @ PRODUTO @ FABRICANTE\", digitar o lote e selecionar o item desejado;", null, null, null],
    ["ATENÇÃO: Fique atento(a) as especificações e a unidade de medida do produto", null, null, 'italic'],
    ["32- Digitar o setor e quantidade. Se necessário, digitar alguma observação;", null, null, null],
    ["33- Notar que enquanto estiver preenchendo os campos, o status fica vermelho PENDENTE;", null, null, null],
    ["34- Quando preencher corretamente, o status torna-se amarelo GRAVAR;", null, null, null],
    ["35- Ao concluir o preenchimento da retirada, clicar no botão amarelo GRAVAR para executar o script;", null, null, null],
    ["36- Confirmar a mensagem \"Deseja retirar fórmulas e gravar SAÍDA?\";", null, null, null],
    ["37- Não mexer, nem clicar em nada enquanto o script estiver em execução;", null, null, null],
    ["38- Se o script for encerrado por algum motivo, ou aparecer alerta \"tempo de execução excedido\", clicar novamente no botão GRAVAR;", null, null, null],
    ["39- Aguardar a mensagem de conclusão \"SAÍDA concluída!\";", null, null, null],
    ["40- Notar que as entradas que estavam amarelo GRAVAR tornam-se verdes RETIRADO.", null, null, null],
    ["ATENÇÃO: A retirada de produtos é por lote! Caso deseje retirar dois lotes diferentes, será necessário fazer duas retiradas.", null, null, 'italic'],
    ["ATENÇÃO: Após enviar, ainda é permitido fazer alterações. Basta deletar o dado incorreto e digitar novamente.", null, null, 'italic'],
    ["ATENÇÃO: Quando o status torna-se roxo ERRO, indica que o produto foi alterado na aba ENTRADA. Corrija o produto na aba SAÍDA para que fique igual a entrada.", null, null, 'italic'],
    ["", null, null, null], // Linha em branco
    ["E) RELATÓRIOS DO ESTOQUE", 'bold', 'underline', null],
    ["41- Abrir a aba DETALHADO;", null, null, null],
    ["42- Conferir o status em ESTOQUE DE PRODUTOS;", null, null, null],
    ["Quant.Atual >= Quant.Mín o status é verde OK", null, null, null],
    ["Quant.Atual < Quant.Mín o status é amarelo CRÍTICO", null, null, null],
    ["Quant.Atual = Zero & Quant.Mín > Zero o status é vermelho EM FALTA", null, null, null],
    ["Quant.Atual = Zero & Quant.Mín = Zero o status é verde OK", null, null, null],
    ["Quant.Atual > Zero & Quant.Mín = Zero o status é azul FAZER DOAÇÃO", null, null, null],
    ["Quant.Atual < Zero o status é laranja FAZER REVISÃO pois não pode haver quantidades negativas no estoque", null, null, null],
    ["43- Conferir periodicamente as tabelas CONSUMO SETORIAL, CONSUMO MENSAL, ABASTECIMENTO MENSAL e LOTES.", null, null, null],
    ["", null, null, null], // Linha em branco
    ["F) LISTA DE INVENTÁRIO:", 'bold', 'underline', null],  
  ];

  // Adicionar as instruções 
  var numLinhas = instrucoes.length; // quant de linhas das instruções
  var leia = ss.getSheetByName("Leia-me"); // Acessar a aba
  leia.getRange(1, 1, numLinhas, 1).setValues(instrucoes.map(function(row) {
    return [row[0]]; // Pega apenas o texto da instrução
  }));

  // Aplicar formatação (negrito, sublinhado e itálico)
  for (var i = 0; i < numLinhas; i++) {
    var range = leia.getRange(i + 1, 1);
    if (instrucoes[i][1] === 'bold') {
      range.setFontWeight('bold');
    }
    if (instrucoes[i][2] === 'underline') {
      range.setFontLine('underline');
    }
    if (instrucoes[i][3] === 'italic') {
      range.setFontStyle('italic');
    }
  }

  // Adicionar fórmula pontual (gerar lista inventário em italico)
  var linhaFormula = numLinhas + 1; // Definir 1ª celula em branco após as instruções
  var maxLinha = leia.getMaxRows(); // Definir máximo de linhas 
  leia.getRange(linhaFormula,1).setFormula(formula35); // adicionar formula inventário na coluna 1
  leia.getRange(linhaFormula,1,maxLinha,1).setFontStyle('italic'); // torna itálico apos as instruções
  
  // Converter string repositGitHub para link 
  var string = leia.getRange('A:A').createTextFinder(repositGitHub); // criar buscador para a string
  var procurarString = string.findAll(); // fazer busca pela string
  for (var i = 0; i < procurarString.length; i++) { // loop
    //Listar intervalos A*:A* contendo string
    var intervaloString = "A"+procurarString[i].getRow()+":A"+procurarString[i].getRow();
    var listaIntervaloString = leia.getRange(intervaloString);
  }  

  //Substituir string pelo link
  listaIntervaloString.setFormula('=HYPERLINK("' + repositGitHub + '")');



  // *** FINALIZAR SEGUNDA ETAPA ***
  // Forçar a atualização da planilha
  SpreadsheetApp.flush();

  // Acessar "propriedades do script" para salvar informações para a próxima etapa
  PropertiesService.getScriptProperties().setProperty("versaoData",versaoData);
  PropertiesService.getScriptProperties().setProperty("repositGitHub",repositGitHub);
  PropertiesService.getScriptProperties().setProperty("id",id);
  PropertiesService.getScriptProperties().setProperty("nomePlanilha",nomePlanilha);
  PropertiesService.getScriptProperties().setProperty("imageUrl",imageUrl);
  PropertiesService.getScriptProperties().setProperty("cellWidthEntrada",cellWidthEntrada);
  PropertiesService.getScriptProperties().setProperty("cellWidthSaida",cellWidthSaida);
  PropertiesService.getScriptProperties().setProperty("anoSeguinte",anoSeguinte);
  
  // Mensagem de conclusão
  Logger.log(
   "SEGUNDA ETAPA finalizada com sucesso!\n" +
   "Planilha " + nomePlanilha + " ID: " + id + "\n" +
   "Pronta para prosseguir para a TERCEIRA ETAPA."
  );

  
}  




//TERCEIRA ETAPA: Configurar os botões e adicionar proteção para abas 
function terceiraEtapa() {

  // Acessar "propriedades do script" para recuperar as informações
  const versaoData = PropertiesService.getScriptProperties().getProperty("versaoData");
  const repositGitHub = PropertiesService.getScriptProperties().getProperty("repositGitHub");
  const id = PropertiesService.getScriptProperties().getProperty("id");
  const nomePlanilha = PropertiesService.getScriptProperties().getProperty("nomePlanilha");
  const imageUrl = PropertiesService.getScriptProperties().getProperty("imageUrl");
  const cellWidthEntradaString = PropertiesService.getScriptProperties().getProperty("cellWidthEntrada");
  const cellWidthSaidaString = PropertiesService.getScriptProperties().getProperty("cellWidthSaida");
  const anoSeguinteString = PropertiesService.getScriptProperties().getProperty("anoSeguinte");
  
  // Converter para inteiro (base 10) 
  const anoSeguinte = parseInt(anoSeguinteString, 10);   
  const cellWidthEntrada = parseInt(cellWidthEntradaString, 10);   
  const cellWidthSaida = parseInt(cellWidthSaidaString, 10); 

  // *** Definições gerais ***
  const ss = SpreadsheetApp.openById(id); // acessar planilha
  const propri = Session.getActiveUser().getEmail(); // Obter conta gmail do proprietário (você)
   
  // *** Formatação da aba "Leia-me" ***
  var leia = ss.getSheetByName("Leia-me"); // Acessar a aba
  var protection = leia.protect(); // Acessar a proteção da aba
  protection.removeEditors(protection.getEditors()); // Remove todos os editores
  protection.addEditor(propri); // Adiciona o proprietário (você) como editor

  // *** Formatação da aba "DETALHADO" ***
  var detalh = ss.getSheetByName("DETALHADO"); // Acessar a aba
  var protection = detalh.protect(); // Acessar a proteção da aba
  protection.removeEditors(protection.getEditors()); // Remove todos os editores
  protection.addEditor(propri); // Adiciona o proprietário (você) como editor

  // *** Formatação da aba "config" ***
  var config = ss.getSheetByName("config"); // Acessar a aba
  var protection = config.protect(); // Acessar a proteção da aba
  protection.removeEditors(protection.getEditors()); // Remove todos os editores
  protection.addEditor(propri); // Adiciona o proprietário (você) como editor
   


  // *** Formatação da aba "invent" ***
  var invent = ss.getSheetByName("invent"); // Acessar a aba

  // Acessar a proteção da aba
  var protectionAba = invent.protect().setDescription('ProtegerAba');
  
  // Definir os intervalos a serem desprotegidos
  var intervaloA3D = invent.getRange("A3:D");
  var intervalosDesprotegidos = [intervaloA3D]; // Crie um array de objetos Range

  // Proteger o intervalo para usuários exceto proprietario
  protectionAba.removeEditors(protection.getEditors()); // remover editores
  protectionAba.addEditor(propri); // adicionar proprietario
  protectionAba.setUnprotectedRanges(intervalosDesprotegidos); // desproteger intervalo para os usuarios




  // *** Formatação da aba "ENTRADA" ***
  var entrada = ss.getSheetByName("ENTRADA"); // Acessar a aba
  
  // Acessar a proteção da aba
  var protectionAba = entrada.protect().setDescription('ProtegerAba');
  
  // Definir os intervalos a serem desprotegidos
  var intervalo1 = entrada.getRange("A2:B2");
  var intervalo2 = entrada.getRange("D3:N");

  // Definir os intervalos a serem desprotegidos usando um array contendo os objetos Range
  var intervalosDesprotegidos = [intervalo1, intervalo2];

  // Proteger o intervalo para usuários exceto proprietario
  protectionAba.removeEditors(protection.getEditors()); // remover editores
  protectionAba.addEditor(propri); // adicionar proprietario
  protectionAba.setUnprotectedRanges(intervalosDesprotegidos); // desproteger intervalo para os usuarios

  // Inserir a imagem ao botão amarelo (inicialmente sem redimensionar)
  var imagemBlob = UrlFetchApp.fetch(imageUrl).getBlob(); // Transforma a imagem para blob (binário)
  const imagem1 = entrada.insertImage(imagemBlob, 1, 2); // coluna (1) linha (2)

  // Obter altura e largura
  var cellHeight = 2 * entrada.getRowHeight(2); // Obter o dobro da altura da segunda linha
   
  // Redimensiona a imagem
  imagem1.setWidth(cellWidthEntrada);
  imagem1.setHeight(cellHeight);
  
  // Associa o script "gravarEntrada" à imagem do botão amarelo
  imagem1.assignScript('gravarEntrada');




  // *** Formatação da aba "SAÍDA" ***
  var saida = ss.getSheetByName("SAÍDA"); // Acessar a aba

  // Acessar a proteção da aba
  var protectionAba = saida.protect().setDescription('ProtegerAba');
  
  // Definir os intervalos a serem desprotegidos
  var intervalo1 = saida.getRange("A2:B2");
  var intervalo2 = saida.getRange("D3:K");

  // Definir os intervalos a serem desprotegidos usando um array contendo os objetos Range
  var intervalosDesprotegidos = [intervalo1, intervalo2];

  // Proteger o intervalo para usuários exceto proprietario
  protectionAba.removeEditors(protection.getEditors()); // remover editores
  protectionAba.addEditor(propri); // adicionar proprietario
  protectionAba.setUnprotectedRanges(intervalosDesprotegidos); // desproteger intervalo para os usuarios

  // Inserir a imagem ao botão amarelo (inicialmente sem redimensionar)
  var imagemBlob = UrlFetchApp.fetch(imageUrl).getBlob(); // Transforma a imagem para blob (binário)
  const imagem2 = saida.insertImage(imagemBlob, 1, 2); // coluna (1) linha (2)

  // Obter altura e largura
  var cellHeight = 2 * saida.getRowHeight(2); // Obter o dobro da altura da segunda linha
   
  // Redimensiona a imagem
  imagem2.setWidth(cellWidthSaida);
  imagem2.setHeight(cellHeight);

  // Associa o script "gravarSaida" à imagem do botão amarelo
  imagem2.assignScript('gravarSaida');




  // *** FINALIZAR TERCEIRA ETAPA ***
  // Forçar a atualização da planilha
  SpreadsheetApp.flush();

  // Acessar "propriedades do script" para salvar informações para a próxima etapa
  PropertiesService.getScriptProperties().deleteProperty("versaoData");
  PropertiesService.getScriptProperties().deleteProperty("repositGitHub");
  PropertiesService.getScriptProperties().deleteProperty("id");
  PropertiesService.getScriptProperties().deleteProperty("nomePlanilha");
  PropertiesService.getScriptProperties().deleteProperty("imageUrl");
  PropertiesService.getScriptProperties().deleteProperty("cellWidthEntrada");
  PropertiesService.getScriptProperties().deleteProperty("cellWidthSaida");
  PropertiesService.getScriptProperties().deleteProperty("anoSeguinte");
  
// Mensagem de conclusão
  Logger.log(
   "TERCEIRA ETAPA finalizada com sucesso!\n" +
   "Planilha " + nomePlanilha + " ID: " + id + "\n" +
   "Pronta para receber o script!\n" +
   "Acesse o editor de script da planilha recém criada e cole o scriptBotõesGravarEntradaSaida."
  );

  
}


