/* Código para automatizar Carteira de Ações
Escrito por Lucas do Vale Bezerra, 
Engenheiro de Computação formado pelo ITA em 2022
*/

// Função que busca elemento e retorna o nó
function buscarElemento(papeis, papel) {
  var i;
  for(i = 0; i < papeis.length; i++) {
    if (papel == papeis[i]) {
      return i;
    }
  }
  if (i == papeis.length)
    return 0;
};

// Função que busca elemento e retorna índice no vetor
function inserirElemento(papeis, qtd, pm, papel, q, p) {
  papeis.push(papel);
  qtd.push(q);
  pm.push(p);
};

// Função que busca elemento e retorna índice no vetor
function removerElemento(papeis, qtd, pm, papel) {
  pos = buscarElemento(papel);
  papeis.splice(pos,1);
  qtd.splice(pos,1);
  pm.splice(pos,1);
};

// Função que muda de página na planilha
function mudarDePag(pag) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  spreadsheet.setActiveSheet(pag, true);
};

// Função que exclui linhas da carteira
function excluirLinha() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('B5:J5').activate();
  spreadsheet.getRange('B5:J5').deleteCells(SpreadsheetApp.Dimension.ROWS);
};

// Função que insere linhas na carteira
function inserirLinha() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('B5:J5').activate();
  spreadsheet.getRange('B5:J5').insertCells(SpreadsheetApp.Dimension.ROWS);
};

// Limpando o preenchimento da Operação
function limparPreenchimento() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('J7:M8').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
};

// Função que atualiza o Extrato
function atualizarExtrato() {
  
  // Declarando planilhas
  var planilha = SpreadsheetApp.getActive();
  var extrato = planilha.getSheetByName('Extrato');
  
  // Lógica da Macro
  extrato.getRange('B7:G7').activate();
  extrato.getRange('B7:G7').insertCells(SpreadsheetApp.Dimension.ROWS);
  extrato.getRange('B7:G7').insertCells(SpreadsheetApp.Dimension.ROWS);
  extrato.getRange('B7').activate();
  extrato.getRange('I7:M8').copyTo(extrato.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  extrato.getRange('B8:G8').activate();
  extrato.getRange('B8:G8').deleteCells(SpreadsheetApp.Dimension.ROWS);
  extrato.getRange('G8').activate();
  extrato.getActiveRange().autoFill(extrato.getRange('G7:G8'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  extrato.getRange('I7:I8').activate();
  limparPreenchimento();
};

// Função que atualiza a Carteira de Ações
function atualizarCarteira() {

  atualizarExtrato();
  // Arrays que armazenarão as leituras dos papeis, quantidade e preço médio das ações
  var papeis = [0];
  var qtd = [0];
  var pm = [0];
    
  // Declarando planilhas
  var planilha = SpreadsheetApp.getActive();
  var extrato = planilha.getSheetByName('Extrato');
  var carteira = planilha.getSheetByName('Carteira');
  var dataExtrato = extrato.getDataRange().getValues();
  var dataCarteira = carteira.getDataRange().getValues();
  
  // Declarando variáveis para cálculos
  var lin, col, i, j, ultimo;
  var papel = 0;
  var qt = 0;
  var preco = 0
  var op;
  
  // Mudando para a planilha da carteira
  mudarDePag(carteira);
  lin = 5;
  col = 2;
  var cont = 0;
  var palavra = dataCarteira[lin - 1][col - 1];
  
  // Verificando se na carteira há linhas suficientes para printar
  while (palavra != 'Caixa Inicial') {
    cont++;
    lin++;
    palavra = dataCarteira[lin - 1][col - 1];
  }
  
  // Lendo o caixa
  var caixaInicial = planilha.getActiveSheet().getRange(lin + 2, col).getValue();
  var caixa = caixaInicial;
  
  // Extraindo os dados do Extrato
  for (lin = dataExtrato.length; lin > 6; lin--) {
    for (col = 3; col < 7 ; col++) {
      if (col == 3) {
        op = dataExtrato[lin - 1][col - 1];
      }
      if (col == 4) {
        papel = dataExtrato[lin - 1][col - 1];
      }
      if (col == 5) {
        preco = dataExtrato[lin - 1][col - 1];
      }
      if (col == 6) {
        qt = dataExtrato[lin - 1][col - 1];
      }
    }
    
    // Procura se papel já tá na carteira
    pos = buscarElemento(papeis, papel);
    if (pos != 0) { // Caso em que papel está na carteira
      if (op == 'Compra') { // Caso em que comprou um papel
        caixa -= qt * preco;
        pm[pos] = (qtd[pos] * pm[pos] + qt * preco) / (qtd[pos] + qt);
        qtd[pos] = qtd[pos] + qt;
      }
      if (op == 'Venda') { // Caso em que vendeu um papel
        caixa += qt * preco;
        pm[pos] = (qtd[pos] * pm[pos] - qt * preco) / (qtd[pos] - qt);
        qtd[pos] = qtd[pos] - qt;
        if (qtd[pos] == 0) {
          papeis.splice(pos,1);
          qtd.splice(pos,1);
          pm.splice(pos,1);
        }
      }
    }
    else { // Caso em que papel não está na carteira
      if (op == 'Compra') {
        caixa -= qt * preco;
        papeis.push(papel);
        qtd.push(qt);
        pm.push(preco);
      }
    }
    // Caso tenha lido todo o extrato, sai da iteração
    if (dataExtrato[lin - 1][col - 1] == ' ') {
      break;
    }
  }
  
  // Revisando carteira
  for (pos = 0; pos < qtd.length; pos++) {
    if (qtd[pos] == 0) {
      papel = papeis[pos];
      removerElemento(papeis, qtd, pm, papel)
    }
  }
  
  // Mudando para a planilha da carteira
  mudarDePag(carteira);
  lin = 5;
  col = 2;
  var cont = 0;
  var palavra = dataCarteira[lin - 1][col - 1];
  
  // Verificando se na carteira há linhas suficientes para printar
  while (palavra != 'Caixa Inicial') {
    cont++;
    lin++;
    palavra = dataCarteira[lin - 1][col - 1];
  }
  // Caso em que há mais linhas que o necessário
  if (cont > papeis.length) {
    for (i = 0; i < (cont - papeis.length); i++) {
      excluirLinha();
    }
  }
  // Caso em que há menos linhas que o necessário
  if (cont < papeis.length) {
    for (i = 0; i < (papeis.length - cont); i++) {
      inserirLinha();
    }
  }
  // Printando as ações na carteira
  pos = 0;
  palavra = '=';
  carteira.getRange('B5').activate();
  for (lin = 5; lin < (5 + papeis.length); lin++) {
    for (col = 2; col < 11; col++) {
      if (col == 2) {
        planilha.getActiveSheet().getRange(lin,col).setValue(papeis[pos]);
        planilha.getActiveSheet().getRange(lin + 1,col + 10).setValue(papeis[pos]);
      }
      if (col == 3)
        planilha.getActiveSheet().getRange(lin,col).setValue(qtd[pos]);
      if (col == 4)
        planilha.getActiveSheet().getRange(lin,col).setValue(pm[pos]);
      if (col == 5) {
        if(papeis[lin - 5].length <= 6)
          planilha.getActiveSheet().getRange(lin,col).setValue('=GOOGLEFINANCE(B' + lin + ')');
        else planilha.getActiveSheet().getRange(lin,col).setValue('=SE(B' + lin + '<>"";PROCV(B' + lin + ';\'Opções\'!D:I;4;FALSO);"")');
      }
      if (col == 6)
        planilha.getActiveSheet().getRange(lin,col).setValue('=C' + lin + ' * D' + lin);
      if (col == 7)
        planilha.getActiveSheet().getRange(lin,col).setValue('= (E' + lin + ' - D' + lin + ') * C' + lin);
      if (col == 8)
        planilha.getActiveSheet().getRange(lin,col).setValue('= G' + lin + ' /F' + lin);
      if (col == 9) 
        planilha.getActiveSheet().getRange(lin,col).setValue('=SEERRO(PROCV(B' + lin + ';\'Papeis da bolsa\'!$A$2:$C$887;3;FALSO);"Opção")');
      if (col == 10) {
        planilha.getActiveSheet().getRange(lin,col).setValue('= (F' + lin + ' + G' + lin + ')/$I$' + (papeis.length + 7));
        planilha.getActiveSheet().getRange(lin + 1,col + 3).setValue('= (F' + lin + ' + G' + lin + ')/$I$' + (papeis.length + 7));
      }
    }
    palavra += ' + C' + lin + ' * E' + lin;
    pos++;
  }
  planilha.getActiveSheet().getRange(lin + 2, col - 2).setValue(palavra);
  planilha.getActiveSheet().getRange(lin + 2, col - 5).setValue(caixa);
};