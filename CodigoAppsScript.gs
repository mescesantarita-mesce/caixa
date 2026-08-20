/*
 * MESCE - Conexão com o Google Sheets
 * ====================================
 * 1) Abra a planilha no Google Sheets.
 * 2) Menu: Extensões > Apps Script.
 * 3) Apague o conteúdo e cole este arquivo inteiro.
 * 4) Implantar > Gerenciar implantações > Editar (lápis) > Versão:
 *    "Nova versão" > Salvar. (Assim a URL do config.js continua a mesma.)
 * 5) Autorize as permissões, se solicitado.
 *
 * Este script:
 *   - grava movimentações na aba "mapa";
 *   - lê a aba "mapa" ou "contas" (GET ?aba=contas);
 *   - cadastra novas contas na aba "contas" (POST acao=nova_conta);
 *   - limpa linhas inválidas gravadas por engano (opcional, manual).
 */

var ABA_MAPA = 'mapa';
var ABA_CONTAS = 'contas';
var ABA_MINISTROS = 'ministros';
var ABA_CONFIG = 'config';

function doGet(e) {
  var aba = (e && e.parameter && e.parameter.aba) || ABA_MAPA;
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var folha = planilha.getSheetByName(aba);
  if (!folha) {
    return jsonResposta({ status: 'erro', mensagem: 'Aba "' + aba + '" não encontrada.' });
  }
  var valores = folha.getDataRange().getValues();

  // Aba "contas": retorna sempre [{ nome, categoria }], ignorando cabeçalho.
  if (aba === ABA_CONTAS) {
    var contas = valores.filter(function (linha) {
      var nome = String(linha[0] || '').trim();
      if (!nome) return false;
      return ['conta', 'nome', 'tipo', 'categoria'].indexOf(nome.toLowerCase()) === -1;
    }).map(function (linha) {
      return {
        nome: String(linha[0] || '').trim(),
        categoria: String(linha[1] || '').trim() || 'Entrada'
      };
    });
    return jsonResposta(contas);
  }

  // Aba "ministros": retorna [{ nome, contato }], ignorando cabeçalho.
  if (aba === ABA_MINISTROS) {
    var ministros = valores.filter(function (linha) {
      var nome = String(linha[0] || '').trim();
      if (!nome) return false;
      return ['nome', 'ministro', 'ministros', 'mesce'].indexOf(nome.toLowerCase()) === -1;
    }).map(function (linha) {
      return {
        nome: String(linha[0] || '').trim(),
        contato: String(linha[1] || '').trim()
      };
    });
    return jsonResposta(ministros);
  }

  // Aba "config": retorna as configurações (ex.: valor da mensalidade).
  if (aba === ABA_CONFIG) {
    return jsonResposta(lerConfig());
  }

  // Aba "mapa": retorna objetos com o nome das colunas como chave e o
  // número da linha na planilha (_linha) para edição/exclusão.
  var cabecalho = valores.shift();
  var linhas = valores.map(function (linha, i) {
    var obj = { _linha: i + 2 };
    cabecalho.forEach(function (coluna, j) {
      obj[String(coluna).trim()] = linha[j];
    });
    return obj;
  });
  return jsonResposta(linhas);
}

function doPost(e) {
  try {
    var dados = JSON.parse(e.postData.contents);
    var acao = dados.acao;
    if (acao === 'nova_conta') {
      return novaConta(dados);
    }
    if (acao === 'nova_ministro') {
      return novoMinistro(dados);
    }
    if (acao === 'excluir_ministro') {
      return excluirMinistro(dados);
    }
    if (acao === 'editar_transacao') {
      return editarTransacao(dados);
    }
    if (acao === 'excluir_transacao') {
      return excluirTransacao(dados);
    }
    if (acao === 'pagamento_recorrente') {
      return pagamentoRecorrente(dados);
    }
    if (acao === 'salvar_config') {
      return salvarConfig(dados);
    }
    return novaMovimentacao(dados);
  } catch (erro) {
    return jsonResposta({ status: 'erro', mensagem: String(erro) });
  }
}

function novaMovimentacao(dados) {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var mapa = planilha.getSheetByName(ABA_MAPA);
  if (!mapa) {
    mapa = planilha.getSheets()[0];
  }
  mapa.appendRow([
    converterData(dados.data),
    String(dados.mesce || ''),
    converterValor(dados.valor),
    String(dados.periodo || ''),
    String(dados.conta || ''),
    String(dados.tipo || '')
  ]);
  return jsonResposta({ status: 'ok', mensagem: 'Registro salvo com sucesso!' });
}

function novoMinistro(dados) {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var folha = planilha.getSheetByName(ABA_MINISTROS);
  if (!folha) {
    return jsonResposta({ status: 'erro', mensagem: 'Aba "ministros" não encontrada. Crie essa aba na planilha.' });
  }
  var nome = String(dados.nome || '').trim();
  if (!nome) {
    return jsonResposta({ status: 'erro', mensagem: 'Informe o nome do ministro.' });
  }
  var contato = String(dados.contato || '').trim();
  folha.appendRow([nome, contato]);
  return jsonResposta({ status: 'ok', mensagem: 'Ministro "' + nome + '" cadastrado com sucesso!' });
}

function editarTransacao(dados) {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var mapa = planilha.getSheetByName(ABA_MAPA) || planilha.getSheets()[0];
  var linha = Number(dados.linha);
  if (!linha || linha < 2) {
    return jsonResposta({ status: 'erro', mensagem: 'Linha inválida para edição.' });
  }
  var valores = [
    converterData(dados.data),
    String(dados.mesce || ''),
    converterValor(dados.valor),
    String(dados.periodo || ''),
    String(dados.conta || ''),
    String(dados.tipo || '')
  ];
  mapa.getRange(linha, 1, 1, valores.length).setValues([valores]);
  return jsonResposta({ status: 'ok', mensagem: 'Movimentação atualizada com sucesso!' });
}

function excluirTransacao(dados) {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var mapa = planilha.getSheetByName(ABA_MAPA) || planilha.getSheets()[0];
  var linha = Number(dados.linha);
  if (!linha || linha < 2) {
    return jsonResposta({ status: 'erro', mensagem: 'Linha inválida para exclusão.' });
  }
  mapa.deleteRow(linha);
  return jsonResposta({ status: 'ok', mensagem: 'Movimentação excluída com sucesso!' });
}

// Pagamento recorrente: grava N lançamentos no mapa (um por mês),
// cada um com o valor dividido pela quantidade de meses.
function pagamentoRecorrente(dados) {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var mapa = planilha.getSheetByName(ABA_MAPA) || planilha.getSheets()[0];
  var mesce = String(dados.mesce || '').trim();
  var meses = dados.meses || [];
  var valorUnitario = Number(dados.valor_unitario);
  var data = converterData(dados.data);
  var conta = String(dados.conta || '').trim() || 'Mensalidade';
  var tipo = String(dados.tipo || '').trim() || 'Entrada';

  if (!mesce) {
    return jsonResposta({ status: 'erro', mensagem: 'Informe o nome do ministro (Mesce).' });
  }
  if (!meses.length) {
    return jsonResposta({ status: 'erro', mensagem: 'Informe a quantidade de meses.' });
  }
  if (isNaN(valorUnitario) || valorUnitario <= 0) {
    return jsonResposta({ status: 'erro', mensagem: 'Informe um valor válido por mês.' });
  }

  meses.forEach(function (mes) {
    mapa.appendRow([data, mesce, valorUnitario, String(mes), conta, tipo]);
  });

  // Se solicitado, atualiza o valor padrão da mensalidade na aba "config".
  if (dados.atualizar_mensalidade && Number(dados.valor_mensalidade) > 0) {
    salvarConfig({ chave: 'valor_mensalidade', valor: Number(dados.valor_mensalidade) });
  }

  var total = Math.round(valorUnitario * meses.length * 100) / 100;
  return jsonResposta({
    status: 'ok',
    mensagem: 'Foram registradas ' + meses.length + ' mensalidades de ' +
      valorUnitario.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' }) +
      ' para ' + mesce + ' (' + meses[0] + ' a ' + meses[meses.length - 1] + '). Total: ' +
      total.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' }) + '.'
  });
}

// Lê a aba "config" (chave/valor). Cria com padrão se não existir.
function lerConfig() {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var folha = planilha.getSheetByName(ABA_CONFIG);
  var config = { valor_mensalidade: 5 };
  if (folha) {
    var valores = folha.getDataRange().getValues();
    for (var i = 0; i < valores.length; i++) {
      var chave = String(valores[i][0] || '').trim().toLowerCase();
      if (chave === 'valor_mensalidade') {
        var num = parseFloat(String(valores[i][1]).replace(',', '.'));
        if (!isNaN(num) && num > 0) config.valor_mensalidade = num;
      }
    }
  }
  return config;
}

// Grava ou atualiza uma chave na aba "config".
function salvarConfig(dados) {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var folha = planilha.getSheetByName(ABA_CONFIG);
  if (!folha) {
    folha = planilha.insertSheet(ABA_CONFIG);
    folha.appendRow(['chave', 'valor']);
  }
  var chave = String(dados.chave || '').trim().toLowerCase();
  var valor = Number(dados.valor);
  if (!chave || isNaN(valor)) {
    return jsonResposta({ status: 'erro', mensagem: 'Chave ou valor inválido.' });
  }
  var valores = folha.getDataRange().getValues();
  for (var i = 0; i < valores.length; i++) {
    if (String(valores[i][0] || '').trim().toLowerCase() === chave) {
      folha.getRange(i + 1, 2).setValue(valor);
      return jsonResposta({ status: 'ok', mensagem: 'Configuração atualizada!' });
    }
  }
  folha.appendRow([chave, valor]);
  return jsonResposta({ status: 'ok', mensagem: 'Configuração atualizada!' });
}

function excluirMinistro(dados) {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var folha = planilha.getSheetByName(ABA_MINISTROS);
  if (!folha) {
    return jsonResposta({ status: 'erro', mensagem: 'Aba "ministros" não encontrada.' });
  }
  var nome = String(dados.nome || '').trim();
  if (!nome) {
    return jsonResposta({ status: 'erro', mensagem: 'Informe o nome do ministro.' });
  }
  var valores = folha.getDataRange().getValues();
  var linha = -1;
  for (var i = 0; i < valores.length; i++) {
    if (String(valores[i][0] || '').trim().toLowerCase() === nome.toLowerCase()) {
      linha = i + 1;
      break;
    }
  }
  if (linha === -1) {
    return jsonResposta({ status: 'erro', mensagem: 'Ministro "' + nome + '" não encontrado na aba "ministros".' });
  }
  folha.deleteRow(linha);
  return jsonResposta({ status: 'ok', mensagem: 'Ministro "' + nome + '" excluído com sucesso!' });
}

function novaConta(dados) {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var contas = planilha.getSheetByName(ABA_CONTAS);
  if (!contas) {
    return jsonResposta({ status: 'erro', mensagem: 'Aba "contas" não encontrada na planilha.' });
  }
  var nome = String(dados.nome || '').trim();
  if (!nome) {
    return jsonResposta({ status: 'erro', mensagem: 'Informe o nome da conta.' });
  }
  var categoria = String(dados.categoria || '').trim() || 'Entrada';
  contas.appendRow([nome, categoria]);
  return jsonResposta({ status: 'ok', mensagem: 'Conta "' + nome + '" cadastrada com sucesso!' });
}

// "dd/mm/aaaa" -> Data real do Google Sheets.
function converterData(d) {
  var partes = String(d || '').split('/');
  if (partes.length === 3) {
    var dia = parseInt(partes[0], 10);
    var mes = parseInt(partes[1], 10) - 1;
    var ano = parseInt(partes[2], 10);
    if (!isNaN(dia) && !isNaN(mes) && !isNaN(ano)) {
      return new Date(ano, mes, dia);
    }
  }
  return String(d || '');
}

// Converte o valor para número (mesmo formato das linhas antigas da planilha).
function converterValor(v) {
  if (typeof v === 'number') return v;
  var s = String(v).trim().replace(/[R$\s]/g, '').replace(/\./g, '').replace(',', '.');
  var num = parseFloat(s);
  return isNaN(num) ? String(v) : num;
}

/*
 * OPCIONAL: remove linhas inválidas gravadas por engano na aba "mapa"
 * (ex.: células com "undefined").
 * Como usar: no editor do Apps Script, escolha esta função na lista e
 * clique em "Executar". Depois pode apagar esta função do projeto.
 */
function limparRegistrosInvalidos() {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var mapa = planilha.getSheetByName(ABA_MAPA) || planilha.getSheets()[0];
  var valores = mapa.getDataRange().getValues();
  var linhasParaRemover = [];
  for (var i = valores.length - 1; i > 0; i--) {
    var v = valores[i];
    var data = String(v[0] || '').trim();
    var mesce = String(v[1] || '').trim();
    var valor = String(v[2] || '').trim();
    if (!data || !mesce || !valor ||
        /undefined/i.test(data) || /undefined/i.test(mesce) || /undefined/i.test(valor)) {
      linhasParaRemover.push(i + 1);
    }
  }
  if (linhasParaRemover.length) {
    linhasParaRemover.forEach(function (linha) { mapa.deleteRow(linha); });
    Logger.log('Removidas ' + linhasParaRemover.length + ' linha(s) inválida(s).');
  } else {
    Logger.log('Nenhuma linha inválida encontrada.');
  }
}

function jsonResposta(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}