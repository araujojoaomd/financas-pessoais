/**
 * Finanças Pessoais — Backend (Apps Script)
 * Banco: a própria planilha onde este script está vinculado.
 *
 * Implantação: Implantar > Nova implantação > App da Web
 *   - Executar como: Eu (dono da planilha)
 *   - Quem pode acessar: Qualquer pessoa
 * Ao republicar, use "Gerenciar implantações" > editar (lápis) > Nova versão,
 * pra manter a MESMA URL.
 */

// ===== CONFIG =====
// Emails autorizados a usar o app (João e Lívia). Ajuste conforme as contas Google reais.
var EMAILS_AUTORIZADOS = [
  'araujo.joaomd@gmail.com',
  'joao.marcelo@recife.pe.gov.br'
  // , 'email-da-livia@gmail.com'
];

// Client ID OAuth (mesmo do painel SETRI; o token é validado contra ele).
var CLIENT_ID = '864961214405-bipsb4fk0sr0sgi8m9ei81u87rhjshc4.apps.googleusercontent.com';

// Definição das abas e seus cabeçalhos. A ORDEM aqui é a ordem das colunas.
var ABAS = {
  Contas:         ['ID','Nome','Perfil','SaldoInicial','LimiteCheque','Ativo'],
  Cartoes:        ['ID','Nome','Perfil','DiaFechamento','DiaVencimento','Limite','Ativo'],
  Compras_Cartao: ['ID','Codigo','Perfil','Natureza','Cartao','DataCompra','Descricao','Grupo','Categoria','ValorParcela','QtdParcelas','CompetenciaInicial','AlvoRef','Criado'],
  Recorrentes:    ['ID','Perfil','Tipo','Grupo','Descricao','Categoria','Conta','ValorPadrao','DiaVenc','VigenciaInicio','VigenciaFim','Ativo','EmFolha','SalarioRef'],
  Movimentos:     ['ID','Codigo','Perfil','Competencia','CompGasto','Data','Tipo','Grupo','Categoria','Conta','Descricao','Valor','Status','Origem','Ref'],
  Emprestimos:    ['ID','Perfil','Instituicao','Contrato','DataContratacao','ValorContratado','TxNominal','TxEfetiva','ValorParcela','QtdParcelas','ParcelasPagas','CompetenciaInicial','Tipo','SalarioRef'],
  Categorias:     ['ID','Tipo','Nivel','Nome','Ativo'],
  Config:         ['ID','Valor']
};

// ===== SETUP (rode uma vez no editor) =====
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(ABAS).forEach(function(nome) {
    var sh = ss.getSheetByName(nome);
    if (!sh) sh = ss.insertSheet(nome);
    var headers = ABAS[nome];
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#1F3864').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
  });
  // Semeia Grupos/Categorias na primeira vez (só se a aba estiver vazia).
  semearCategorias_(ss);
  // Remove a aba padrão "Página1"/"Sheet1" se estiver vazia
  ['Página1','Sheet1','Folha1'].forEach(function(n) {
    var s = ss.getSheetByName(n);
    if (s && s.getLastRow() <= 1 && Object.keys(ABAS).indexOf(n) === -1) {
      try { ss.deleteSheet(s); } catch (e) {}
    }
  });
  return 'Setup concluído: ' + Object.keys(ABAS).join(', ');
}

// Preenche a aba Categorias com os grupos/categorias padrão, só se ainda não houver dados.
function semearCategorias_(ss) {
  var sh = ss.getSheetByName('Categorias');
  if (!sh || sh.getLastRow() > 1) return; // já tem dados
  var seed = [
    ['ENTRADA','GRUPO','SALÁRIO'],
    ['ENTRADA','GRUPO','OUTRAS RENDAS'],
    ['SAÍDA','GRUPO','CARRO'],
    ['SAÍDA','GRUPO','APT PARNAMIRIM'],
    ['SAÍDA','GRUPO','FLAT CARNEIROS'],
    ['SAÍDA','GRUPO','APT JAQUEIRA'],
    ['SAÍDA','GRUPO','CASA ALDEIA'],
    ['SAÍDA','GRUPO','GASTOS PESSOAIS'],
    ['SAÍDA','CATEGORIA','ALIMENTAÇÃO'],
    ['SAÍDA','CATEGORIA','MORADIA'],
    ['SAÍDA','CATEGORIA','TRANSPORTE'],
    ['SAÍDA','CATEGORIA','SAÚDE'],
    ['SAÍDA','CATEGORIA','GASTO PESSOAL'],
    ['SAÍDA','CATEGORIA','OBRAS CASA DE ALDEIA'],
    ['SAÍDA','CATEGORIA','SAQUE']
  ];
  var linhas = seed.map(function(r, i) {
    return ['CG-' + (i + 1), r[0], r[1], r[2], 'SIM'];
  });
  sh.getRange(2, 1, linhas.length, 5).setValues(linhas);
}

// ===== ENTRADA HTTP =====
function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents || '{}');
    var auth = verificar(body.idToken);
    if (!auth.ok) return json(auth);

    var acao = body.acao;
    var r;
    switch (acao) {
      case 'ler':              r = lerTudo(); break;
      case 'salvar':           r = inserir(body.aba, body.registro); break;
      case 'editar':           r = editar(body.aba, body.registro); break;
      case 'excluir':          r = excluir(body.aba, body.id); break;
      default: r = { ok: false, erro: 'acao_desconhecida: ' + acao };
    }
    r.usuario = auth.email;
    return json(r);
  } catch (err) {
    return json({ ok: false, erro: String(err) });
  }
}

function doGet() {
  return json({ ok: true, msg: 'Finanças Pessoais API online' });
}

// ===== AUTENTICAÇÃO =====
function verificar(idToken) {
  if (!idToken) return { ok: false, erro: 'sem_token' };
  try {
    var resp = UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(idToken));
    var info = JSON.parse(resp.getContentText());
    if (info.aud !== CLIENT_ID) return { ok: false, erro: 'token_invalido' };
    var email = (info.email || '').toLowerCase();
    if (EMAILS_AUTORIZADOS.map(function(x){return x.toLowerCase();}).indexOf(email) === -1) {
      return { ok: false, erro: 'sem_acesso' };
    }
    return { ok: true, email: email };
  } catch (err) {
    return { ok: false, erro: 'token_invalido' };
  }
}

// ===== LEITURA =====
function lerTudo() {
  var out = { ok: true };
  var mapa = { Contas: 'contas', Cartoes: 'cartoes', Compras_Cartao: 'compras', Recorrentes: 'recorrentes', Movimentos: 'movimentos', Emprestimos: 'emprestimos', Categorias: 'categorias', Config: 'config' };
  Object.keys(ABAS).forEach(function(nome) {
    out[mapa[nome]] = lerAba(nome);
  });
  return out;
}

function lerAba(nome) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nome);
  if (!sh) return [];
  var valores = sh.getDataRange().getValues();
  if (valores.length < 2) return [];
  var headers = valores[0];
  var linhas = [];
  for (var i = 1; i < valores.length; i++) {
    var row = valores[i];
    if (row.join('') === '') continue;
    var obj = {};
    for (var c = 0; c < headers.length; c++) {
      var v = row[c];
      obj[headers[c]] = (v instanceof Date) ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd/MM/yyyy') : v;
    }
    linhas.push(obj);
  }
  return linhas;
}

// ===== ESCRITA (CRUD genérico) =====
function abaOuErro(nome) {
  if (!ABAS[nome]) return null;
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nome);
}

function inserir(nome, registro) {
  var sh = abaOuErro(nome);
  if (!sh) return { ok: false, erro: 'aba_invalida' };
  var headers = ABAS[nome];
  if (!registro.ID) registro.ID = novoID(nome);
  var linha = headers.map(function(h) { return (registro[h] !== undefined && registro[h] !== null) ? registro[h] : ''; });
  sh.appendRow(linha);
  return { ok: true, id: registro.ID };
}

function editar(nome, registro) {
  var sh = abaOuErro(nome);
  if (!sh) return { ok: false, erro: 'aba_invalida' };
  var headers = ABAS[nome];
  var idx = encontrarLinha(sh, registro.ID);
  if (idx === -1) return { ok: false, erro: 'id_nao_encontrado' };
  var atual = sh.getRange(idx, 1, 1, headers.length).getValues()[0];
  var linha = headers.map(function(h, c) {
    return (registro[h] !== undefined) ? registro[h] : atual[c];
  });
  sh.getRange(idx, 1, 1, headers.length).setValues([linha]);
  return { ok: true, id: registro.ID };
}

function excluir(nome, id) {
  var sh = abaOuErro(nome);
  if (!sh) return { ok: false, erro: 'aba_invalida' };
  var idx = encontrarLinha(sh, id);
  if (idx === -1) return { ok: false, erro: 'id_nao_encontrado' };
  sh.deleteRow(idx);
  return { ok: true, id: id };
}

function encontrarLinha(sh, id) {
  var col = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
  for (var i = 1; i < col.length; i++) {
    if (String(col[i][0]) === String(id)) return i + 1;
  }
  return -1;
}

// Gera ID com prefixo por aba + timestamp curto + aleatório.
function novoID(nome) {
  var pre = { Contas: 'CC', Cartoes: 'CT', Compras_Cartao: 'CP', Recorrentes: 'RC', Movimentos: 'MV', Emprestimos: 'EM', Categorias: 'CG', Config: 'CF' }[nome] || 'XX';
  return pre + '-' + Date.now().toString(36).toUpperCase() + '-' + Math.floor(Math.random() * 1296).toString(36).toUpperCase();
}

// ===== UTIL =====
function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
