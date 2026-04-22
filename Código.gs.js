// ════════════════════════════════════════════════════════════════
//  Recrutamento Interno PRF — Google Apps Script (Código.gs)
// ════════════════════════════════════════════════════════════════

const SPREADSHEET_ID = '1OruPZa972HXynMYQz5jmSnlPZ1OlxVSSphpiUSeKXls';

// Cabeçalho da aba "respostas"
// Respostas brutas do BFI ficam em "bfi_resultados" (aba separada)
const HEADERS_RESPOSTAS = [
  'Data/Hora','E-mail','Nome','Matrícula','Cargo',
  'Unidade Oportunidade','Conhecimento da Unidade',
  'DDD','Telefone','Cônjuge',
  'Graduações','Pós-Graduações','Mestrados','Doutorados','Currículo SouGov',
  'Cônjuge Nome','Cônjuge Matrícula','Cônjuge E-mail','Tipo de União','Data da União',
  'Endereço Cônjuge 1','Endereço Cônjuge 2',
  'Lotação Cônjuge 1','Lotação Cônjuge 2',
  'URL Comprov. União','URL Comprob. Coabitação',
  'Extroversão','Amabilidade','Conscienciosidade','Estab. Emocional','Abertura',
  'Assinatura','Status',
  'ID Confirmação','Status Cônjuge','Data Confirmação Cônjuge',
  'PDF Respostas','PDF Termo','PDF Concordância Cônjuge'
];

// ─── Votação — constantes ─────────────────────────────────────────
const SHEET_VOTOS    = 'votos';
const HEADERS_VOTOS  = [
  'Data/Hora','Email Eleitor','Nome Eleitor',
  'Email Candidato','Nome Candidato','Matrícula Candidato',
  'Confiança','Lealdade','Amizade','Ego','Família',
  'Média','Comentário'
];

// ─── Web App ─────────────────────────────────────────────────────
function doGet(e) {
  var params = (e && e.parameter) ? e.parameter : {};
  var acao   = params.acao || '';

  // Confirmação / recusa de cônjuge via link de e-mail
  if (acao === 'confirmar' || acao === 'recusar') {
    return _processarRespostaConjuge(params.id || '', acao);
  }

  if (params.pagina === 'painel') {
    return HtmlService.createTemplateFromFile('Painel')
      .evaluate()
      .setTitle('Painel de Avaliação — PRF')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  if (params.pagina === 'votacao') {
    return HtmlService.createTemplateFromFile('Votacao')
      .evaluate()
      .setTitle('Votação Final — PRF')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  if (params.pagina === 'acompanhamento') {
    return HtmlService.createTemplateFromFile('Acompanhamento')
      .evaluate()
      .setTitle('Acompanhamento — PRF')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  var arquivo = params.pagina === 'formulario' ? 'Index' : 'Login';
  return HtmlService.createTemplateFromFile(arquivo)
    .evaluate()
    .setTitle('Recrutamento PRF')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ─── Validação de acesso (tela de login) ─────────────────────────
function validarAcesso(email) {
  if (!email || !email.toLowerCase().endsWith('@prf.gov.br')) {
    throw new Error('Acesso restrito a servidores com e-mail @prf.gov.br.');
  }
  // Retorna a URL do formulário com o parâmetro de página
  return ScriptApp.getService().getUrl() + '?pagina=formulario';
}

// ─── 2FA — Aba de registro ────────────────────────────────────────
const SHEET_2FA      = 'codigos_2fa';
const VALIDADE_2FA_H = 12; // horas

function _get2FASheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh   = ss.getSheetByName(SHEET_2FA);
  if (!sh) {
    sh = ss.insertSheet(SHEET_2FA);
    sh.appendRow(['Data/Hora Envio', 'E-mail', 'Código', 'Expira Em', 'Verificado Em', 'Status']);
    sh.setFrozenRows(1);
    sh.getRange('A1:F1').setFontWeight('bold').setBackground('#1d1a5b').setFontColor('#ffffff');
    sh.setColumnWidth(1, 160).setColumnWidth(2, 220).setColumnWidth(3, 80)
      .setColumnWidth(4, 160).setColumnWidth(5, 160).setColumnWidth(6, 90);
  }
  return sh;
}

// ─── Envio do código 2FA por e-mail (geração server-side) ─────────
function enviarEmail2FA(email) {
  if (!email || !email.toLowerCase().endsWith('@prf.gov.br')) {
    throw new Error('E-mail institucional @prf.gov.br obrigatório.');
  }

  // Gera código de 6 dígitos server-side
  const codigo  = String(Math.floor(100000 + Math.random() * 900000));
  const agora   = new Date();
  const expira  = new Date(agora.getTime() + VALIDADE_2FA_H * 60 * 60 * 1000);
  const tz      = Session.getScriptTimeZone();
  const fmtDt   = function(d) { return Utilities.formatDate(d, tz, "dd/MM/yyyy HH:mm:ss"); };

  // ── Salva (ou atualiza) na aba codigos_2fa ─────────────────────
  const sh   = _get2FASheet();
  const rows = sh.getDataRange().getValues();
  let targetRow = -1;

  // Procura entrada anterior do mesmo e-mail para sobrescrever
  for (let i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][1] || '').toLowerCase() === email.toLowerCase()) {
      targetRow = i + 1;
      break;
    }
  }

  const rowData = [fmtDt(agora), email, codigo, fmtDt(expira), '', 'Pendente'];
  if (targetRow > 0) {
    sh.getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);
  } else {
    sh.appendRow(rowData);
  }

  // ── Envia e-mail ───────────────────────────────────────────────
  const corpo = '<!DOCTYPE html><html lang="pt-BR"><body style="margin:0;padding:0;background:#f5f6f7;font-family:Arial,sans-serif;">'
    + '<div style="max-width:460px;margin:24px auto;">'
    + '<div style="background:#1d1a5b;padding:16px 24px;border-radius:10px 10px 0 0;">'
    + '<span style="color:#f2c200;font-weight:bold;font-size:22px;">PRF</span>'
    + '<span style="color:rgba(255,255,255,.85);font-size:14px;margin-left:14px;">Recrutamento Interno</span>'
    + '</div>'
    + '<div style="background:#fff;padding:28px 24px;border:1px solid #e0e0e0;">'
    + '<p style="font-size:15px;color:#1a2332;margin-bottom:6px;">Seu <strong>código de verificação</strong> é:</p>'
    + '<div style="background:#f5f6f7;border:2px dashed #1d1a5b;border-radius:8px;padding:20px;text-align:center;margin:18px 0;">'
    + '<span style="font-size:44px;font-weight:bold;color:#1d1a5b;letter-spacing:16px;">' + codigo + '</span>'
    + '</div>'
    + '<p style="font-size:13px;color:#555;margin-bottom:4px;">&#x23F0; Válido por <strong>' + VALIDADE_2FA_H + ' horas</strong> — expira em <strong>' + fmtDt(expira) + '</strong>.</p>'
    + '<p style="font-size:13px;color:#555;margin:0;">Se você não solicitou este código, ignore este e-mail.</p>'
    + '</div>'
    + '<div style="background:#f2c200;padding:8px 24px;border-radius:0 0 10px 10px;font-size:11px;color:#1d1a5b;">'
    + 'Polícia Rodoviária Federal | DIPROM/DGP &mdash; mensagem automática, não responda.'
    + '</div></div></body></html>';

  MailApp.sendEmail({ to: email, subject: '[PRF Recrutamento] Seu código: ' + codigo, htmlBody: corpo, name: 'Recrutamento PRF' });
}

// ═══════════════════════════════════════════════════════════════════
//  LOGIN VIA GOOGLE WORKSPACE (substitui fluxo 2FA)
// ═══════════════════════════════════════════════════════════════════
//
//  Para reativar o 2FA por e-mail:
//   1. Substitua a chamada a iniciarSessao() no Login.html pela
//      chamada a enviarEmail2FA() (descomente o bloco [2FA_OFF]).
//   2. Certifique-se de que o DKIM do domínio prf.gov.br está ativo:
//      Admin Console → Apps → Google Workspace → Gmail →
//      Autenticar e-mail → Iniciar autenticação.
//
// ─── Detecta e-mails de área/setor (ex: sgp.go@prf.gov.br) ───────
function _isEmailDeArea(email) {
  const local  = String(email || '').toLowerCase().split('@')[0];
  const partes = local.split('.');

  // Siglas de unidades organizacionais conhecidas da PRF
  const UORGS = new Set([
    'dgp','diprom','dicop','dicat','drgp','dare','dasp',
    'cgp','cgcsp','cgti','cgrh','cglog','cgfin','cgad',
    'direx','dint','dor','dop','daf','dpp','dprf',
    'gate','niop','niac','nucrim','nefaz','nesp','nug','nud','nup',
    'sgp','srf','srp','sop','seop','set','stt','sri','sti',
    'naf','cci','gab','sec','adm','fin','log','saf','sas','sal','sap'
  ]);

  // Códigos de UF brasileiras
  const UFS = new Set([
    'ac','al','am','ap','ba','ce','df','es','go',
    'ma','mg','ms','mt','pa','pb','pe','pi','pr',
    'rj','rn','ro','rr','rs','sc','se','sp','to'
  ]);

  // Bloqueia se qualquer segmento é uma sigla de uorg conhecida
  for (const p of partes) {
    if (UORGS.has(p)) return true;
  }

  // Bloqueia padrão sigla_curta.uf — ex: sgp.go, naf.sp
  if (partes.length === 2 && partes[0].length <= 5 && UFS.has(partes[1])) return true;

  return false;
}

// ─── Inicia sessão usando a conta Google Workspace logada ──────────
function iniciarSessao() {
  const email = Session.getActiveUser().getEmail();
  if (!email) {
    throw new Error('Não foi possível identificar o usuário. Acesse pelo link institucional com sua conta @prf.gov.br.');
  }
  if (!email.toLowerCase().endsWith('@prf.gov.br')) {
    throw new Error('Acesso restrito a servidores com e-mail @prf.gov.br.');
  }
  if (_isEmailDeArea(email)) {
    throw new Error(
      'O e-mail "' + email + '" parece ser de uma área ou setor (ex: sgp.go@prf.gov.br). ' +
      'Acesse com seu e-mail pessoal institucional.'
    );
  }

  const baseUrl    = ScriptApp.getService().getUrl();
  const painelInfo = _checarPerfilPainel(email);

  let votacaoUrl = null;
  try {
    const cRows = _getCredenciaisSheet().getDataRange().getValues();
    for (let j = cRows.length - 1; j >= 1; j--) {
      const ce = String(cRows[j][0] || '').toLowerCase();
      const cp = String(cRows[j][2] || '').toUpperCase();
      const cs = String(cRows[j][5] || '').toLowerCase();
      if (ce === email.toLowerCase() && cs === 'ativo'
          && (cp === 'VOTADOR' || cp === 'ADMINISTRADOR')) {
        votacaoUrl = baseUrl + '?pagina=votacao';
        break;
      }
    }
  } catch(e) {}

  let acompanhamentoUrl = null;
  try {
    const shResp = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('respostas');
    if (shResp && shResp.getLastRow() > 1) {
      const rows = shResp.getRange(1, 1, shResp.getLastRow(), 2).getValues().slice(1);
      for (let k = rows.length - 1; k >= 0; k--) {
        if (String(rows[k][1] || '').toLowerCase() === email.toLowerCase()) {
          acompanhamentoUrl = baseUrl + '?pagina=acompanhamento';
          break;
        }
      }
    }
  } catch(e) {}

  return {
    email:             email,
    formUrl:           baseUrl + '?pagina=formulario',
    painelUrl:         painelInfo ? (baseUrl + '?pagina=painel') : null,
    votacaoUrl:        votacaoUrl,
    acompanhamentoUrl: acompanhamentoUrl,
    perfil:            painelInfo ? painelInfo.perfil : null,
    nome:              painelInfo ? painelInfo.nome   : email.split('@')[0]
  };
}

// ─── Checa se o e-mail tem perfil no painel (sem lançar erro) ─────
function _checarPerfilPainel(email) {
  try {
    const rows = _getCredenciaisSheet().getDataRange().getValues();
    for (let i = rows.length - 1; i >= 1; i--) {
      if (String(rows[i][0]).toLowerCase() === email.toLowerCase()) {
        if (String(rows[i][5]).toLowerCase() !== 'ativo') return null;
        return { perfil: String(rows[i][2]), nome: String(rows[i][1] || email) };
      }
    }
  } catch (e) { /* aba credenciais pode não existir ainda */ }
  return null;
}

// ─── Verificação do código 2FA (server-side) ──────────────────────
function verificarCodigo2FA(email, codigoDigitado) {
  if (!email || !codigoDigitado) throw new Error('Dados inválidos.');

  const sh   = _get2FASheet();
  const rows = sh.getDataRange().getValues();
  const tz   = Session.getScriptTimeZone();
  const fmtDt = function(d) { return Utilities.formatDate(d, tz, "dd/MM/yyyy HH:mm:ss"); };

  // Busca a linha mais recente do e-mail
  for (let i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][1] || '').toLowerCase() !== email.toLowerCase()) continue;

    const codigoSalvo = String(rows[i][2] || '');
    const expiraStr   = String(rows[i][3] || '');
    const status      = String(rows[i][5] || '');

    if (status === 'Usado') {
      throw new Error('Este código já foi utilizado. Solicite um novo.');
    }

    // Verifica expiração (coluna D guarda data/hora como string "dd/MM/yyyy HH:mm:ss")
    const partes  = expiraStr.split(/[/ :]/);
    // partes: [dd, MM, yyyy, HH, mm, ss]
    const expira  = new Date(partes[2], partes[1]-1, partes[0], partes[3], partes[4], partes[5]);
    if (new Date() > expira) {
      // Marca como expirado na planilha
      sh.getRange(i + 1, 6).setValue('Expirado');
      throw new Error('Código expirado. Solicite um novo código de verificação.');
    }

    if (codigoDigitado.trim() !== codigoSalvo) {
      throw new Error('Código incorreto. Verifique e tente novamente.');
    }

    // Código válido — registra uso
    const agora = new Date();
    sh.getRange(i + 1, 5).setValue(fmtDt(agora));
    sh.getRange(i + 1, 6).setValue('Usado');

    const baseUrl    = ScriptApp.getService().getUrl();
    const formUrl    = baseUrl + '?pagina=formulario';
    const painelInfo = _checarPerfilPainel(email);

    // Verifica perfil de votador (VOTADOR ou ADMINISTRADOR na aba credenciais)
    let votacaoUrl = null;
    try {
      const cRows = _getCredenciaisSheet().getDataRange().getValues();
      for (let j = cRows.length - 1; j >= 1; j--) {
        const ce = String(cRows[j][0] || '').toLowerCase();
        const cp = String(cRows[j][2] || '').toUpperCase();
        const cs = String(cRows[j][5] || '').toLowerCase();
        if (ce === email.toLowerCase() && cs === 'ativo'
            && (cp === 'VOTADOR' || cp === 'ADMINISTRADOR')) {
          votacaoUrl = baseUrl + '?pagina=votacao';
          break;
        }
      }
    } catch(e) { /* aba pode não existir ainda */ }

    // Verifica se o candidato tem inscrição para exibir link de acompanhamento
    let acompanhamentoUrl = null;
    try {
      const shResp = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('respostas');
      if (shResp && shResp.getLastRow() > 1) {
        const respRows = shResp.getRange(1, 1, shResp.getLastRow(), 2).getValues().slice(1);
        for (let k = respRows.length - 1; k >= 0; k--) {
          if (String(respRows[k][1] || '').toLowerCase() === email.toLowerCase()) {
            acompanhamentoUrl = baseUrl + '?pagina=acompanhamento';
            break;
          }
        }
      }
    } catch(e) { /* aba pode não existir ainda */ }

    return {
      formUrl:           formUrl,
      painelUrl:         painelInfo  ? (baseUrl + '?pagina=painel') : null,
      votacaoUrl:        votacaoUrl,
      acompanhamentoUrl: acompanhamentoUrl,
      perfil:            painelInfo  ? painelInfo.perfil : null,
      nome:              painelInfo  ? painelInfo.nome   : (votacaoUrl ? email.split('@')[0] : null)
    };
  }

  throw new Error('Nenhum código encontrado para este e-mail. Solicite um novo.');
}

// ─── Inicialização combinada (listas + inscrição existente) ───────
// Abre o SpreadsheetApp uma única vez e repassa para as subfunções,
// evitando múltiplos roundtrips à API do Sheets.
function inicializarFormulario() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  return {
    listas:    getListasFormacao(ss),
    inscricao: verificarInscricaoExistente(ss)
  };
}

// ─── Verifica inscrição anterior pelo e-mail logado ───────────────
function verificarInscricaoExistente(ss) {
  const email = Session.getActiveUser().getEmail();
  if (!email) return null;

  if (!ss) ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('respostas');
  if (!sheet || sheet.getLastRow() < 2) return null;

  // Lê apenas as colunas necessárias (A–Z = colunas 1–26).
  // Evita propositalmente a coluna 32 (assinaturaJson) que pode conter
  // dezenas de KB de base64 e causaria lentidão/timeout desnecessário.
  const lastRow = sheet.getLastRow();
  const data    = sheet.getRange(1, 1, lastRow, 27).getValues();

  // Busca de baixo para cima para pegar a entrada mais recente
  let row = null;
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][1] || '').toLowerCase() === email.toLowerCase()) {
      row = data[i];
      break;
    }
  }
  if (!row) return null;

  const split = function (val) {
    return val ? String(val).split(' | ').filter(Boolean) : [];
  };

  // Buscar respostas brutas do BFI na aba separada
  let bfiResponses = [];
  const sheetBFI = ss.getSheetByName('bfi_resultados');
  if (sheetBFI && sheetBFI.getLastRow() > 1) {
    const bfiData = sheetBFI.getDataRange().getValues();
    for (let j = bfiData.length - 1; j >= 1; j--) {
      if (String(bfiData[j][1] || '').toLowerCase() === email.toLowerCase()) {
        const brutas = String(bfiData[j][11] || ''); // coluna "Respostas Brutas"
        bfiResponses = brutas ? brutas.split(',').map(Number) : [];
        break;
      }
    }
  }

  return {
    nome:                  row[2]  || '',
    matricula:             row[3]  || '',
    cargo:                 row[4]  || '',
    unidadeOportunidade:   row[5]  || '',
    conhecimentoUnidade:   row[6]  || '',
    ddd:                   row[7]  || '',
    telefone:              row[8]  || '',
    conjuge:               String(row[9] || '').toLowerCase() === 'sim' ? 'sim' : 'nao',
    graduacao:             split(row[10]),
    pos:                   split(row[11]),
    mestrado:              split(row[12]),
    doutorado:             split(row[13]),
    sougovUrl:             row[14] || '',
    conjugeNome:           row[15] || '',
    conjugeMatricula:      row[16] || '',
    conjugeEmail:          row[17] || '',
    tipoUniao:             row[18] || '',
    dataUniao:             row[19] || '',
    enderecoConjuge1:      row[20] || '',
    enderecoConjuge2:      row[21] || '',
    lotacaoConjuge1:       row[22] || '',
    lotacaoConjuge2:       row[23] || '',
    urlComprovUniao:       row[24] || '',
    urlComprovCoab:        row[25] || '',
    bfiResponses:          bfiResponses
  };
}

// ─── Salvar rascunho parcial do BFI ──────────────────────────────
function salvarRascunhoBFI(bfiParcial) {
  const email = Session.getActiveUser().getEmail();
  if (!email || !email.toLowerCase().endsWith('@prf.gov.br')) {
    throw new Error('Acesso restrito a e-mails @prf.gov.br.');
  }

  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet   = ss.getSheetByName('rascunhos_bfi');
  if (!sheet) {
    sheet = ss.insertSheet('rascunhos_bfi');
    sheet.appendRow(['Data/Hora', 'E-mail', 'Respostas (rascunho)', 'Respondidas']);
  }

  const agora       = new Date();
  const respondidas = (bfiParcial || []).filter(function (v) { return v > 0; }).length;
  const rowData     = [agora, email, (bfiParcial || []).join(','), respondidas + '/44'];

  // Sobrescrever rascunho existente ou adicionar novo
  const existing = sheet.getDataRange().getValues();
  let targetRow  = -1;
  for (let i = existing.length - 1; i >= 1; i--) {
    if (String(existing[i][1] || '').toLowerCase() === email.toLowerCase()) {
      targetRow = i + 1; // 1-indexed
      break;
    }
  }

  if (targetRow > 0) {
    sheet.getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
  }

  return 'Rascunho salvo: ' + respondidas + '/44 questões respondidas.';
}

// ─── Listas para autocomplete ─────────────────────────────────────
function getListasFormacao(ss) {
  // Verifica cache primeiro — listas raramente mudam, TTL de 6 horas
  const cache       = CacheService.getScriptCache();
  const cacheKey    = 'listas_formacao_v1';
  const cachedJson  = cache.get(cacheKey);
  if (cachedJson) {
    try { return JSON.parse(cachedJson); } catch (e) { /* cache corrompido, recalcula */ }
  }

  if (!ss) ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  const sheetGrad   = ss.getSheetByName('graduacao');
  let graduacoes    = [];
  let posGraduacoes = [];
  if (sheetGrad) {
    const data    = sheetGrad.getDataRange().getValues();
    graduacoes    = data.slice(1).map(function (r) { return r[0]; }).filter(String);
    posGraduacoes = data.slice(1).map(function (r) { return r[1]; }).filter(String);
  }

  // UORGs — coluna A da aba "uorgs" (para lotação do cônjuge)
  const sheetUorgs = ss.getSheetByName('uorgs');
  let uorgs = [];
  if (sheetUorgs && sheetUorgs.getLastRow() > 1) {
    uorgs = sheetUorgs
      .getRange(2, 1, sheetUorgs.getLastRow() - 1, 1)
      .getValues()
      .flat()
      .filter(String);
  }

  // Unidades de oportunidade — coluna A da aba "unidades_oportunidade"
  const sheetOport = ss.getSheetByName('unidades_oportunidade');
  let unidadesOportunidade = [];
  if (sheetOport && sheetOport.getLastRow() > 1) {
    unidadesOportunidade = sheetOport
      .getRange(2, 1, sheetOport.getLastRow() - 1, 1)
      .getValues()
      .flat()
      .filter(String);
  }

  const resultado = {
    graduacoes:           [...new Set(graduacoes)],
    posGraduacoes:        [...new Set(posGraduacoes)],
    uorgs:                [...new Set(uorgs)],
    unidadesOportunidade: [...new Set(unidadesOportunidade)]
  };

  // Salva no cache por 6 horas (21600 segundos)
  try { cache.put(cacheKey, JSON.stringify(resultado), 21600); } catch (e) { /* ignora se JSON for grande demais */ }

  return resultado;
}

// ─── Upload de arquivo genérico ───────────────────────────────────
function salvarArquivoGenerico(base64Str, nomeArquivo, nomePasta) {
  const email = Session.getActiveUser().getEmail();
  if (!email || !email.toLowerCase().endsWith('@prf.gov.br')) {
    throw new Error('Acesso restrito a e-mails @prf.gov.br.');
  }

  const decoded = Utilities.base64Decode(base64Str);

  let mime = 'application/octet-stream';
  const ext = nomeArquivo.toLowerCase().split('.').pop();
  if (ext === 'pdf')                    mime = 'application/pdf';
  else if (ext === 'jpg' || ext === 'jpeg') mime = 'image/jpeg';
  else if (ext === 'png')               mime = 'image/png';

  const blob    = Utilities.newBlob(decoded, mime, nomeArquivo);
  const pasta   = getOuCriarPasta(nomePasta);
  const arquivo = pasta.createFile(blob);
  arquivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return arquivo.getUrl();
}

function salvarArquivoSouGov(base64Str, nomeArquivo) {
  return salvarArquivoGenerico(base64Str, nomeArquivo, 'Curriculos SouGov - Recrutamento PRF');
}

function getOuCriarPasta(nomePasta) {
  const folders = DriveApp.getFoldersByName(nomePasta);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(nomePasta);
}

// ─── Salvar progresso parcial (chamado a cada avanço de tela) ────
function salvarProgressoInscricao(dados) {
  const email = Session.getActiveUser().getEmail();
  if (!email || !email.toLowerCase().endsWith('@prf.gov.br')) {
    throw new Error('Acesso restrito a e-mails @prf.gov.br.');
  }

  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const agora = new Date();
  const fmt   = function(v) { return Array.isArray(v) ? v.filter(String).join(' | ') : (v || ''); };

  let sheetResp = ss.getSheetByName('respostas');
  if (!sheetResp) {
    sheetResp = ss.insertSheet('respostas');
    sheetResp.appendRow(HEADERS_RESPOSTAS);
    sheetResp.setFrozenRows(1);
  }

  const conjugeAtivo = dados.conjuge === 'sim' || dados.conjuge === true;

  // Calcula BFI se as 44 respostas já foram preenchidas
  let bfiResult = null;
  const bfiRespostas = Array.isArray(dados.bfiResponses) && dados.bfiResponses.length === 44
    ? dados.bfiResponses : [];
  if (bfiRespostas.length === 44) bfiResult = calcularBFI(bfiRespostas);

  const rowData = [
    agora,
    email,
    (dados.nome || '').toUpperCase(),
    dados.matricula            || '',
    dados.cargo                || '',
    dados.unidadeOportunidade  || '',
    dados.conhecimentoUnidade  || '',
    dados.ddd                  || '',
    dados.telefone             || '',
    conjugeAtivo ? 'Sim' : (dados.conjuge === 'nao' ? 'Não' : ''),
    fmt(dados.graduacao),
    fmt(dados.pos),
    fmt(dados.mestrado),
    fmt(dados.doutorado),
    dados.sougovUrl            || '',
    conjugeAtivo ? (dados.conjugeNome        || '') : '',
    conjugeAtivo ? (dados.conjugeMatricula   || '') : '',
    conjugeAtivo ? (dados.conjugeEmail       || '') : '',
    conjugeAtivo ? (dados.tipoUniao          || '') : '',
    conjugeAtivo ? (dados.dataUniao          || '') : '',
    conjugeAtivo ? (dados.enderecoConjuge1   || '') : '',
    conjugeAtivo ? (dados.enderecoConjuge2   || '') : '',
    conjugeAtivo ? (dados.lotacaoConjuge1    || '') : '',
    conjugeAtivo ? (dados.lotacaoConjuge2    || '') : '',
    conjugeAtivo ? (dados.urlComprovUniao    || '') : '',
    conjugeAtivo ? (dados.urlComprovCoab     || '') : '',
    bfiResult ? bfiResult.ext.toFixed(2)   : '',
    bfiResult ? bfiResult.amab.toFixed(2)  : '',
    bfiResult ? bfiResult.cons.toFixed(2)  : '',
    bfiResult ? bfiResult.estab.toFixed(2) : '',
    bfiResult ? bfiResult.abert.toFixed(2) : '',
    '',              // Assinatura — só gerada na finalização
    'Em preenchimento',
    '',              // ID Confirmação
    conjugeAtivo ? 'Pendente' : '',
    '',              // Data Confirmação Cônjuge
    '',              // PDF Respostas
    '',              // PDF Termo
    ''               // PDF Concordância Cônjuge
  ];

  // Sobrescreve linha existente, preservando campos que não devem ser perdidos
  const allRows = sheetResp.getDataRange().getValues();
  for (let i = allRows.length - 1; i >= 1; i--) {
    if (String(allRows[i][1] || '').toLowerCase() !== email.toLowerCase()) continue;

    const prev = allRows[i];
    // Preserva ID de confirmação (gerado na inscrição final)
    if (prev[33]) rowData[33] = prev[33];
    // Preserva status de cônjuge se já foi confirmado/recusado
    if (prev[34] && String(prev[34]).toLowerCase() !== 'pendente') {
      rowData[34] = prev[34];
      rowData[35] = prev[35]; // Data Confirmação
      rowData[38] = prev[38]; // PDF Concordância
    }
    // Preserva PDFs já gerados
    if (prev[36]) rowData[36] = prev[36];
    if (prev[37]) rowData[37] = prev[37];
    // Preserva status final se inscrição já foi concluída
    if (String(prev[32]).toLowerCase() === 'inscrito') rowData[32] = 'Inscrito';

    sheetResp.getRange(i + 1, 1, 1, rowData.length).setValues([rowData]);
    return 'atualizado';
  }

  sheetResp.appendRow(rowData);
  return 'criado';
}

// ─── Processamento principal ──────────────────────────────────────
function processarInscricao(dados) {
  const email = Session.getActiveUser().getEmail();
  if (!email || !email.toLowerCase().endsWith('@prf.gov.br')) {
    throw new Error('Acesso restrito a e-mails @prf.gov.br.');
  }

  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const agora = new Date();
  const fmt   = function (v) { return Array.isArray(v) ? v.filter(String).join(' | ') : (v || ''); };

  // ── 1. Aba "respostas" ───────────────────────────────────────
  let sheetResp = ss.getSheetByName('respostas');
  if (!sheetResp) {
    sheetResp = ss.insertSheet('respostas');
    sheetResp.appendRow(HEADERS_RESPOSTAS);
    sheetResp.setFrozenRows(1);
  } else {
    // Garantir que o cabeçalho tenha todas as colunas novas
    const cabecalho = sheetResp.getRange(1, 1, 1, HEADERS_RESPOSTAS.length);
    if (sheetResp.getLastColumn() < HEADERS_RESPOSTAS.length) {
      cabecalho.setValues([HEADERS_RESPOSTAS]);
    }
  }

  const conjugeAtivo = dados.conjuge === 'sim' || dados.conjuge === true;
  const confirmId    = conjugeAtivo ? Utilities.getUuid() : '';

  // BFI
  let bfiResult = null;
  const bfiRespostas = Array.isArray(dados.bfiResponses) && dados.bfiResponses.length === 44
    ? dados.bfiResponses : [];
  if (bfiRespostas.length === 44) {
    bfiResult = calcularBFI(bfiRespostas);
  }

  // Assinatura
  let assinaturaJson = '';
  if (dados.assinatura && typeof dados.assinatura === 'object') {
    assinaturaJson = JSON.stringify(dados.assinatura);
  }

  // ── Montar linha (sem PDFs ainda) ────────────────────────────
  dados._email = email; // disponibiliza e-mail no contexto dos geradores de PDF
  const nomeSanitizado = (dados.nome || 'Candidato').replace(/[^a-zA-Z0-9À-ú ]/g, '').trim();

  const rowData = [
    agora,
    email,
    (dados.nome || '').toUpperCase(),
    dados.matricula            || '',
    dados.cargo                || '',
    dados.unidadeOportunidade  || '',
    dados.conhecimentoUnidade  || '',
    dados.ddd                  || '',
    dados.telefone             || '',
    conjugeAtivo ? 'Sim' : 'Não',
    fmt(dados.graduacao),
    fmt(dados.pos),
    fmt(dados.mestrado),
    fmt(dados.doutorado),
    dados.sougovUrl            || '',
    // Cônjuge
    conjugeAtivo ? (dados.conjugeNome        || '') : '',
    conjugeAtivo ? (dados.conjugeMatricula   || '') : '',
    conjugeAtivo ? (dados.conjugeEmail       || '') : '',
    conjugeAtivo ? (dados.tipoUniao          || '') : '',
    conjugeAtivo ? (dados.dataUniao          || '') : '',
    conjugeAtivo ? (dados.enderecoConjuge1   || '') : '',
    conjugeAtivo ? (dados.enderecoConjuge2   || '') : '',
    conjugeAtivo ? (dados.lotacaoConjuge1    || '') : '',
    conjugeAtivo ? (dados.lotacaoConjuge2    || '') : '',
    conjugeAtivo ? (dados.urlComprovUniao    || '') : '',
    conjugeAtivo ? (dados.urlComprovCoab     || '') : '',
    // BFI — apenas escores calculados (respostas brutas ficam em bfi_resultados)
    bfiResult ? bfiResult.ext.toFixed(2)   : '',
    bfiResult ? bfiResult.amab.toFixed(2)  : '',
    bfiResult ? bfiResult.cons.toFixed(2)  : '',
    bfiResult ? bfiResult.estab.toFixed(2) : '',
    bfiResult ? bfiResult.abert.toFixed(2) : '',
    assinaturaJson,
    'Inscrito',
    // Confirmação cônjuge
    confirmId,
    conjugeAtivo ? 'Pendente' : '',
    '',           // Data Confirmação Cônjuge (preenchida quando cônjuge responde)
    '',           // PDF Respostas  — preenchido logo abaixo
    '',           // PDF Termo      — preenchido logo abaixo
    ''            // PDF Concordância Cônjuge (gerado quando cônjuge confirma)
  ];

  // ── Salvar linha PRIMEIRO (garante a inscrição mesmo se PDFs falharem) ──
  const allRows = sheetResp.getDataRange().getValues();
  let savedRowIdx = -1;
  for (let i = allRows.length - 1; i >= 1; i--) {
    if (String(allRows[i][1] || '').toLowerCase() === email.toLowerCase()) {
      savedRowIdx = i + 1;
      break;
    }
  }

  if (savedRowIdx > 0) {
    sheetResp.getRange(savedRowIdx, 1, 1, rowData.length).setValues([rowData]);
  } else {
    sheetResp.appendRow(rowData);
    savedRowIdx = sheetResp.getLastRow();
  }

  // ── PDFs — gerados após salvar a linha ───────────────────────
  let pdfRespostasUrl = '';
  let pdfTermoUrl     = '';
  let pdfRespostasBlob = null;
  let pdfTermoBlob     = null;
  let pdfErro          = '';

  try {
    const pdfRespostasResult = _gerarESalvarPdf(
      _htmlPdfRespostas(dados, bfiResult, agora),
      'Ficha_' + nomeSanitizado
    );
    pdfRespostasUrl  = pdfRespostasResult.url;
    pdfRespostasBlob = pdfRespostasResult.blob;

    const pdfTermoResult = _gerarESalvarPdf(
      _htmlPdfTermo(dados),
      'Termo_' + nomeSanitizado
    );
    pdfTermoUrl  = pdfTermoResult.url;
    pdfTermoBlob = pdfTermoResult.blob;
  } catch (pdfErr) {
    pdfErro = pdfErr.message;
    Logger.log('[processarInscricao] Falha nos PDFs: ' + pdfErro);
  }

  // Retroalimentar as colunas de PDF na linha já salva (índices 36 e 37 → colunas 37 e 38)
  if (pdfRespostasUrl || pdfTermoUrl) {
    sheetResp.getRange(savedRowIdx, 37, 1, 2).setValues([[pdfRespostasUrl, pdfTermoUrl]]);
  }

  // ── 2. Aba "bfi_resultados" (separada para análise) ──────────
  if (bfiResult) {
    let sheetBFI = ss.getSheetByName('bfi_resultados');
    if (!sheetBFI) {
      sheetBFI = ss.insertSheet('bfi_resultados');
      sheetBFI.appendRow([
        'Data/Hora','E-mail','Nome','Matrícula','Cargo','Área Desejada',
        'Extroversão','Amabilidade','Conscienciosidade',
        'Estab. Emocional','Abertura','Respostas Brutas'
      ]);
      sheetBFI.setFrozenRows(1);
    }

    const bfiRow = [
      agora, email,
      (dados.nome || '').toUpperCase(),
      dados.matricula    || '',
      dados.cargo        || '',
      dados.areaDesejada || '',
      bfiResult.ext.toFixed(2),
      bfiResult.amab.toFixed(2),
      bfiResult.cons.toFixed(2),
      bfiResult.estab.toFixed(2),
      bfiResult.abert.toFixed(2),
      bfiRespostas.join(',')
    ];

    const bfiRows = sheetBFI.getDataRange().getValues();
    let bfiExisting = -1;
    for (let i = bfiRows.length - 1; i >= 1; i--) {
      if (String(bfiRows[i][1] || '').toLowerCase() === email.toLowerCase()) {
        bfiExisting = i + 1;
        break;
      }
    }

    if (bfiExisting > 0) {
      sheetBFI.getRange(bfiExisting, 1, 1, bfiRow.length).setValues([bfiRow]);
    } else {
      sheetBFI.appendRow(bfiRow);
    }
  }

  // ── 3. E-mail de confirmação ao candidato ────────────────────
  const anexos = [pdfRespostasBlob, pdfTermoBlob].filter(Boolean);
  enviarEmailConfirmacao(email, dados, bfiResult, agora, anexos);

  // ── 4. E-mail de confirmação ao cônjuge ──────────────────────
  if (conjugeAtivo && dados.conjugeEmail) {
    _enviarEmailConjuge(dados, confirmId);
  }

  return pdfErro ? ('ok|aviso:PDF não gerado — ' + pdfErro) : 'ok';
}

// ─── Cálculo BFI-44 ───────────────────────────────────────────────
function calcularBFI(responses) {
  const invertidos = [2,6,8,9,12,18,23,24,27,31,34,35,37,41,43];
  const adj = responses.map(function (v, i) {
    return invertidos.includes(i + 1) ? (6 - Number(v)) : Number(v);
  });
  const s = adj;
  const ext   = (s[0]+s[5]+s[10]+s[15]+s[20]+s[25]+s[30]+s[35]) / 8;
  const amab  = (s[1]+s[6]+s[11]+s[16]+s[21]+s[26]+s[31]+s[36]+s[41]) / 9;
  const cons  = (s[2]+s[7]+s[12]+s[17]+s[22]+s[27]+s[32]+s[37]+s[42]) / 9;
  const estab = (s[3]+s[8]+s[13]+s[18]+s[23]+s[28]+s[33]+s[38]) / 8;
  const abert = (s[4]+s[9]+s[14]+s[19]+s[24]+s[29]+s[34]+s[39]+s[40]+s[43]) / 10;
  return { ext, amab, cons, estab, abert };
}

// ─── Interpretação ────────────────────────────────────────────────
function interpretar(v) {
  if (v >= 3.6) return { nivel:'Alto',  emoji:'⬆️', cor:'#1b5e20' };
  if (v >= 2.5) return { nivel:'Médio', emoji:'➡️', cor:'#e65100' };
  return              { nivel:'Baixo', emoji:'⬇️', cor:'#b71c1c' };
}

const DIMS_INFO = {
  ext:   { nome:'Extroversão',            cor:'#1565C0', alto:'Comunicativo(a), assertivo(a) e cheio(a) de energia. Prospera em ambientes sociais.',                    medio:'Equilibrado(a) entre sociabilidade e introspecção. Adapta-se bem a diferentes contextos.',      baixo:'Reservado(a) e introspectivo(a). Prefere ambientes calmos e foco individual.' },
  amab:  { nome:'Amabilidade',            cor:'#2e7d32', alto:'Cooperativo(a), gentil e empático(a). Valoriza o trabalho em equipe e o bem-estar coletivo.',            medio:'Equilibrado(a) entre cooperação e assertividade, adaptando-se conforme a situação.',           baixo:'Direto(a) e independente. Pode ser crítico(a) e focado(a) em resultados objetivos.' },
  cons:  { nome:'Conscienciosidade',      cor:'#e65100', alto:'Organizado(a), disciplinado(a) e confiável. Planeja e cumpre responsabilidades com rigor.',              medio:'Razoavelmente organizado(a). Equilibra estrutura e flexibilidade conforme a demanda.',          baixo:'Flexível e espontâneo(a). Prefere adaptar-se às circunstâncias em vez de planos rígidos.' },
  estab: { nome:'Estabilidade Emocional', cor:'#6a1b9a', alto:'Calmo(a), equilibrado(a) e resistente ao estresse. Lida bem com pressão e adversidades.',               medio:'Relativamente estável. Pode sentir alguma tensão em situações de alta pressão.',               baixo:'Mais sensível emocionalmente. Pode experimentar ansiedade com maior frequência.' },
  abert: { nome:'Abertura à Experiência', cor:'#00695c', alto:'Criativo(a), curioso(a) e aberto(a) ao novo. Aprecia aprender e explorar perspectivas diferentes.',     medio:'Equilibrado(a) entre tradição e inovação. Aberto(a) a novas ideias bem fundamentadas.',        baixo:'Prático(a) e convencional. Prefere abordagens comprovadas e ambientes previsíveis.' }
};

// ─── Helpers de e-mail ────────────────────────────────────────────
function _secao(titulo, cor) {
  cor = cor || '#FFC400';
  return '<h2 style="color:#002244;font-size:16px;margin:28px 0 10px;border-left:4px solid ' + cor + ';padding-left:12px;">' + titulo + '</h2>';
}
function _linha(label, valor, negrito) {
  if (!valor) return '';
  return '<tr>'
    + '<td style="color:#888;width:180px;padding:4px 0;vertical-align:top;font-size:13px;">' + label + '</td>'
    + '<td style="font-size:13px;color:#1a2332;padding:4px 0;">' + (negrito ? '<strong>' + valor + '</strong>' : valor) + '</td>'
    + '</tr>';
}
function _linksArquivos(urlsStr, rotulo) {
  if (!urlsStr) return '';
  var urls = String(urlsStr).split('|').filter(Boolean);
  if (!urls.length) return '';
  var links = urls.map(function(u, i) {
    var n = urls.length > 1 ? ' ' + (i + 1) : '';
    return '<a href="' + u + '" style="color:#1565C0;margin-right:10px;">📎 ' + rotulo + n + '</a>';
  }).join('');
  return '<tr><td style="color:#888;width:180px;padding:4px 0;font-size:13px;vertical-align:top;">' + rotulo + (urls.length > 1 ? 's' : '') + ':</td>'
    + '<td style="font-size:13px;padding:4px 0;">' + links + '</td></tr>';
}

// ─── E-mail HTML ──────────────────────────────────────────────────
function enviarEmailConfirmacao(email, dados, bfiResult, data, anexos) {
  const tz        = Session.getScriptTimeZone();
  const dtStr     = Utilities.formatDate(data, tz, "dd/MM/yyyy 'às' HH:mm");
  const conjAtivo = dados.conjuge === 'sim' || dados.conjuge === true;

  // ── Bloco BFI completo ─────────────────────────────────────────
  let bfiHtml = '';
  if (bfiResult) {
    // Dimensão dominante
    const dims = [
      { k:'ext', v:bfiResult.ext },{ k:'amab', v:bfiResult.amab },
      { k:'cons', v:bfiResult.cons },{ k:'estab', v:bfiResult.estab },
      { k:'abert', v:bfiResult.abert }
    ];
    const dominante = dims.reduce(function(a,b){ return b.v > a.v ? b : a; });

    const rows = ['ext','amab','cons','estab','abert'].map(function (k) {
      const info   = DIMS_INFO[k];
      const val    = bfiResult[k];
      const interp = interpretar(val);
      const pct    = Math.round((val / 5) * 100);
      const desc   = val >= 3.6 ? info.alto : (val >= 2.5 ? info.medio : info.baixo);
      const isDom  = k === dominante.k;
      return '<tr style="' + (isDom ? 'background:#fffde7;' : '') + '">'
        + '<td style="padding:12px 16px;border-bottom:1px solid #eee;">'
        + '<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:5px;">'
        + '<strong style="color:' + info.cor + ';font-size:13px;">' + info.nome + (isDom ? ' ⭐' : '') + '</strong>'
        + '<span style="background:' + interp.cor + ';color:#fff;padding:2px 9px;border-radius:10px;font-size:11px;font-weight:700;">'
        + interp.nivel + ' — ' + val.toFixed(2) + '</span></div>'
        + '<div style="background:#e8e8e8;border-radius:4px;height:6px;margin-bottom:7px;overflow:hidden;">'
        + '<div style="background:' + info.cor + ';height:6px;width:' + pct + '%;border-radius:4px;"></div></div>'
        + '<p style="margin:0;font-size:12px;color:#666;line-height:1.5;">' + desc + '</p>'
        + '</td></tr>';
    }).join('');

    bfiHtml = _secao('Perfil Comportamental — BFI-44', '#8e24aa')
      + '<div style="background:#f3e5f5;border:1px solid #ce93d8;border-radius:8px;padding:12px 16px;margin-bottom:12px;">'
      + '<p style="margin:0;font-size:13px;color:#4a148c;"><strong>Dimensão dominante:</strong> '
      + DIMS_INFO[dominante.k].nome + ' (' + dominante.v.toFixed(2) + ')</p>'
      + '<p style="margin:5px 0 0;font-size:12px;color:#555;">Os resultados são uma <strong>ferramenta de autoconhecimento</strong> e não determinam aprovação ou eliminação.</p>'
      + '</div>'
      + '<table style="width:100%;border-collapse:collapse;border:1px solid #e0e0e0;border-radius:10px;overflow:hidden;">' + rows + '</table>'
      + '<p style="font-size:11px;color:#aaa;margin-top:8px;font-style:italic;">Escala 1–5 &nbsp;|&nbsp; Alto ≥ 3,6 &nbsp;|&nbsp; Médio 2,5–3,5 &nbsp;|&nbsp; Baixo ≤ 2,4 &nbsp;|&nbsp; ⭐ Dimensão mais alta</p>';
  }

  // ── Formação ───────────────────────────────────────────────────
  const fRows = [];
  if ((dados.graduacao||[]).filter(String).length)
    fRows.push(_linha('Graduação', dados.graduacao.filter(String).map(function(v,i){ return (i+1)+'. '+v; }).join('<br>')));
  if ((dados.pos||[]).filter(String).length)
    fRows.push(_linha('Pós-Graduação', dados.pos.filter(String).map(function(v,i){ return (i+1)+'. '+v; }).join('<br>')));
  if ((dados.mestrado||[]).filter(String).length)
    fRows.push(_linha('Mestrado', dados.mestrado.filter(String).map(function(v,i){ return (i+1)+'. '+v; }).join('<br>')));
  if ((dados.doutorado||[]).filter(String).length)
    fRows.push(_linha('Doutorado', dados.doutorado.filter(String).map(function(v,i){ return (i+1)+'. '+v; }).join('<br>')));

  const formacaoHtml = fRows.length
    ? _secao('Formação Acadêmica')
      + '<table style="width:100%;border-collapse:collapse;">' + fRows.join('') + '</table>'
    : '';

  // ── Cônjuge ────────────────────────────────────────────────────
  let conjugeHtml = '';
  if (conjAtivo) {
    const cRows = [
      _linha('Nome do Cônjuge',   dados.conjugeNome,       true),
      _linha('Matrícula',         dados.conjugeMatricula),
      _linha('E-mail Institucional', dados.conjugeEmail),
      _linha('Tipo de União',     dados.tipoUniao),
      _linha('Data da União',     dados.dataUniao),
      _linha('Endereço 1',        dados.enderecoConjuge1),
      _linha('Endereço 2',        dados.enderecoConjuge2),
      _linha('Lotação 1',         dados.lotacaoConjuge1),
      _linha('Lotação 2',         dados.lotacaoConjuge2),
      _linksArquivos(dados.urlComprovUniao,  'Comprov. União'),
      _linksArquivos(dados.urlComprovCoab,   'Comprov. Coabitação'),
      _linha('Concordância do Cônjuge', 'Aguardando confirmação via e-mail', false),
    ].filter(Boolean).join('');

    conjugeHtml = _secao('Acompanhamento de Cônjuge')
      + '<table style="width:100%;border-collapse:collapse;">' + cRows + '</table>';
  }

  // ── Assinatura do Termo ────────────────────────────────────────
  let assinaturaHtml = '';
  if (dados.assinatura && dados.assinatura.nome) {
    const a = dados.assinatura;
    assinaturaHtml = _secao('Termo de Compromisso', '#43a047')
      + '<div style="background:linear-gradient(135deg,#e8f0fe,#f0f7ff);border:2px solid #1565C0;border-radius:10px;padding:16px 20px;">'
      + '<div style="text-align:center;color:#1565C0;font-weight:700;font-size:12px;margin-bottom:12px;padding-bottom:10px;border-bottom:1px solid rgba(21,101,192,.2);">'
      + '✅ COMPROMISSO ACEITO E ASSINADO ELETRONICAMENTE</div>'
      + '<table style="width:100%;font-size:13px;color:#333;line-height:1.9;border-collapse:collapse;">'
      + _linha('Nome', a.nome, true)
      + _linha('Cargo', a.cargo)
      + _linha('Unidade', a.area)
      + _linha('Data/Hora', a.dataHora)
      + '</table></div>';
  }

  // ── Corpo completo do e-mail ───────────────────────────────────
  const corpo = '<!DOCTYPE html><html lang="pt-BR"><body style="margin:0;padding:0;background:#eef2f7;font-family:\'Segoe UI\',Roboto,Arial,sans-serif;">'
    + '<div style="max-width:620px;margin:24px auto;background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 4px 20px rgba(0,34,68,.12);">'

    // Cabeçalho
    + '<div style="background:linear-gradient(145deg,#001228,#00429a);padding:26px 24px;text-align:center;border-bottom:4px solid #FFC400;">'
    + '<div style="display:inline-block;background:#FFC400;color:#002244;font-size:10px;font-weight:700;padding:3px 14px;border-radius:20px;text-transform:uppercase;letter-spacing:1px;margin-bottom:10px;">Recrutamento Interno</div>'
    + '<h1 style="color:#fff;margin:0;font-size:18px;text-transform:uppercase;letter-spacing:.5px;">Polícia Rodoviária Federal</h1>'
    + '<p style="color:rgba(255,255,255,.7);margin:6px 0 0;font-size:13px;">Inscrição confirmada com sucesso ✅</p>'
    + '</div>'

    // Corpo
    + '<div style="padding:24px 26px;">'
    + '<p style="font-size:15px;color:#1a2332;margin-bottom:4px;">Olá, <strong>' + (dados.nome||'') + '</strong>!</p>'
    + '<p style="color:#555;font-size:13px;margin-top:0;">Sua inscrição no <strong>Recrutamento Interno da PRF</strong> foi registrada em <strong>' + dtStr + '</strong>.</p>'

    // ── Dados pessoais
    + _secao('Dados Pessoais e da Inscrição')
    + '<table style="width:100%;border-collapse:collapse;">'
    + _linha('Nome', (dados.nome||'').toUpperCase(), true)
    + _linha('Matrícula', dados.matricula, true)
    + _linha('Cargo', dados.cargo)
    + _linha('E-mail', email)
    + _linha('Telefone', '(' + (dados.ddd||'') + ') ' + (dados.telefone||''))
    + _linha('Acompan. Cônjuge', conjAtivo ? 'Sim' : 'Não')
    + '</table>'

    // ── Oportunidade
    + _secao('Oportunidade de Lotação')
    + '<table style="width:100%;border-collapse:collapse;">'
    + _linha('Unidade Escolhida', dados.unidadeOportunidade, true)
    + _linha('Conhecimento da Unidade', dados.conhecimentoUnidade)
    + (dados.sougovUrl ? _linksArquivos(dados.sougovUrl, 'Currículo SouGov') : '')
    + '</table>'

    + formacaoHtml
    + conjugeHtml
    + bfiHtml
    + assinaturaHtml

    // Aviso
    + '<div style="background:#fff8e1;border-left:4px solid #FFC400;padding:12px 15px;border-radius:0 8px 8px 0;margin-top:26px;">'
    + '<p style="margin:0;font-size:12px;color:#5d4037;"><strong>⚠️ Atenção:</strong> Guarde este e-mail como comprovante. Em caso de dúvidas, contate a equipe responsável pelo processo seletivo.</p>'
    + '</div>'
    + '</div>'

    // Rodapé
    + '<div style="background:#002244;padding:13px 24px;text-align:center;">'
    + '<p style="color:rgba(255,255,255,.4);font-size:11px;margin:0;">E-mail automático — Não responda.<br>Polícia Rodoviária Federal | DIPROM/DGP | Recrutamento Interno</p>'
    + '</div>'

    + '</div></body></html>';

  const msgObj = {
    to:       email,
    subject:  '✅ Inscrição Confirmada — Recrutamento PRF | ' + (dados.nome||'') + ' | ' + dtStr,
    htmlBody: corpo,
    name:     'Recrutamento PRF'
  };
  if (anexos && anexos.length) msgObj.attachments = anexos;
  MailApp.sendEmail(msgObj);
}

// ─── E-mail de confirmação ao cônjuge ────────────────────────────
function _enviarEmailConjuge(dados, confirmId) {
  const baseUrl   = ScriptApp.getService().getUrl();
  const urlSim    = baseUrl + '?acao=confirmar&id=' + encodeURIComponent(confirmId);
  const urlNao    = baseUrl + '?acao=recusar&id='   + encodeURIComponent(confirmId);

  const nomeConjuge   = dados.conjugeNome        || 'Cônjuge';
  const nomeCandidato = dados.nome               || '';
  const vaga          = dados.unidadeOportunidade|| '';

  const btnBase = 'display:inline-block;padding:14px 32px;border-radius:8px;font-size:15px;font-weight:700;text-decoration:none;letter-spacing:.3px;';
  const btnSim  = btnBase + 'background:#1b5e20;color:#ffffff;margin-right:12px;';
  const btnNao  = btnBase + 'background:#b71c1c;color:#ffffff;';

  const corpo = '<!DOCTYPE html><html lang="pt-BR"><body style="margin:0;padding:0;background:#eef2f7;font-family:\'Segoe UI\',Roboto,Arial,sans-serif;">'
    + '<div style="max-width:600px;margin:24px auto;background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 4px 20px rgba(0,34,68,.12);">'

    // Cabeçalho
    + '<div style="background:linear-gradient(145deg,#001228,#00429a);padding:26px 24px;text-align:center;border-bottom:4px solid #FFC400;">'
    + '<div style="display:inline-block;background:#FFC400;color:#002244;font-size:10px;font-weight:700;padding:3px 14px;border-radius:20px;text-transform:uppercase;letter-spacing:1px;margin-bottom:10px;">Recrutamento Interno</div>'
    + '<h1 style="color:#fff;margin:0;font-size:18px;text-transform:uppercase;letter-spacing:.5px;">Polícia Rodoviária Federal</h1>'
    + '<p style="color:rgba(255,255,255,.7);margin:6px 0 0;font-size:13px;">Solicitação de Acompanhamento de Cônjuge</p>'
    + '</div>'

    // Corpo
    + '<div style="padding:28px 30px;color:#333;">'
    + '<p style="font-size:15px;line-height:1.7;margin-top:0;">Olá, <strong>' + nomeConjuge + '</strong>.</p>'
    + '<p style="font-size:14px;line-height:1.8;color:#444;">Informamos que <strong>' + nomeCandidato + '</strong> está participando de um processo de recrutamento e seleção para a vaga de <strong>' + vaga + '</strong>.</p>'
    + '<p style="font-size:14px;line-height:1.8;color:#444;">Durante o cadastro, foi manifestado o interesse em sua inclusão como acompanhante, caso o candidato seja selecionado para a posição. Para prosseguirmos com esta solicitação, precisamos da sua <strong>confirmação oficial de concordância</strong>.</p>'

    // Destaque
    + '<div style="background:#e8f0fe;border-left:4px solid #1565C0;border-radius:0 8px 8px 0;padding:14px 18px;margin:20px 0;">'
    + '<p style="margin:0;font-size:14px;font-weight:600;color:#1565C0;">Você concorda com esta solicitação de acompanhamento?</p>'
    + '</div>'

    // Botões
    + '<div style="text-align:center;margin:28px 0;">'
    + '<a href="' + urlSim + '" style="' + btnSim + '">✅ SIM, EU CONCORDO</a>'
    + '<a href="' + urlNao + '" style="' + btnNao + '">❌ NÃO CONCORDO</a>'
    + '</div>'

    + '<p style="font-size:12px;color:#777;line-height:1.6;">Ao clicar em uma das opções acima, sua resposta será registrada automaticamente em nosso sistema. Caso os botões não funcionem, copie e cole um dos links abaixo no seu navegador:</p>'
    + '<p style="font-size:11px;color:#aaa;word-break:break-all;">Concordar: ' + urlSim + '</p>'
    + '<p style="font-size:11px;color:#aaa;word-break:break-all;">Recusar: '   + urlNao + '</p>'
    + '</div>'

    // Rodapé
    + '<div style="background:#002244;padding:13px 24px;text-align:center;">'
    + '<p style="color:rgba(255,255,255,.5);font-size:11px;margin:0;">Atenciosamente, Equipe de Recrutamento<br>Polícia Rodoviária Federal | DIPROM/DGP<br>E-mail automático — Não responda.</p>'
    + '</div>'
    + '</div></body></html>';

  MailApp.sendEmail({
    to:       dados.conjugeEmail,
    subject:  '📋 Solicitação de Acompanhamento — Recrutamento PRF | ' + nomeCandidato,
    htmlBody: corpo,
    name:     'Recrutamento PRF'
  });
}

// ─── Processamento da resposta do cônjuge (via URL params) ────────
function _processarRespostaConjuge(id, acao) {
  if (!id) {
    return _paginaRespostaConjuge('Link inválido',
      'O link utilizado não é válido ou está incompleto. Entre em contato com a equipe de recrutamento.',
      '#c0392b', '❌');
  }

  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('respostas');
  if (!sheet) {
    return _paginaRespostaConjuge('Erro interno',
      'Não foi possível localizar os dados. Entre em contato com a equipe de recrutamento.',
      '#c0392b', '⚠️');
  }

  const COL_ID          = 33; // 0-indexed (HEADERS: 'ID Confirmação')
  const COL_STATUS      = 34; // 0-indexed (HEADERS: 'Status Cônjuge')
  const COL_DATA        = 35; // 0-indexed (HEADERS: 'Data Confirmação Cônjuge')
  const COL_PDF_CONCORD = 38; // 0-indexed (HEADERS: 'PDF Concordância Cônjuge')

  const allRows = sheet.getDataRange().getValues();
  let rowIdx = -1;
  for (let i = 1; i < allRows.length; i++) {
    if (String(allRows[i][COL_ID] || '') === String(id)) {
      rowIdx = i + 1; // 1-indexed para o sheet
      break;
    }
  }

  if (rowIdx === -1) {
    return _paginaRespostaConjuge('Solicitação não encontrada',
      'Não encontramos uma solicitação correspondente a este link. Ele pode ter expirado ou já foi processado.',
      '#e67e22', '🔍');
  }

  const statusAtual = String(allRows[rowIdx - 1][COL_STATUS] || '');
  if (statusAtual === 'Confirmado' || statusAtual === 'Recusado') {
    return _paginaRespostaConjuge('Resposta já registrada',
      'Sua resposta (<strong>' + statusAtual + '</strong>) já foi registrada anteriormente em nosso sistema. Não é necessária nenhuma ação adicional.',
      '#e67e22', '📋');
  }

  const novoStatus = (acao === 'confirmar') ? 'Confirmado' : 'Recusado';
  const dataHora   = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm:ss');

  sheet.getRange(rowIdx, COL_STATUS + 1).setValue(novoStatus);
  sheet.getRange(rowIdx, COL_DATA   + 1).setValue(dataHora);

  // Gerar PDF de concordância apenas quando o cônjuge confirma
  if (acao === 'confirmar') {
    try {
      const rowData     = allRows[rowIdx - 1];
      const nomeCand    = String(rowData[2]  || '');
      const nomeConj    = String(rowData[15] || '');
      const emailConj   = String(rowData[17] || '');
      const nomeSan     = (nomeConj || 'Conjuge').replace(/[^a-zA-Z0-9À-ú ]/g, '').trim();
      const pdfResult   = _gerarESalvarPdf(
        _htmlPdfConcordancia(rowData, nomeConj, emailConj, dataHora),
        'Concordancia_' + nomeSan
      );
      if (pdfResult.url) {
        sheet.getRange(rowIdx, COL_PDF_CONCORD + 1).setValue(pdfResult.url);
      }
    } catch(errPdf) {
      Logger.log('[concordância PDF] ' + errPdf.message);
    }

    return _paginaRespostaConjuge('Confirmação registrada',
      'Obrigado! Sua concordância com o acompanhamento foi registrada com sucesso em nosso sistema. Não é necessária nenhuma ação adicional.',
      '#1b5e20', '✅');
  } else {
    return _paginaRespostaConjuge('Recusa registrada',
      'Sua resposta foi registrada. A solicitação de acompanhamento <strong>não será processada</strong>. Obrigado pelo retorno.',
      '#b71c1c', '❌');
  }
}

// ─── Página HTML simples para resposta do cônjuge ─────────────────
function _paginaRespostaConjuge(titulo, mensagem, corDestaque, icone) {
  const html = '<!DOCTYPE html><html lang="pt-BR"><head>'
    + '<meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">'
    + '<title>Recrutamento PRF</title>'
    + '<style>'
    + 'body{margin:0;padding:0;background:#eef2f7;font-family:\'Segoe UI\',Roboto,Arial,sans-serif;display:flex;align-items:center;justify-content:center;min-height:100vh;}'
    + '.card{background:#fff;border-radius:14px;box-shadow:0 4px 24px rgba(0,34,68,.13);max-width:440px;width:90%;overflow:hidden;}'
    + '.topo{background:linear-gradient(145deg,#001228,#00429a);padding:22px 24px;text-align:center;border-bottom:4px solid #FFC400;}'
    + '.topo p{color:rgba(255,255,255,.65);margin:4px 0 0;font-size:12px;}'
    + '.topo h2{color:#fff;margin:0;font-size:15px;text-transform:uppercase;letter-spacing:.5px;}'
    + '.corpo{padding:30px 28px;text-align:center;}'
    + '.icone{font-size:52px;margin-bottom:12px;}'
    + '.titulo{font-size:19px;font-weight:700;color:' + corDestaque + ';margin:0 0 14px;}'
    + '.msg{font-size:14px;color:#555;line-height:1.75;margin:0;}'
    + '.rodape{background:#f5f5f5;padding:12px;text-align:center;font-size:11px;color:#aaa;border-top:1px solid #eee;}'
    + '</style></head><body>'
    + '<div class="card">'
    + '<div class="topo">'
    + '<h2>Polícia Rodoviária Federal</h2>'
    + '<p>Recrutamento Interno — Confirmação de Cônjuge</p>'
    + '</div>'
    + '<div class="corpo">'
    + '<div class="icone">' + (icone || '📋') + '</div>'
    + '<h3 class="titulo">' + titulo + '</h3>'
    + '<p class="msg">' + mensagem + '</p>'
    + '</div>'
    + '<div class="rodape">DIPROM/DGP | Polícia Rodoviária Federal</div>'
    + '</div></body></html>';

  return HtmlService.createHtmlOutput(html)
    .setTitle('Recrutamento PRF — Confirmação')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ════════════════════════════════════════════════════════════════
//  GERAÇÃO DE PDF
// ════════════════════════════════════════════════════════════════

// ID da pasta no Drive onde os PDFs serão salvos.
// Deixe vazio ('') para criar/usar automaticamente a pasta "Recrutamento PRF".
const DRIVE_FOLDER_ID = '';

function _getPastaRecrutamento() {
  if (DRIVE_FOLDER_ID) return DriveApp.getFolderById(DRIVE_FOLDER_ID);
  const iter = DriveApp.getFoldersByName('Recrutamento PRF');
  return iter.hasNext() ? iter.next() : DriveApp.createFolder('Recrutamento PRF');
}

// Lê a lista de e-mails configurados na aba "config" da planilha ativa.
// Retorna array de strings (pode ser vazio).
function _lerEmailsResposta() {
  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('config');
    if (!sheet) return [];
    const rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0] || '').trim().toLowerCase() === 'emailsresposta') {
        return String(rows[i][1] || '').split(',')
          .map(function(e){ return e.trim(); })
          .filter(Boolean);
      }
    }
  } catch(e) {
    Logger.log('[_lerEmailsResposta] ' + e.message);
  }
  return [];
}

// Converte HTML em PDF via Drive API, salva na pasta e retorna { url, blob }.
// Lança exceção em caso de falha (para que processarInscricao possa registrar o erro).
function _gerarESalvarPdf(htmlContent, nomeBase) {
  const token    = ScriptApp.getOAuthToken();
  const boundary = 'PRF' + Utilities.getUuid().replace(/-/g, '');
  const metaJson = JSON.stringify({ name: nomeBase, mimeType: MimeType.GOOGLE_DOCS });

  // Monta o payload multipart como array de bytes para suporte correto a UTF-8
  const cabecalho =
    '--' + boundary + '\r\n' +
    'Content-Type: application/json; charset=UTF-8\r\n\r\n' +
    metaJson + '\r\n' +
    '--' + boundary + '\r\n' +
    'Content-Type: text/html; charset=UTF-8\r\n\r\n';
  const rodape = '\r\n--' + boundary + '--';

  const payloadBytes = Utilities.newBlob(cabecalho, 'text/plain').getBytes()
    .concat(Utilities.newBlob(htmlContent, MimeType.HTML).getBytes())
    .concat(Utilities.newBlob(rodape, 'text/plain').getBytes());

  const resp = UrlFetchApp.fetch(
    'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart',
    {
      method:             'POST',
      headers:            {
        Authorization:  'Bearer ' + token,
        'Content-Type': 'multipart/related; boundary=' + boundary
      },
      payload:            payloadBytes,
      muteHttpExceptions: true
    }
  );

  if (resp.getResponseCode() !== 200) {
    throw new Error(
      'Drive API retornou HTTP ' + resp.getResponseCode() +
      ': ' + resp.getContentText().substring(0, 300)
    );
  }

  const docId   = JSON.parse(resp.getContentText()).id;
  const pdfBlob = DriveApp.getFileById(docId)
    .getAs(MimeType.PDF)
    .setName(nomeBase + '.pdf');
  DriveApp.getFileById(docId).setTrashed(true);

  const pasta   = _getPastaRecrutamento();
  const pdfFile = pasta.createFile(pdfBlob);
  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return { url: pdfFile.getUrl(), blob: pdfFile.getBlob().setName(nomeBase + '.pdf') };
}

// Regenera PDFs para todas as linhas que estejam sem URL nos campos de PDF.
// Pode ser chamada manualmente via GAS Editor (Executar > regerarPdfsFaltantes).
function regerarPdfsFaltantes() {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('respostas');
  if (!sheet || sheet.getLastRow() <= 1) { Logger.log('Nenhuma linha encontrada.'); return; }

  const rows = sheet.getDataRange().getValues();
  const HDR  = rows[0];
  const iRespostas = HDR.indexOf('PDF Respostas');  // coluna 36 (0-based)
  const iTermo     = HDR.indexOf('PDF Termo');      // coluna 37

  if (iRespostas === -1 || iTermo === -1) {
    Logger.log('Colunas PDF Respostas / PDF Termo não encontradas no cabeçalho.');
    return;
  }

  let atualizados = 0;
  for (let i = 1; i < rows.length; i++) {
    const r             = rows[i];
    const urlRespostas  = String(r[iRespostas] || '').trim();
    const urlTermo      = String(r[iTermo]     || '').trim();
    if (urlRespostas && urlTermo) continue; // já tem ambos, pula

    // Reconstruir objeto "dados" mínimo a partir da linha
    const dados = {
      nome:               String(r[2]  || ''),
      matricula:          String(r[3]  || ''),
      cargo:              String(r[4]  || ''),
      unidadeOportunidade:String(r[5]  || ''),
      conhecimentoUnidade:String(r[6]  || ''),
      ddd:                String(r[7]  || ''),
      telefone:           String(r[8]  || ''),
      conjuge:            String(r[9]  || '').toLowerCase() === 'sim' ? 'sim' : 'não',
      sougovUrl:          String(r[14] || ''),
      conjugeNome:        String(r[15] || ''),
      conjugeMatricula:   String(r[16] || ''),
      conjugeEmail:       String(r[17] || ''),
      tipoUniao:          String(r[18] || ''),
      dataUniao:          String(r[19] || ''),
      enderecoConjuge1:   String(r[20] || ''),
      enderecoConjuge2:   String(r[21] || ''),
      lotacaoConjuge1:    String(r[22] || ''),
      lotacaoConjuge2:    String(r[23] || ''),
      urlComprovUniao:    String(r[24] || ''),
      urlComprovCoab:     String(r[25] || ''),
    };
    dados._email = String(r[1] || '');

    // BFI
    const bfiValues = [r[26], r[27], r[28], r[29], r[30]].map(Number);
    const bfiResult = bfiValues.every(function(v) { return !isNaN(v) && v > 0; })
      ? { ext: bfiValues[0], amab: bfiValues[1], cons: bfiValues[2], estab: bfiValues[3], abert: bfiValues[4] }
      : null;

    const agora = r[0] instanceof Date ? r[0] : new Date();
    const nomeSanitizado = dados.nome.replace(/[^a-zA-Z0-9À-ú ]/g, '').trim();

    try {
      const pdfR = urlRespostas ? { url: urlRespostas } : _gerarESalvarPdf(_htmlPdfRespostas(dados, bfiResult, agora), 'Ficha_' + nomeSanitizado);
      const pdfT = urlTermo     ? { url: urlTermo }     : _gerarESalvarPdf(_htmlPdfTermo(dados),                     'Termo_' + nomeSanitizado);
      sheet.getRange(i + 1, iRespostas + 1, 1, 2).setValues([[pdfR.url, pdfT.url]]);
      Logger.log('Linha ' + (i + 1) + ' (' + dados.nome + '): PDFs gerados OK.');
      atualizados++;
    } catch (err) {
      Logger.log('Linha ' + (i + 1) + ' (' + dados.nome + '): ERRO — ' + err.message);
    }
  }
  Logger.log('regerarPdfsFaltantes: ' + atualizados + ' linha(s) atualizada(s).');
}

// ─── Auxiliares de formatação para PDFs ──────────────────────────
function _pdfCabecalho(titulo) {
  return '<div style="text-align:center;margin-bottom:20px;padding-bottom:12px;border-bottom:3px solid #002244;">'
    + '<h2 style="color:#002244;margin:0 0 4px;font-size:16pt;font-family:Arial;">POLÍCIA RODOVIÁRIA FEDERAL</h2>'
    + '<p style="color:#666;margin:0 0 10px;font-size:10pt;font-family:Arial;">DIPROM/DGP — Recrutamento Interno</p>'
    + '<h3 style="color:#001228;margin:0;font-size:13pt;font-family:Arial;">' + titulo + '</h3>'
    + '</div>';
}

function _pdfSecao(titulo) {
  return '<h4 style="background:#002244;color:#fff;padding:5px 10px;font-size:11pt;'
    + 'font-family:Arial;margin:18px 0 6px;border-radius:3px;">' + titulo + '</h4>';
}

function _pdfLinha(label, valor) {
  if (!valor && valor !== 0) return '';
  return '<tr>'
    + '<td style="padding:4px 10px 4px 0;font-weight:bold;width:38%;color:#333;font-size:10pt;font-family:Arial;vertical-align:top;">' + label + ':</td>'
    + '<td style="padding:4px 0;color:#555;font-size:10pt;font-family:Arial;vertical-align:top;">' + valor + '</td>'
    + '</tr>';
}

function _pdfTabela(linhas) {
  if (!linhas) return '';
  return '<table style="width:100%;border-collapse:collapse;margin-bottom:6px;">' + linhas + '</table>';
}

// ─── Bloco reutilizável: Termo de Compromisso ─────────────────────
function _htmlTermoCompromisso(dados) {
  const assinatura = (typeof dados.assinatura === 'string')
    ? JSON.parse(dados.assinatura || '{}') : (dados.assinatura || {});

  const dataFmt = assinatura.dataHora || Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm:ss');

  return _pdfSecao('Termo de Compromisso')
    + '<div style="border:1px solid #002244;padding:14px 16px;margin:6px 0;border-radius:4px;">'
    + '<p style="text-align:justify;font-size:10pt;font-family:Arial;line-height:1.7;margin:0 0 10px;">'
    + 'Eu, <strong>' + (dados.nome || '') + '</strong>, matrícula <strong>'
    + (dados.matricula || '') + '</strong>, cargo <strong>' + (dados.cargo || '') + '</strong>, '
    + 'declaro, para os devidos fins, que estou ciente e de acordo com as normas e critérios '
    + 'estabelecidos para o presente processo de recrutamento interno da Polícia Rodoviária Federal, '
    + 'e que as informações prestadas neste formulário são verdadeiras, sob pena de responsabilização '
    + 'disciplinar e/ou legal em caso de falsidade.'
    + '</p>'
    + '<p style="text-align:justify;font-size:10pt;font-family:Arial;line-height:1.7;margin:0;">'
    + 'Comprometo-me a cumprir todas as etapas e exigências do processo seletivo, bem como a aceitar '
    + 'as decisões da Comissão de Recrutamento e a apresentar, quando solicitado, a documentação '
    + 'comprobatória das informações declaradas.'
    + '</p>'
    + '</div>'
    + (assinatura.nome
      ? '<div style="margin-top:14px;padding-top:10px;border-top:1px solid #ccc;">'
        + '<p style="font-size:10pt;font-family:Arial;font-weight:bold;margin:0 0 6px;">✅ Assinado eletronicamente:</p>'
        + _pdfTabela(
            _pdfLinha('Nome', assinatura.nome)
            + _pdfLinha('Cargo', assinatura.cargo)
            + _pdfLinha('Unidade', assinatura.area)
            + _pdfLinha('Data/Hora', dataFmt)
          )
        + '</div>'
      : '');
}

// ─── PDF 1: Ficha completa de respostas ───────────────────────────
function _htmlPdfRespostas(dados, bfiResult, agora) {
  const dtStr = Utilities.formatDate(agora, 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm:ss');
  const conj  = dados.conjuge === 'sim' || dados.conjuge === true;

  let html = '<html><body style="font-family:Arial,sans-serif;font-size:10pt;color:#333;margin:20px;">';
  html += _pdfCabecalho('Ficha de Inscrição — Processo de Recrutamento Interno');
  html += '<p style="text-align:right;font-size:9pt;color:#888;margin:0 0 10px;">Gerado em: ' + dtStr + '</p>';

  // Dados Pessoais
  html += _pdfSecao('Dados Pessoais');
  html += _pdfTabela(
    _pdfLinha('Nome', dados.nome)
    + _pdfLinha('Matrícula', dados.matricula)
    + _pdfLinha('Cargo', dados.cargo)
    + _pdfLinha('E-mail', dados._email || '')
    + _pdfLinha('Telefone', dados.ddd ? '(' + dados.ddd + ') ' + dados.telefone : dados.telefone)
  );

  // Oportunidade
  html += _pdfSecao('Oportunidade Pleiteada');
  html += _pdfTabela(
    _pdfLinha('Unidade Desejada', dados.unidadeOportunidade)
    + _pdfLinha('Conhecimento da Unidade', dados.conhecimentoUnidade)
  );

  // Formação
  const fLinhas = [
    _pdfLinha('Graduação', (dados.graduacao || []).join('; ')),
    _pdfLinha('Pós-Graduação', (dados.pos || []).join('; ')),
    _pdfLinha('Mestrado', (dados.mestrado || []).join('; ')),
    _pdfLinha('Doutorado', (dados.doutorado || []).join('; '))
  ].join('');
  if (fLinhas) {
    html += _pdfSecao('Formação Acadêmica');
    html += _pdfTabela(fLinhas);
  }

  // Cônjuge
  if (conj) {
    html += _pdfSecao('Dados do Cônjuge');
    html += _pdfTabela(
      _pdfLinha('Nome', dados.conjugeNome)
      + _pdfLinha('Matrícula', dados.conjugeMatricula)
      + _pdfLinha('E-mail', dados.conjugeEmail)
      + _pdfLinha('Tipo de União', dados.tipoUniao)
      + _pdfLinha('Data da União', dados.dataUniao)
      + _pdfLinha('Endereço Atual', dados.enderecoConjuge1)
      + _pdfLinha('Endereço Destino', dados.enderecoConjuge2)
      + _pdfLinha('Lotação Atual', dados.lotacaoConjuge1)
      + _pdfLinha('Lotação Destino', dados.lotacaoConjuge2)
    );
  }

  // BFI
  if (bfiResult) {
    const dims = [
      { n: 'Extroversão',            v: bfiResult.ext   },
      { n: 'Amabilidade',            v: bfiResult.amab  },
      { n: 'Conscienciosidade',      v: bfiResult.cons  },
      { n: 'Estabilidade Emocional', v: bfiResult.estab },
      { n: 'Abertura à Experiência', v: bfiResult.abert }
    ];
    const maxVal = Math.max.apply(null, dims.map(function(d){ return d.v; }));
    html += _pdfSecao('Resultado BFI-44');
    html += _pdfTabela(dims.map(function(d) {
      const pct  = Math.round((d.v / 5) * 100);
      const dest = d.v === maxVal ? ' ⭐' : '';
      return _pdfLinha(d.n + dest, d.v.toFixed(2) + ' / 5.00  (' + pct + '%)');
    }).join(''));
  }

  // Termo
  html += _htmlTermoCompromisso(dados);

  html += '</body></html>';
  return html;
}

// ─── PDF 2: Termo de Compromisso (isolado) ────────────────────────
function _htmlPdfTermo(dados) {
  const dtStr = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm:ss');
  return '<html><body style="font-family:Arial,sans-serif;font-size:10pt;color:#333;margin:20px;">'
    + _pdfCabecalho('Termo de Compromisso — Recrutamento Interno PRF')
    + '<p style="text-align:right;font-size:9pt;color:#888;margin:0 0 10px;">Gerado em: ' + dtStr + '</p>'
    + _pdfSecao('Identificação do Servidor')
    + _pdfTabela(
        _pdfLinha('Nome', dados.nome)
        + _pdfLinha('Matrícula', dados.matricula)
        + _pdfLinha('Cargo', dados.cargo)
        + _pdfLinha('E-mail', dados._email || '')
        + _pdfLinha('Unidade Pleiteada', dados.unidadeOportunidade)
      )
    + _htmlTermoCompromisso(dados)
    + '</body></html>';
}

// ─── PDF 3: Declaração de Concordância do Cônjuge ─────────────────
function _htmlPdfConcordancia(rowData, nomeConj, emailConj, dataConfirmacao) {
  const nomeCand = String(rowData[2]  || '');
  const matricCand = String(rowData[3] || '');
  const cargoCand  = String(rowData[4] || '');
  const unidade    = String(rowData[5] || '');

  return '<html><body style="font-family:Arial,sans-serif;font-size:10pt;color:#333;margin:20px;">'
    + _pdfCabecalho('Declaração de Concordância — Acompanhamento de Cônjuge')
    + '<p style="text-align:right;font-size:9pt;color:#888;margin:0 0 10px;">Confirmado em: ' + dataConfirmacao + '</p>'

    + _pdfSecao('Dados da Solicitação')
    + _pdfTabela(
        _pdfLinha('Candidato', nomeCand)
        + _pdfLinha('Matrícula', matricCand)
        + _pdfLinha('Cargo', cargoCand)
        + _pdfLinha('Unidade Desejada', unidade)
      )

    + _pdfSecao('Declarante')
    + _pdfTabela(
        _pdfLinha('Nome do Cônjuge', nomeConj)
        + _pdfLinha('E-mail Institucional', emailConj)
      )

    + _pdfSecao('Declaração')
    + '<div style="border:1px solid #002244;padding:14px 16px;margin:6px 0;border-radius:4px;">'
    + '<p style="text-align:justify;font-size:10pt;font-family:Arial;line-height:1.7;margin:0 0 10px;">'
    + 'Eu, <strong>' + nomeConj + '</strong>, servidor(a) da Polícia Rodoviária Federal, '
    + 'declaro expressamente que estou ciente e <strong>concordo</strong> com a solicitação de '
    + 'acompanhamento de cônjuge apresentada pelo(a) servidor(a) <strong>' + nomeCand + '</strong>, '
    + 'matrícula ' + matricCand + ', no processo de recrutamento interno para a vaga de '
    + '<strong>' + unidade + '</strong>.'
    + '</p>'
    + '<p style="text-align:justify;font-size:10pt;font-family:Arial;line-height:1.7;margin:0;">'
    + 'Esta declaração foi gerada eletronicamente a partir da confirmação realizada por meio de '
    + 'link enviado ao e-mail institucional (@prf.gov.br) do declarante.'
    + '</p>'
    + '</div>'

    + '<div style="margin-top:16px;padding-top:10px;border-top:1px solid #ccc;">'
    + '<p style="font-size:10pt;font-family:Arial;font-weight:bold;margin:0 0 6px;">✅ Confirmação eletrônica:</p>'
    + _pdfTabela(
        _pdfLinha('E-mail do confirmante', emailConj)
        + _pdfLinha('Data/Hora', dataConfirmacao)
        + _pdfLinha('Método', 'Clique em link enviado ao e-mail institucional (@prf.gov.br)')
      )
    + '</div>'
    + '</body></html>';
}

// ════════════════════════════════════════════════════════════════
//  CRIAR NOVO RECRUTAMENTO
// ════════════════════════════════════════════════════════════════

function criarNovoRecrutamento(dados) {
  const u = verificarAcessoPainel();
  if (u.perfil !== 'ADMINISTRADOR') throw new Error('Acesso restrito a administradores.');

  // Aceita string simples (legado) ou objeto estruturado
  var dadosObj = (typeof dados === 'string') ? { nome: dados } : (dados || {});
  const nome = ((dadosObj.nome || '') + '').trim();
  if (!nome) throw new Error('Informe o nome do recrutamento.');

  const uorgs               = dadosObj.uorgs               || [];
  const unidadesOportunidade = dadosObj.unidadesOportunidade || [];
  const graduacoes           = dadosObj.graduacoes           || [];
  const posGraduacoes        = dadosObj.posGraduacoes        || [];
  const emailsResposta       = dadosObj.emailsResposta       || '';

  // ── Cria a planilha ──────────────────────────────────────────
  const ss  = SpreadsheetApp.create(nome);
  const id  = ss.getId();
  const url = ss.getUrl();

  // ── Formata cabeçalho de qualquer aba ───────────────────────
  const _cab = function(sheet, headers) {
    const hdr = sheet.getRange(1, 1, 1, headers.length);
    hdr.setValues([headers]);
    hdr.setFontWeight('bold').setBackground('#002244').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    sheet.setColumnWidths(1, headers.length, 180);
    SpreadsheetApp.flush();
  };

  // ── Aba INSTRUÇÕES (renomeia a Sheet1 padrão) ────────────────
  const sheetInst = ss.getSheets()[0];
  sheetInst.setName('INSTRUÇÕES');
  _preencherInstrucoesNovoRecrutamento(sheetInst, id, url, nome);

  // ── Aba respostas ────────────────────────────────────────────
  const sRespostas = ss.insertSheet('respostas');
  _cab(sRespostas, [
    'Data/Hora','E-mail','Nome','Matrícula','Cargo',
    'Unidade Oportunidade','Conhecimento da Unidade',
    'DDD','Telefone','Cônjuge',
    'Graduações','Pós-Graduações','Mestrados','Doutorados','Currículo SouGov',
    'Cônjuge Nome','Cônjuge Matrícula','Cônjuge E-mail','Tipo de União','Data da União',
    'Endereço Cônjuge 1','Endereço Cônjuge 2',
    'Lotação Cônjuge 1','Lotação Cônjuge 2',
    'URL Comprov. União','URL Comprob. Coabitação',
    'Extroversão','Amabilidade','Conscienciosidade','Estab. Emocional','Abertura',
    'Assinatura','Status',
    'ID Confirmação','Status Cônjuge','Data Confirmação Cônjuge',
    'PDF Respostas','PDF Termo','PDF Concordância Cônjuge'
  ]);

  // ── Aba bfi_resultados ───────────────────────────────────────
  const sBfi = ss.insertSheet('bfi_resultados');
  _cab(sBfi, ['Data/Hora','E-mail','Nome','Matrícula','Cargo','Área Desejada',
    'Extroversão','Amabilidade','Conscienciosidade','Estab. Emocional','Abertura','Respostas Brutas']);

  // ── Aba credenciais (já insere o criador como admin) ─────────
  const sCred = ss.insertSheet('credenciais');
  _cab(sCred, ['E-mail','Nome','Perfil','Unidades Autorizadas','Data Cadastro','Status']);
  const agora = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm:ss');
  sCred.appendRow([u.email, u.nome || 'Administrador Principal', 'ADMINISTRADOR', '*', agora, 'Ativo']);

  // ── Aba avaliacoes ───────────────────────────────────────────
  const sAval = ss.insertSheet('avaliacoes');
  _cab(sAval, [
    'E-mail Candidato','Nome Candidato','Unidade Candidato',
    'E-mail Avaliador','Nome Avaliador','Tipo','Data/Hora',
    'Nota Geral',
    'Nota Comunicação','Nota Int. Emocional','Nota Postura','Nota Ling. Corporal',
    'Unidade Indicada','Obs Positivas','Obs Negativas','Obs Livres','Status'
  ]);

  // ── Aba graduacao ────────────────────────────────────────────
  const sGrad = ss.insertSheet('graduacao');
  _cab(sGrad, ['Graduação', 'Pós-Graduação / Especialização']);
  if (graduacoes.length > 0 || posGraduacoes.length > 0) {
    const maxLen = Math.max(graduacoes.length, posGraduacoes.length);
    const gradRows = [];
    for (var gi = 0; gi < maxLen; gi++) {
      gradRows.push([ graduacoes[gi] || '', posGraduacoes[gi] || '' ]);
    }
    sGrad.getRange(2, 1, gradRows.length, 2).setValues(gradRows);
  } else {
    [['Direito','Pós-Graduação em Direito Público'],
     ['Administração','MBA em Gestão Pública'],
     ['Ciência da Computação','Pós-Graduação em Segurança da Informação'],
     ['Engenharia','Especialização em Gestão de Projetos'],
     ['Psicologia','Especialização em Psicologia Organizacional'],
     ['Contabilidade','MBA em Auditoria e Controle'],
     ['Economia','Pós-Graduação em Economia do Setor Público']
    ].forEach(function(r){ sGrad.appendRow(r); });
  }

  // ── Aba uorgs ────────────────────────────────────────────────
  const sUorgs = ss.insertSheet('uorgs');
  _cab(sUorgs, ['Sigla / Nome', 'Nome Completo (referência)']);
  if (uorgs.length > 0) {
    const uorgRows = uorgs.map(function(u){ return [String(u.sigla || ''), String(u.nome || '')]; });
    sUorgs.getRange(2, 1, uorgRows.length, 2).setValues(uorgRows);
  } else {
    sUorgs.getRange('A2').setValue('← Preencha com as unidades da PRF. Ex: SRPRF/DF, DPRF/SP...');
  }

  // ── Aba unidades_oportunidade ────────────────────────────────
  const sOport = ss.insertSheet('unidades_oportunidade');
  _cab(sOport, ['Unidade de Oportunidade']);
  if (unidadesOportunidade.length > 0) {
    const oportRows = unidadesOportunidade.map(function(v){ return [String(v)]; });
    sOport.getRange(2, 1, oportRows.length, 1).setValues(oportRows);
  } else {
    sOport.getRange('A2').setValue('← Preencha com as unidades que têm vagas abertas neste recrutamento.');
  }

  // ── Aba rascunhos_bfi ────────────────────────────────────────
  const sRasc = ss.insertSheet('rascunhos_bfi');
  _cab(sRasc, ['Data/Hora','E-mail','Respostas (rascunho)','Respondidas']);

  // ── Aba config ───────────────────────────────────────────────
  const sCfg = ss.insertSheet('config');
  const cfgHdr = sCfg.getRange(1, 1, 1, 2);
  cfgHdr.setValues([['Chave', 'Valor']]);
  cfgHdr.setFontWeight('bold').setBackground('#002244').setFontColor('#ffffff');
  sCfg.setFrozenRows(1);
  sCfg.setColumnWidth(1, 260);
  sCfg.setColumnWidth(2, 420);
  sCfg.appendRow(['emailsResposta', emailsResposta]);
  sCfg.getRange('A2').setNote(
    'E-mails que receberão notificação quando o cônjuge confirmar a concordância. ' +
    'Separe múltiplos com vírgula. Ex: fulano@prf.gov.br, ciclano@prf.gov.br'
  );

  SpreadsheetApp.flush();

  // ── Tenta enviar e-mail com o link e instruções ──────────────
  var emailEnviado = false;
  try {
    MailApp.sendEmail({
      to:      u.email,
      subject: '[PRF Recrutamento] Nova planilha criada: ' + nome,
      body:
        'Sua nova planilha de recrutamento foi criada com sucesso.\n\n'
        + '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n'
        + 'NOME: ' + nome + '\n'
        + 'LINK: ' + url + '\n'
        + 'ID:   ' + id  + '\n'
        + '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n'
        + 'PRÓXIMOS PASSOS:\n\n'
        + '1. Abra o projeto GAS (no formulário original: Extensões > Apps Script)\n'
        + '2. Localize a linha: const SPREADSHEET_ID = \'...\'\n'
        + '3. Substitua pelo ID acima: \'' + id + '\'\n'
        + '4. Salve (Ctrl+S)\n'
        + '5. Clique em Implantar > Gerenciar implantações > lápis (Editar) > Nova versão > Implantar\n'
        + '6. Copie o URL da implantação e distribua para os candidatos\n\n'
        + 'Veja a aba "INSTRUÇÕES" da nova planilha para orientações completas.'
    });
    emailEnviado = true;
  } catch (e) {
    // MailApp pode não ter autorização — não é erro crítico
    emailEnviado = false;
  }

  return { nome: nome, url: url, id: id, emailEnviado: emailEnviado };
}

// ── Preenche a aba de Instruções da nova planilha ────────────────
function _preencherInstrucoesNovoRecrutamento(sheet, id, url, nome) {
  sheet.setColumnWidth(1, 240);
  sheet.setColumnWidth(2, 520);

  const dados = [
    // [linha de início, col início, valor, fontWeight, fontSize, bgColor, fontColor, wrapStrategy]
  ];

  // Cabeçalho principal
  sheet.getRange('A1:B1').merge().setValue('SISTEMA DE RECRUTAMENTO INTERNO — PRF')
    .setBackground('#002244').setFontColor('#FFC400')
    .setFontWeight('bold').setFontSize(14).setHorizontalAlignment('center');
  sheet.setRowHeight(1, 38);

  // Subtítulo
  sheet.getRange('A2:B2').merge().setValue(nome)
    .setBackground('#1565C0').setFontColor('#ffffff')
    .setFontWeight('bold').setFontSize(11).setHorizontalAlignment('center');
  sheet.setRowHeight(2, 28);

  var r = 4; // linha atual

  // ── Bloco: ID e URL ──────────────────────────────────────────
  sheet.getRange(r, 1, 1, 2).merge().setValue('📋  IDENTIFICAÇÃO DA PLANILHA')
    .setBackground('#002244').setFontColor('#ffffff').setFontWeight('bold').setFontSize(10);
  sheet.setRowHeight(r, 24); r++;

  sheet.getRange(r, 1).setValue('ID da Planilha').setFontWeight('bold');
  sheet.getRange(r, 2).setValue(id)
    .setBackground('#FFF9C4').setFontWeight('bold').setFontSize(11);
  sheet.setRowHeight(r, 24); r++;

  sheet.getRange(r, 1).setValue('URL de Acesso').setFontWeight('bold');
  sheet.getRange(r, 2).setValue(url);
  sheet.setRowHeight(r, 22); r += 2;

  // ── Bloco: Como configurar o script ─────────────────────────
  sheet.getRange(r, 1, 1, 2).merge().setValue('⚙️  COMO CONFIGURAR O SCRIPT PARA ESTA PLANILHA')
    .setBackground('#002244').setFontColor('#ffffff').setFontWeight('bold').setFontSize(10);
  sheet.setRowHeight(r, 24); r++;

  var passos = [
    ['Passo 1', 'No formulário original, abra o editor: menu Extensões > Apps Script'],
    ['Passo 2', 'No editor, localize a linha (aprox. linha 5):\n    const SPREADSHEET_ID = \'ID_ANTIGO\';'],
    ['Passo 3', 'Substitua o ID pelo ID desta planilha (célula amarela acima):\n    const SPREADSHEET_ID = \'' + id + '\';'],
    ['Passo 4', 'Salve o projeto (Ctrl+S ou ícone de disquete)'],
    ['Passo 5', 'Clique em "Implantar" > "Gerenciar implantações"'],
    ['Passo 6', 'Clique no lápis (Editar) da implantação ativa'],
    ['Passo 7', 'Em "Versão", selecione "Nova versão" e clique em "Implantar"'],
    ['Passo 8', 'Copie o URL exibido — este é o link do formulário para os candidatos']
  ];
  passos.forEach(function(p) {
    sheet.getRange(r, 1).setValue(p[0]).setFontWeight('bold').setBackground('#E3F2FD');
    sheet.getRange(r, 2).setValue(p[1]).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
    sheet.setRowHeight(r, p[1].indexOf('\n') !== -1 ? 44 : 22); r++;
  });
  r++;

  // ── Bloco: Como encontrar o link após o deploy ───────────────
  sheet.getRange(r, 1, 1, 2).merge().setValue('🔗  COMO ENCONTRAR O LINK DO FORMULÁRIO APÓS REIMPLANTAR')
    .setBackground('#002244').setFontColor('#ffffff').setFontWeight('bold').setFontSize(10);
  sheet.setRowHeight(r, 24); r++;

  var linkPassos = [
    ['1.', 'No GAS: Implantar > Gerenciar implantações'],
    ['2.', 'Localize a implantação do tipo "Aplicativo da Web"'],
    ['3.', 'Clique no ícone de cópia ao lado do URL'],
    ['4.', 'O URL do FORMULÁRIO termina com ?pagina=formulario\n    Ex: https://script.google.com/macros/s/{ID}/exec?pagina=formulario'],
    ['5.', 'O URL do PAINEL DE AVALIAÇÃO termina com ?pagina=painel\n    Ex: https://script.google.com/macros/s/{ID}/exec?pagina=painel']
  ];
  linkPassos.forEach(function(p) {
    sheet.getRange(r, 1).setValue(p[0]).setFontWeight('bold').setBackground('#E8F5E9');
    sheet.getRange(r, 2).setValue(p[1]).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
    sheet.setRowHeight(r, p[1].indexOf('\n') !== -1 ? 44 : 22); r++;
  });
  r++;

  // ── Bloco: Abas da planilha ──────────────────────────────────
  sheet.getRange(r, 1, 1, 2).merge().setValue('📑  ORIENTAÇÕES POR ABA')
    .setBackground('#002244').setFontColor('#ffffff').setFontWeight('bold').setFontSize(10);
  sheet.setRowHeight(r, 24); r++;

  var abas = [
    ['respostas',             'Preenchida AUTOMATICAMENTE pelo formulário. Não editar manualmente.'],
    ['bfi_resultados',        'Resultados do inventário de personalidade BFI-44. Preenchida automaticamente.'],
    ['credenciais',           'Lista de avaliadores e administradores autorizados.\nAdicione e-mails @prf.gov.br com Perfil = AVALIADOR ou ADMINISTRADOR.\nCampo "Unidades Autorizadas": use * para todas, ou informe os nomes separados por vírgula.'],
    ['avaliacoes',            'Avaliações de currículo e entrevista. Preenchida pelo painel de avaliação.'],
    ['graduacao',             'PREENCHER ANTES DE ATIVAR.\nColuna A: Cursos de graduação\nColuna B: Pós-graduações / Especializações\n(Usadas nos autocompletes do formulário)'],
    ['uorgs',                 'PREENCHER ANTES DE ATIVAR.\nColuna A: Siglas/nomes das UORGs da PRF.\n(Usada no autocomplete de lotação de cônjuge e nas unidades indicadas da entrevista)'],
    ['unidades_oportunidade', 'PREENCHER ANTES DE ATIVAR.\nColuna A: Unidades que têm vagas neste recrutamento.\n(Aparece no campo "Unidade de Oportunidade" do formulário)'],
    ['rascunhos_bfi',         'Rascunhos do questionário BFI. Preenchida automaticamente. Não editar.'],
    ['config',                'Configurações do recrutamento. Linha "emailsResposta": e-mails separados por vírgula que receberão notificação quando o cônjuge confirmar a concordância.']
  ];
  abas.forEach(function(a, i) {
    var bg = (i % 2 === 0) ? '#FAFAFA' : '#F0F4F8';
    sheet.getRange(r, 1).setValue(a[0]).setFontWeight('bold').setBackground(bg)
      .setFontFamily('Courier New').setFontSize(10);
    sheet.getRange(r, 2).setValue(a[1]).setBackground(bg)
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
    var linhas = (a[1].match(/\n/g) || []).length;
    sheet.setRowHeight(r, 22 + linhas * 18); r++;
  });
  r++;

  // ── Bloco: Aviso final ───────────────────────────────────────
  sheet.getRange(r, 1, 1, 2).merge()
    .setValue('⚠️  ATENÇÃO: Após configurar o script e implantar, teste o formulário com um e-mail @prf.gov.br antes de divulgar.')
    .setBackground('#FFF3E0').setFontColor('#E65100').setFontWeight('bold')
    .setFontSize(10).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
    .setHorizontalAlignment('center');
  sheet.setRowHeight(r, 36);

  SpreadsheetApp.flush();
}

// ════════════════════════════════════════════════════════════════
//  PAINEL DE AVALIAÇÃO — BACKEND
// ════════════════════════════════════════════════════════════════

const SHEET_CREDENCIAIS   = 'credenciais';
const HEADERS_CREDENCIAIS = ['E-mail','Nome','Perfil','Unidades Autorizadas','Data Cadastro','Status'];
const PERFIL_ADMIN        = 'ADMINISTRADOR';
const PERFIL_AVALIADOR    = 'AVALIADOR';

// ─── Cache de SpreadsheetApp dentro de uma execução ─────────────
// GAS recria o ambiente a cada request, mas dentro de um único request
// evitar openById() repetido reduz latência significativamente.
var _ssCache = null;
function _getSSCached() {
  if (!_ssCache) _ssCache = SpreadsheetApp.openById(SPREADSHEET_ID);
  return _ssCache;
}

// ─── Aba credenciais ─────────────────────────────────────────────
function _getCredenciaisSheet() {
  const ss  = _getSSCached();
  let sheet = ss.getSheetByName(SHEET_CREDENCIAIS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_CREDENCIAIS);
    // Agrupa todas as chamadas de formatação antes do flush para reduzir roundtrips
    const hdr = sheet.getRange(1, 1, 1, HEADERS_CREDENCIAIS.length);
    hdr.setValues([HEADERS_CREDENCIAIS]);
    hdr.setFontWeight('bold');
    hdr.setBackground('#002244');
    hdr.setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    SpreadsheetApp.flush(); // executa as formatações acima em lote
    // Registra o dono do script como primeiro ADMINISTRADOR
    const ownerEmail = Session.getEffectiveUser().getEmail();
    const agora = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm:ss');
    sheet.appendRow([ownerEmail, 'Administrador Principal', PERFIL_ADMIN, '*', agora, 'Ativo']);
  }
  return sheet;
}

// ─── Verificar acesso e retornar perfil ──────────────────────────
function verificarAcessoPainel() {
  const email = Session.getActiveUser().getEmail();
  if (!email) {
    throw new Error('Sessão expirada ou acesso não autorizado. Faça login novamente.');
  }
  if (!email.toLowerCase().endsWith('@prf.gov.br')) {
    throw new Error('Acesso restrito a servidores com e-mail @prf.gov.br. E-mail detectado: ' + email);
  }
  const rows = _getCredenciaisSheet().getDataRange().getValues();
  for (let i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][0]).toLowerCase() === email.toLowerCase()) {
      if (String(rows[i][5]).toLowerCase() !== 'ativo') {
        throw new Error('Seu acesso ao Painel está inativo. Contate o administrador.');
      }
      const unidadesRaw = String(rows[i][3] || '*');
      return {
        email:    String(rows[i][0]),
        nome:     String(rows[i][1] || email),
        perfil:   String(rows[i][2]),
        unidades: unidadesRaw.trim() === '*'
          ? '*'
          : unidadesRaw.split(';').map(function(u){ return u.trim(); }).filter(Boolean)
      };
    }
  }
  throw new Error('Você não tem acesso ao Painel de Avaliação. Solicite acesso ao administrador do sistema.');
}

// ─── Filtrar linhas da planilha por unidades autorizadas ─────────
function _filtrarCandidatos(rows, unidades) {
  return rows.filter(function(r) {
    if (!String(r[1] || '').trim()) return false; // linha sem e-mail
    if (unidades === '*') return true;
    return unidades.indexOf(String(r[5] || '').trim()) !== -1;
  });
}

// ─── Diagnóstico da planilha ─────────────────────────────────────
function diagnosticarPlanilha() {
  verificarAcessoPainel(); // requer acesso
  const ss   = _getSSCached();
  const abas = ss.getSheets().map(function(sh) {
    return {
      nome:   sh.getName(),
      linhas: sh.getLastRow(),
      colunas: sh.getLastColumn(),
      gid:    sh.getSheetId()
    };
  });
  const abaRespostas  = ss.getSheetByName('respostas');
  const cabecalhoReal = abaRespostas
    ? abaRespostas.getRange(1, 1, 1, abaRespostas.getLastColumn()).getValues()[0].join(' | ')
    : null;
  return {
    abas:           abas,
    spreadsheetId:  SPREADSHEET_ID,
    abaEsperada:    'respostas',
    abaEncontrada:  !!abaRespostas,
    totalLinhas:    abaRespostas ? abaRespostas.getLastRow() : 0,
    cabecalhoReal:  cabecalhoReal
  };
}

// ─── Dados do Dashboard ──────────────────────────────────────────
function inicializarPainel() {
  const usuario = verificarAcessoPainel();
  const ss      = _getSSCached();
  const sheet   = ss.getSheetByName('respostas');

  const vazio = {
    total: 0, totalUnidades: 0, comConjuge: 0, posGraduados: 0,
    porUnidade: [], porFormacao: [], statusConjuge: {},
    _sheetEncontrada: !!sheet,
    _totalLinhasBruto: sheet ? sheet.getLastRow() : 0
  };
  if (!sheet || sheet.getLastRow() <= 1) {
    return { usuario: usuario, dashboard: vazio };
  }

  const allRows    = sheet.getDataRange().getValues().slice(1);
  const candidatos = _filtrarCandidatos(allRows, usuario.unidades);

  const unidadesMap = {};
  let comConjuge    = 0;
  let posGraduados  = 0;
  const formCounts  = { graduacao: 0, pos: 0, mestrado: 0, doutorado: 0 };
  const statusConj  = { Confirmado: 0, Pendente: 0, Recusado: 0, 'N/A': 0 };

  candidatos.forEach(function(r) {
    // Por unidade de oportunidade
    const u = String(r[5] || 'Não informado').trim();
    unidadesMap[u] = (unidadesMap[u] || 0) + 1;

    // Cônjuge
    if (String(r[9] || '').toLowerCase() === 'sim') {
      comConjuge++;
      const st = String(r[34] || 'Pendente').trim(); // Status Cônjuge (col 35, idx 34)
      statusConj[st] = (statusConj[st] || 0) + 1;
    } else {
      statusConj['N/A']++;
    }

    // Formação
    if (String(r[10] || '').trim()) formCounts.graduacao++;
    if (String(r[11] || '').trim()) { formCounts.pos++;      posGraduados++; }
    if (String(r[12] || '').trim()) { formCounts.mestrado++; posGraduados++; }
    if (String(r[13] || '').trim()) { formCounts.doutorado++; posGraduados++; }
  });

  const porUnidade = Object.keys(unidadesMap)
    .map(function(k) { return { unidade: k, count: unidadesMap[k] }; })
    .sort(function(a, b) { return b.count - a.count; });

  const porFormacao = [
    { label: 'Graduação',    count: formCounts.graduacao },
    { label: 'Pós-Graduação', count: formCounts.pos      },
    { label: 'Mestrado',     count: formCounts.mestrado  },
    { label: 'Doutorado',    count: formCounts.doutorado }
  ];

  return {
    usuario:   usuario,
    uorgs:     getListasFormacao(ss).uorgs,   // usa CacheService — rápido
    dashboard: {
      total:             candidatos.length,
      totalUnidades:     Object.keys(unidadesMap).length,
      comConjuge:        comConjuge,
      posGraduados:      posGraduados,
      porUnidade:        porUnidade,
      porFormacao:       porFormacao,
      statusConjuge:     statusConj,
      _sheetEncontrada:  true,
      _totalLinhasBruto: sheet.getLastRow()
    }
  };
}

// ─── Listar candidatos (Parte 2) ─────────────────────────────────
function listarCandidatos() {
  const u     = verificarAcessoPainel();
  const ss    = _getSSCached();
  const sheet = ss.getSheetByName('respostas');
  if (!sheet || sheet.getLastRow() <= 1) return [];

  const allRows    = sheet.getDataRange().getValues().slice(1);
  const candidatos = _filtrarCandidatos(allRows, u.unidades);

  return candidatos.map(function(r) {
    const temGrad  = !!String(r[10]||'').trim();
    const temPos   = !!String(r[11]||'').trim();
    const temMest  = !!String(r[12]||'').trim();
    const temDout  = !!String(r[13]||'').trim();
    let formacao   = 'Não informado';
    if (temGrad)  formacao = 'Graduação';
    if (temPos)   formacao = 'Pós-Graduação';
    if (temMest)  formacao = 'Mestrado';
    if (temDout)  formacao = 'Doutorado';

    const temConjuge = String(r[9]||'').toLowerCase() === 'sim';
    return {
      dataHora:     String(r[0]  || ''),
      email:        String(r[1]  || ''),
      nome:         String(r[2]  || ''),
      matricula:    String(r[3]  || ''),
      cargo:        String(r[4]  || ''),
      unidade:      String(r[5]  || ''),
      temConjuge:   temConjuge,
      formacao:     formacao,
      statusConjuge: temConjuge ? String(r[34]||'Pendente').trim() : 'N/A',
      status:       String(r[32] || '')
    };
  }).sort(function(a, b) { return a.nome.localeCompare(b.nome, 'pt-BR'); });
}

// ─── Detalhe de um candidato (Parte 2) ───────────────────────────
function obterCandidato(email) {
  const u     = verificarAcessoPainel();
  const ss    = _getSSCached();
  const sheet = ss.getSheetByName('respostas');
  if (!sheet || sheet.getLastRow() <= 1) throw new Error('Candidato não encontrado.');

  const allRows    = sheet.getDataRange().getValues().slice(1);
  const candidatos = _filtrarCandidatos(allRows, u.unidades);

  for (let i = 0; i < candidatos.length; i++) {
    const r = candidatos[i];
    if (String(r[1]||'').toLowerCase() !== email.toLowerCase()) continue;

    const temConjuge = String(r[9]||'').toLowerCase() === 'sim';
    const ext   = parseFloat(r[26]) || 0;
    const amab  = parseFloat(r[27]) || 0;
    const cons  = parseFloat(r[28]) || 0;
    const estab = parseFloat(r[29]) || 0;
    const abert = parseFloat(r[30]) || 0;

    return {
      dataHora:         String(r[0]  || ''),
      email:            String(r[1]  || ''),
      nome:             String(r[2]  || ''),
      matricula:        String(r[3]  || ''),
      cargo:            String(r[4]  || ''),
      unidade:          String(r[5]  || ''),
      conhecimento:     String(r[6]  || ''),
      ddd:              String(r[7]  || ''),
      telefone:         String(r[8]  || ''),
      temConjuge:       temConjuge,
      // Formação
      graduacoes:       String(r[10] || ''),
      posGraduacoes:    String(r[11] || ''),
      mestrados:        String(r[12] || ''),
      doutorados:       String(r[13] || ''),
      curriculo:        String(r[14] || ''),
      // Big Five (escala 1–5)
      bfi: { ext: ext, amab: amab, cons: cons, estab: estab, abert: abert },
      // Status
      assinatura:       String(r[31] || ''),
      status:           String(r[32] || ''),
      // Cônjuge
      conjugeNome:      String(r[15] || ''),
      conjugeMatricula: String(r[16] || ''),
      conjugeEmail:     String(r[17] || ''),
      tipoUniao:        String(r[18] || ''),
      dataUniao:        String(r[19] || ''),
      endConjuge1:      String(r[20] || ''),
      endConjuge2:      String(r[21] || ''),
      lotConjuge1:      String(r[22] || ''),
      lotConjuge2:      String(r[23] || ''),
      urlUniao:         String(r[24] || ''),
      urlCoabitacao:    String(r[25] || ''),
      statusConjuge:    String(r[34] || ''),
      dataConfirmacao:  String(r[35] || ''),
      // PDFs
      pdfRespostas:     String(r[36] || ''),
      pdfTermo:         String(r[37] || ''),
      pdfConcordancia:  String(r[38] || '')
    };
  }
  throw new Error('Candidato não encontrado ou sem permissão de acesso.');
}

// ─── Gerenciamento de credenciais (apenas ADMINISTRADOR) ──────────
function listarCredenciais() {
  const u = verificarAcessoPainel();
  if (u.perfil !== PERFIL_ADMIN) throw new Error('Acesso restrito a administradores.');
  const sheet = _getCredenciaisSheet();
  if (sheet.getLastRow() <= 1) return []; // apenas cabeçalho
  return sheet.getDataRange().getValues().slice(1)
    .filter(function(r) { return !!String(r[0] || '').trim(); }) // ignora linhas em branco
    .map(function(r) {
      // Normaliza dataCadastro para string (evita problemas de serialização de Date)
      var dt = r[4];
      if (dt instanceof Date) {
        dt = Utilities.formatDate(dt, 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm:ss');
      } else {
        dt = String(dt || '');
      }
      return {
        email:       String(r[0] || ''),
        nome:        String(r[1] || ''),
        perfil:      String(r[2] || ''),
        unidades:    String(r[3] || ''),
        dataCadastro: dt,
        status:      String(r[5] || 'Ativo')
      };
    });
}

function salvarCredencial(dados) {
  const u = verificarAcessoPainel();
  if (u.perfil !== PERFIL_ADMIN) throw new Error('Acesso restrito a administradores.');
  if (!dados.email || !dados.email.toLowerCase().endsWith('@prf.gov.br')) {
    throw new Error('O e-mail deve ser @prf.gov.br.');
  }
  if (dados.perfil !== PERFIL_ADMIN && dados.perfil !== PERFIL_AVALIADOR) {
    throw new Error('Perfil inválido. Use ADMINISTRADOR ou AVALIADOR.');
  }
  const sheet    = _getCredenciaisSheet();
  const rows     = sheet.getDataRange().getValues();
  const agora    = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm:ss');
  const unidades = (dados.perfil === PERFIL_ADMIN) ? '*' : (dados.unidades || '');
  const status   = dados.status || 'Ativo';

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).toLowerCase() === dados.email.toLowerCase()) {
      sheet.getRange(i + 1, 1, 1, 6).setValues([[
        dados.email, dados.nome || '', dados.perfil, unidades, rows[i][4] || agora, status
      ]]);
      return 'atualizado';
    }
  }
  sheet.appendRow([dados.email, dados.nome || '', dados.perfil, unidades, agora, status]);
  return 'criado';
}

function revogarCredencial(email) {
  const u = verificarAcessoPainel();
  if (u.perfil !== PERFIL_ADMIN) throw new Error('Acesso restrito a administradores.');
  if (email.toLowerCase() === u.email.toLowerCase()) {
    throw new Error('Você não pode revogar seu próprio acesso.');
  }
  const sheet = _getCredenciaisSheet();
  const rows  = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).toLowerCase() === email.toLowerCase()) {
      sheet.getRange(i + 1, 6).setValue('Inativo');
      return 'ok';
    }
  }
  throw new Error('Usuário não encontrado.');
}

// ════════════════════════════════════════════════════════════════
//  PARTE 3 — AVALIAÇÕES (Currículo e Entrevista)
// ════════════════════════════════════════════════════════════════

const SHEET_AVALIACOES    = 'avaliacoes';
const HEADERS_AVALIACOES  = [
  'E-mail Candidato', 'Nome Candidato', 'Unidade Candidato',
  'E-mail Avaliador',  'Nome Avaliador', 'Tipo',
  'Data/Hora',
  'Nota Geral',
  'Nota Comunicação', 'Nota Int. Emocional', 'Nota Postura', 'Nota Ling. Corporal',
  'Unidade Indicada',
  'Obs Positivas', 'Obs Negativas', 'Obs Livres',
  'Status'
];

function _getAvaliacoesSheet() {
  const ss  = _getSSCached();
  let sheet = ss.getSheetByName(SHEET_AVALIACOES);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_AVALIACOES);
    const hdr = sheet.getRange(1, 1, 1, HEADERS_AVALIACOES.length);
    hdr.setValues([HEADERS_AVALIACOES]);
    hdr.setFontWeight('bold');
    hdr.setBackground('#002244');
    hdr.setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    sheet.setColumnWidths(1, HEADERS_AVALIACOES.length, 160);
  }
  return sheet;
}

// Retorna lista de candidatos autorizados já enriquecida com a avaliação
// do usuário atual (se existir) para o tipo dado ('curriculo' | 'entrevista').
function listarParaAvaliar(tipo) {
  const u = verificarAcessoPainel();

  // ── Candidatos ───────────────────────────────────────────────
  const ss       = _getSSCached();
  const sheetR   = ss.getSheetByName('respostas');
  if (!sheetR || sheetR.getLastRow() <= 1) return [];
  const allRows  = sheetR.getDataRange().getValues().slice(1);
  const cands    = _filtrarCandidatos(allRows, u.unidades);

  // ── Avaliações (minha + contagem de concluídas de todos) ─────
  const sheetA  = _getAvaliacoesSheet();
  const meusMap  = {};   // emailCandidato → minha avaliação
  const conclMap = {};   // emailCandidato → {emailAvaliador: true} p/ avaliações Concluídas
  if (sheetA.getLastRow() > 1) {
    sheetA.getDataRange().getValues().slice(1)
      .filter(function(r) { return String(r[5]).toLowerCase() === tipo.toLowerCase(); })
      .forEach(function(r) {
        const ec = String(r[0]).toLowerCase();
        const ea = String(r[3]).toLowerCase();
        const st = String(r[16] || '');

        // Contagem de avaliadores distintos que concluíram
        if (st.toLowerCase() === 'concluída') {
          if (!conclMap[ec]) conclMap[ec] = {};
          conclMap[ec][ea] = true;
        }

        // Minha avaliação
        if (ea === u.email.toLowerCase()) {
          meusMap[ec] = {
            notaGeral:        r[7]  !== '' ? Number(r[7])  : 0,
            notaComunicacao:  r[8]  !== '' ? Number(r[8])  : 0,
            notaIntEmocional: r[9]  !== '' ? Number(r[9])  : 0,
            notaPostura:      r[10] !== '' ? Number(r[10]) : 0,
            notaLingCorporal: r[11] !== '' ? Number(r[11]) : 0,
            unidadeIndicada:  String(r[12] || ''),
            obsPositivas:     String(r[13] || ''),
            obsNegativas:     String(r[14] || ''),
            obsLivres:        String(r[15] || ''),
            status:           st || 'Rascunho'
          };
        }
      });
  }

  // ── Merge ────────────────────────────────────────────────────
  return cands.map(function(r) {
    const email = String(r[1] || '').toLowerCase();
    const aval  = meusMap[email] || null;

    // Nota para exibição na tabela
    var notaExibir = null;
    if (aval) {
      if (tipo === 'curriculo') {
        notaExibir = aval.notaGeral || null;
      } else {
        var ns = [aval.notaComunicacao, aval.notaIntEmocional, aval.notaPostura, aval.notaLingCorporal]
          .filter(function(n){ return n > 0; });
        notaExibir = ns.length > 0 ? +(ns.reduce(function(a,b){return a+b;},0) / ns.length).toFixed(1) : null;
      }
    }

    return {
      email:          String(r[1] || ''),
      nome:           String(r[2] || ''),
      matricula:      String(r[3] || ''),
      cargo:          String(r[4] || ''),
      unidade:        String(r[5] || ''),
      dataHora:       String(r[0] || ''),
      sougovUrl:      String(r[14] || ''),
      avalStatus:     aval ? aval.status : 'Não avaliado',
      notaExibir:     notaExibir,
      avaliacao:      aval,
      avalConcluidas: conclMap[email] ? Object.keys(conclMap[email]).length : 0
    };
  }).sort(function(a, b){ return a.nome.localeCompare(b.nome, 'pt-BR'); });
}

// ─── Resumo Geral — consolida avaliações de todos os avaliadores ───
function listarResumo() {
  const u      = verificarAcessoPainel();
  const ss     = _getSSCached();
  const sheetR = ss.getSheetByName('respostas');
  if (!sheetR || sheetR.getLastRow() <= 1) return [];

  const allRows = sheetR.getDataRange().getValues().slice(1);
  const cands   = _filtrarCandidatos(allRows, u.unidades);

  // Mapa de candidatos (keyed by email lowercase)
  const candMap = {};
  cands.forEach(function(r) {
    const email = String(r[1] || '').toLowerCase();
    if (!email) return;
    candMap[email] = {
      email:     String(r[1] || ''),
      nome:      String(r[2] || ''),
      matricula: String(r[3] || ''),
      cargo:     String(r[4] || ''),
      unidade:   String(r[5] || ''),
      sougovUrl: String(r[14] || ''),
      _cur: { notas: [], avaliacoes: [] },
      _ent: { criterios: [], unidades: {}, avaliacoes: [] }
    };
  });

  // Percorre avaliacoes em uma única leitura
  const sheetA = _getAvaliacoesSheet();
  if (sheetA.getLastRow() > 1) {
    sheetA.getDataRange().getValues().slice(1).forEach(function(r) {
      const ec = String(r[0] || '').toLowerCase();
      if (!candMap[ec]) return;

      const tipo  = String(r[5] || '').toLowerCase();
      const st    = String(r[16] || '');
      const isConc = st.toLowerCase() === 'concluída';
      const dt    = r[6] instanceof Date
        ? Utilities.formatDate(r[6], 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm')
        : String(r[6] || '');

      if (tipo === 'curriculo') {
        const nota = r[7] !== '' ? Number(r[7]) : 0;
        candMap[ec]._cur.avaliacoes.push({
          nomeAvaliador: String(r[4] || ''),
          nota:          nota,
          obsPositivas:  String(r[13] || ''),
          obsNegativas:  String(r[14] || ''),
          obsLivres:     String(r[15] || ''),
          status:        st,
          dataHora:      dt
        });
        if (isConc && nota > 0) candMap[ec]._cur.notas.push(nota);

      } else if (tipo === 'entrevista') {
        const nc  = r[8]  !== '' ? Number(r[8])  : 0;
        const nie = r[9]  !== '' ? Number(r[9])  : 0;
        const np  = r[10] !== '' ? Number(r[10]) : 0;
        const nlc = r[11] !== '' ? Number(r[11]) : 0;
        const uni = String(r[12] || '');
        candMap[ec]._ent.avaliacoes.push({
          nomeAvaliador:    String(r[4] || ''),
          notaComunicacao:  nc,
          notaIntEmocional: nie,
          notaPostura:      np,
          notaLingCorporal: nlc,
          unidadeIndicada:  uni,
          obsLivres:        String(r[15] || ''),
          status:           st,
          dataHora:         dt
        });
        if (isConc) {
          if (nc > 0 || nie > 0 || np > 0 || nlc > 0) {
            candMap[ec]._ent.criterios.push({ nc: nc, nie: nie, np: np, nlc: nlc });
          }
          if (uni) {
            uni.split(';').map(function(s){ return s.trim(); }).filter(Boolean)
              .forEach(function(uStr) {
                candMap[ec]._ent.unidades[uStr] = (candMap[ec]._ent.unidades[uStr] || 0) + 1;
              });
          }
        }
      }
    });
  }

  // Calcula médias e monta objeto limpo para o cliente
  return Object.values(candMap).map(function(c) {
    const cur = c._cur;
    const ent = c._ent;

    const mediaCur = cur.notas.length > 0
      ? +(cur.notas.reduce(function(a,b){return a+b;},0) / cur.notas.length).toFixed(1)
      : null;

    var mediaComun = null, mediaIntE = null, mediaPost = null, mediaLing = null, mediaEnt = null;
    if (ent.criterios.length > 0) {
      var len = ent.criterios.length;
      mediaComun = +(ent.criterios.reduce(function(a,b){return a+b.nc; },0) / len).toFixed(1);
      mediaIntE  = +(ent.criterios.reduce(function(a,b){return a+b.nie;},0) / len).toFixed(1);
      mediaPost  = +(ent.criterios.reduce(function(a,b){return a+b.np; },0) / len).toFixed(1);
      mediaLing  = +(ent.criterios.reduce(function(a,b){return a+b.nlc;},0) / len).toFixed(1);
      mediaEnt   = +((Number(mediaComun)+Number(mediaIntE)+Number(mediaPost)+Number(mediaLing))/4).toFixed(1);
    }

    // Contagem de avaliadores que concluíram (por tipo)
    const curConc = cur.avaliacoes.filter(function(a){ return a.status.toLowerCase() === 'concluída'; }).length;
    const entConc = ent.avaliacoes.filter(function(a){ return a.status.toLowerCase() === 'concluída'; }).length;

    return {
      email:     c.email,
      nome:      c.nome,
      matricula: c.matricula,
      cargo:     c.cargo,
      unidade:   c.unidade,
      sougovUrl: c.sougovUrl,
      curriculo: {
        concluidas: curConc,
        mediaNota:  mediaCur,
        avaliacoes: cur.avaliacoes
      },
      entrevista: {
        concluidas:        entConc,
        mediaGeral:        mediaEnt,
        mediaComunicacao:  mediaComun,
        mediaIntEmocional: mediaIntE,
        mediaPostura:      mediaPost,
        mediaLingCorporal: mediaLing,
        unidadesIndicadas: Object.keys(ent.unidades)
          .map(function(k){ return { unidade: k, votos: ent.unidades[k] }; })
          .sort(function(a,b){ return b.votos - a.votos; }),
        avaliacoes: ent.avaliacoes
      }
    };
  }).sort(function(a, b){ return a.nome.localeCompare(b.nome, 'pt-BR'); });
}

function salvarAvaliacao(dados) {
  const u = verificarAcessoPainel();
  if (!dados.emailCandidato) throw new Error('E-mail do candidato não informado.');
  const tipo = String(dados.tipo || '').toLowerCase();
  if (tipo !== 'curriculo' && tipo !== 'entrevista') throw new Error('Tipo inválido.');

  const sheet = _getAvaliacoesSheet();
  const agora = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm:ss');

  const rowData = [
    dados.emailCandidato,
    dados.nomeCandidato      || '',
    dados.unidadeCandidato   || '',
    u.email,
    u.nome,
    tipo,
    agora,
    dados.notaGeral          || '',
    dados.notaComunicacao    || '',
    dados.notaIntEmocional   || '',
    dados.notaPostura        || '',
    dados.notaLingCorporal   || '',
    dados.unidadeIndicada    || '',
    dados.obsPositivas       || '',
    dados.obsNegativas       || '',
    dados.obsLivres          || '',
    dados.status             || 'Rascunho'
  ];

  // Atualizar linha existente (mesmo candidato + avaliador + tipo)
  if (sheet.getLastRow() > 1) {
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).toLowerCase() === dados.emailCandidato.toLowerCase()
       && String(rows[i][3]).toLowerCase() === u.email.toLowerCase()
       && String(rows[i][5]).toLowerCase() === tipo) {
        sheet.getRange(i + 1, 1, 1, rowData.length).setValues([rowData]);
        return 'atualizado';
      }
    }
  }
  sheet.appendRow(rowData);
  return 'criado';
}

function excluirCredencial(email) {
  const u = verificarAcessoPainel();
  if (u.perfil !== PERFIL_ADMIN) throw new Error('Acesso restrito a administradores.');
  if (email.toLowerCase() === u.email.toLowerCase()) {
    throw new Error('Você não pode excluir seu próprio cadastro.');
  }
  const sheet = _getCredenciaisSheet();
  const rows  = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).toLowerCase() === email.toLowerCase()) {
      sheet.deleteRow(i + 1);
      return 'ok';
    }
  }
  throw new Error('Usuário não encontrado.');
}

// ════════════════════════════════════════════════════════════════
//  VOTAÇÃO FINAL
// ════════════════════════════════════════════════════════════════

// ─── Verifica se o usuário tem perfil VOTADOR ou ADMINISTRADOR ────
function _verificarAcessoVotacao() {
  const email = Session.getActiveUser().getEmail();
  if (!email || !email.toLowerCase().endsWith('@prf.gov.br')) {
    throw new Error('Acesso restrito a servidores @prf.gov.br.');
  }
  const rows = _getCredenciaisSheet().getDataRange().getValues();
  for (let i = rows.length - 1; i >= 1; i--) {
    const e = String(rows[i][0] || '').toLowerCase();
    const p = String(rows[i][2] || '').toUpperCase();
    const s = String(rows[i][5] || '').toLowerCase();
    if (e === email.toLowerCase() && s === 'ativo'
        && (p === 'VOTADOR' || p === 'ADMINISTRADOR')) {
      return { email, nome: String(rows[i][1] || email), perfil: p };
    }
  }
  throw new Error('Você não tem permissão para acessar a votação. Solicite acesso ao administrador.');
}

// ─── Inicializa a tela de votação ─────────────────────────────────
function inicializarVotacao() {
  const usuario = _verificarAcessoVotacao();
  const ss      = SpreadsheetApp.openById(SPREADSHEET_ID);

  // Carrega candidatos com status "Inscrito"
  const sheetResp = ss.getSheetByName('respostas');
  const candidatos = [];
  if (sheetResp && sheetResp.getLastRow() > 1) {
    const lastRow = sheetResp.getLastRow();
    const dados   = sheetResp.getRange(1, 1, lastRow, 33).getValues().slice(1);
    dados.forEach(function(r) {
      if (String(r[32] || '').toLowerCase() === 'inscrito' && String(r[1] || '').trim()) {
        candidatos.push({
          email:     String(r[1] || ''),
          nome:      String(r[2] || ''),
          matricula: String(r[3] || ''),
          unidade:   String(r[5] || '')
        });
      }
    });
  }

  // Verifica quais candidatos este eleitor já votou
  const jaVotei = [];
  const sheetVotos = ss.getSheetByName(SHEET_VOTOS);
  if (sheetVotos && sheetVotos.getLastRow() > 1) {
    sheetVotos.getDataRange().getValues().slice(1).forEach(function(r) {
      if (String(r[1] || '').toLowerCase() === usuario.email.toLowerCase()) {
        jaVotei.push(String(r[3] || '').toLowerCase());
      }
    });
  }

  return { eleitor: usuario.nome || usuario.email, candidatos, jaVotei };
}

// ─── Salva o voto de um eleitor para um candidato ─────────────────
function salvarVoto(emailCandidato, notas, comentario) {
  const usuario = _verificarAcessoVotacao();
  const ss      = SpreadsheetApp.openById(SPREADSHEET_ID);

  // Garante a aba "votos"
  let sheetVotos = ss.getSheetByName(SHEET_VOTOS);
  if (!sheetVotos) {
    sheetVotos = ss.insertSheet(SHEET_VOTOS);
    sheetVotos.appendRow(HEADERS_VOTOS);
    sheetVotos.setFrozenRows(1);
    sheetVotos.getRange('A1:M1')
      .setFontWeight('bold')
      .setBackground('#1d1a5b')
      .setFontColor('#ffffff');
  }

  // Verifica duplicidade
  if (sheetVotos.getLastRow() > 1) {
    const rows = sheetVotos.getDataRange().getValues().slice(1);
    for (const r of rows) {
      if (String(r[1] || '').toLowerCase() === usuario.email.toLowerCase() &&
          String(r[3] || '').toLowerCase() === emailCandidato.toLowerCase()) {
        throw new Error('Você já registrou seu voto para este candidato.');
      }
    }
  }

  // Busca dados do candidato na aba "respostas"
  let nomeC = '', matriculaC = '';
  const sheetResp = ss.getSheetByName('respostas');
  if (sheetResp && sheetResp.getLastRow() > 1) {
    const rows = sheetResp.getRange(1, 1, sheetResp.getLastRow(), 4).getValues().slice(1);
    for (const r of rows) {
      if (String(r[1] || '').toLowerCase() === emailCandidato.toLowerCase()) {
        nomeC      = String(r[2] || '');
        matriculaC = String(r[3] || '');
        break;
      }
    }
  }

  // Calcula média
  const soma  = (notas.confianca || 0) + (notas.lealdade || 0) + (notas.amizade || 0)
              + (notas.ego || 0)       + (notas.familia  || 0);
  const media = parseFloat((soma / 5).toFixed(2));

  sheetVotos.appendRow([
    new Date(),
    usuario.email,
    usuario.nome || usuario.email,
    emailCandidato,
    nomeC,
    matriculaC,
    notas.confianca  || 0,
    notas.lealdade   || 0,
    notas.amizade    || 0,
    notas.ego        || 0,
    notas.familia    || 0,
    media,
    comentario       || ''
  ]);

  return { media };
}

// ════════════════════════════════════════════════════════════════
//  ACOMPANHAMENTO DO CANDIDATO
// ════════════════════════════════════════════════════════════════

// ─── Constantes da aba de entrevistas ────────────────────────────
const SHEET_ENTREVISTAS  = 'entrevistas';
const HEADERS_ENTREVISTAS = [
  'Email Candidato','Nome Candidato','Matrícula',
  'Email Entrevistador 1','Nome Entrevistador 1',
  'Email Entrevistador 2','Nome Entrevistador 2',
  'Email Entrevistador 3','Nome Entrevistador 3',
  'Data Entrevista','Hora Início','Duração (min)','Link Meet',
  'Status'
];

const SHEET_SOLICITACOES  = 'solicitacoes_alteracao';
const HEADERS_SOLICITACOES = [
  'Data/Hora','Email Candidato','Nome Candidato',
  'Email Entrevistador','Nome Entrevistador',
  'Justificativa','Status'
];

// ─── Inicializa página de acompanhamento ─────────────────────────
function inicializarAcompanhamento() {
  const email = Session.getActiveUser().getEmail();
  if (!email) throw new Error('Não foi possível identificar o usuário.');

  const ss        = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheetResp = ss.getSheetByName('respostas');
  if (!sheetResp || sheetResp.getLastRow() < 2) {
    throw new Error('Nenhuma inscrição encontrada para este e-mail.');
  }

  // Busca inscrição do candidato (colunas 0-32 para ter Status)
  const lastRow = sheetResp.getLastRow();
  const data    = sheetResp.getRange(1, 1, lastRow, 33).getValues();
  let rowCandidato = null;
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][1] || '').toLowerCase() === email.toLowerCase()) {
      rowCandidato = data[i];
      break;
    }
  }
  if (!rowCandidato) throw new Error('Nenhuma inscrição encontrada para este e-mail.');

  const status = String(rowCandidato[32] || '');

  // Busca dados da entrevista do candidato
  let entrevista = null;
  const sheetEnt = ss.getSheetByName(SHEET_ENTREVISTAS);
  if (sheetEnt && sheetEnt.getLastRow() > 1) {
    const entRows = sheetEnt.getDataRange().getValues().slice(1);
    for (const r of entRows) {
      if (String(r[0] || '').toLowerCase() === email.toLowerCase()) {
        entrevista = {
          data:           String(r[9]  || ''),
          hora:           String(r[10] || ''),
          duracao:        String(r[11] || ''),
          linkMeet:       String(r[12] || ''),
          statusEnt:      String(r[13] || ''),
          entrevistadores: [
            r[3] ? { email: String(r[3]), nome: String(r[4] || '') } : null,
            r[5] ? { email: String(r[5]), nome: String(r[6] || '') } : null,
            r[7] ? { email: String(r[7]), nome: String(r[8] || '') } : null
          ].filter(Boolean)
        };
        break;
      }
    }
  }

  // Busca solicitações anteriores deste candidato
  let solicitacoes = [];
  const sheetSol = ss.getSheetByName(SHEET_SOLICITACOES);
  if (sheetSol && sheetSol.getLastRow() > 1) {
    sheetSol.getDataRange().getValues().slice(1).forEach(function(r) {
      if (String(r[1] || '').toLowerCase() === email.toLowerCase()) {
        solicitacoes.push({
          data:              String(r[0] || ''),
          emailEntrevistador: String(r[3] || ''),
          nomeEntrevistador:  String(r[4] || ''),
          justificativa:      String(r[5] || ''),
          statusSol:          String(r[6] || '')
        });
      }
    });
  }

  return {
    nome:        String(rowCandidato[2]  || ''),
    matricula:   String(rowCandidato[3]  || ''),
    unidade:     String(rowCandidato[5]  || ''),
    status:      status,
    entrevista:  entrevista,
    solicitacoes: solicitacoes
  };
}

// ─── Solicita alteração de entrevistador ─────────────────────────
function solicitarAlteracaoEntrevistador(emailEntrevistador, nomeEntrevistador, justificativa) {
  const email = Session.getActiveUser().getEmail();
  if (!email) throw new Error('Não foi possível identificar o usuário.');
  if (!emailEntrevistador || !justificativa || justificativa.trim().length < 20) {
    throw new Error('Preencha o entrevistador e a justificativa (mínimo 20 caracteres).');
  }

  const ss        = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheetResp = ss.getSheetByName('respostas');
  let nomeCandidato = '';
  if (sheetResp && sheetResp.getLastRow() > 1) {
    const rows = sheetResp.getRange(1, 1, sheetResp.getLastRow(), 3).getValues().slice(1);
    for (const r of rows) {
      if (String(r[1] || '').toLowerCase() === email.toLowerCase()) {
        nomeCandidato = String(r[2] || '');
        break;
      }
    }
  }

  // Verifica se já existe solicitação pendente para o mesmo entrevistador
  let sheetSol = ss.getSheetByName(SHEET_SOLICITACOES);
  if (!sheetSol) {
    sheetSol = ss.insertSheet(SHEET_SOLICITACOES);
    sheetSol.appendRow(HEADERS_SOLICITACOES);
    sheetSol.setFrozenRows(1);
    sheetSol.getRange(1, 1, 1, HEADERS_SOLICITACOES.length)
      .setFontWeight('bold').setBackground('#1d1a5b').setFontColor('#ffffff');
  } else if (sheetSol.getLastRow() > 1) {
    const rows = sheetSol.getDataRange().getValues().slice(1);
    for (const r of rows) {
      if (String(r[1] || '').toLowerCase() === email.toLowerCase() &&
          String(r[3] || '').toLowerCase() === emailEntrevistador.toLowerCase() &&
          String(r[6] || '').toLowerCase() === 'pendente') {
        throw new Error('Você já possui uma solicitação pendente para este entrevistador.');
      }
    }
  }

  sheetSol.appendRow([
    new Date(),
    email,
    nomeCandidato,
    emailEntrevistador,
    nomeEntrevistador || emailEntrevistador,
    justificativa.trim(),
    'Pendente'
  ]);

  return { ok: true };
}
