// ============================================================
// PREMIER LOGISTICS — GOOGLE APPS SCRIPT BACKEND
// Cole este código no Google Apps Script Editor (script.google.com)
// Siga o SETUP_GUIDE.md para configuração completa
// ============================================================

// ⚠️ CONFIGURE ESTAS CONSTANTES ANTES DE PUBLICAR:
const CONFIG = {
  SPREADSHEET_ID:  '1sU9aEz5p0JDhjSEowJjTq69Qur6rnP3JdWVmf7vCp68',
  FORM_URL:        'https://ntipremierlog-sys.github.io/Ficha-de-Cadastro_PremierLog/',
  ADMIN_SECRET:    'Premier2025AdminSecret!',
  EMAIL_FROM_NAME: 'RH — Premier Logistics',
  EMAIL_SUBJECT:   'Premier Logistics — Ficha de Cadastro de Admissão'
};

const SHEET_CANDIDATES = 'Candidatos';
const SHEET_RESPONSES  = 'Respostas';

// ============================================================
// ROTEADOR GET
// ============================================================
function doGet(e) {
  const action = (e.parameter.action || '').trim();
  try {
    switch (action) {
      case 'ping':            return jsonResponse({ ok: true, ts: new Date().toISOString() });
      case 'validateToken':   return jsonResponse(validateToken(e.parameter.token));
      case 'getCandidates':   return jsonResponse(getCandidates(e.parameter.secret));
      case 'getResponse':     return jsonResponse(getResponse(e.parameter.token, e.parameter.secret));
      case 'getResponses':    return jsonResponse(getAllResponses(e.parameter.secret));
      case 'addCandidate':    return jsonResponse(addCandidate(e.parameter));
      case 'resendEmail':     return jsonResponse(resendEmail(e.parameter.token, e.parameter.secret));
      case 'deleteCandidate': return jsonResponse(deleteCandidate(e.parameter.token, e.parameter.secret));
      default:                return jsonResponse({ error: 'Ação inválida: ' + action });
    }
  } catch (err) {
    console.error('doGet error:', err);
    return jsonResponse({ error: err.message });
  }
}

// ============================================================
// ROTEADOR POST (form submission)
// ============================================================
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.action === 'submitForm') return jsonResponse(submitForm(data));
    return jsonResponse({ error: 'Ação POST inválida' });
  } catch (err) {
    console.error('doPost error:', err);
    return jsonResponse({ error: err.message });
  }
}

// ============================================================
// FUNÇÃO: Adicionar Candidato
// ============================================================
function addCandidate(params) {
  if (params.secret !== CONFIG.ADMIN_SECRET) return { error: 'Não autorizado' };

  const nome      = (params.nome  || '').trim();
  const email     = (params.email || '').trim();
  const vaga      = (params.vaga  || '').trim();
  const sendEmail = params.sendEmail === 'true';

  if (!nome || !email || !vaga) return { error: 'Nome, e-mail e vaga são obrigatórios.' };

  ensureSheets();

  const ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_CANDIDATES);
  const token = generateToken();
  const now   = formatDate(new Date());
  const id    = Date.now().toString();

  sheet.appendRow([id, nome, email, vaga, token, 'Pendente', now, '']);

  const formUrl = CONFIG.FORM_URL + '?token=' + token;
  let emailSent = false;
  let emailError = null;

  if (sendEmail && email) {
    try {
      sendCandidateEmail(nome, email, token, vaga, formUrl);
      emailSent = true;
    } catch (err) {
      emailError = err.message;
    }
  }

  return { success: true, token: token, formUrl: formUrl, id: id, emailSent: emailSent, emailError: emailError };
}

// ============================================================
// FUNÇÃO: Reenviar E-mail
// ============================================================
function resendEmail(token, secret) {
  if (secret !== CONFIG.ADMIN_SECRET) return { error: 'Não autorizado' };
  if (!token) return { error: 'Token obrigatório' };

  const ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_CANDIDATES);
  const data  = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (data[i][4] === token) {
      var nome    = data[i][1];
      var email   = data[i][2];
      var vaga    = data[i][3];
      var formUrl = CONFIG.FORM_URL + '?token=' + token;
      sendCandidateEmail(nome, email, token, vaga, formUrl);
      return { success: true };
    }
  }
  return { error: 'Candidato não encontrado' };
}

// ============================================================
// FUNÇÃO: Validar Token (candidato)
// ============================================================
function validateToken(token) {
  if (!token) return { valid: false, reason: 'no_token', message: 'Token não fornecido.' };
  ensureSheets();

  var ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEET_CANDIDATES);
  var data  = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (data[i][4] === token) {
      if (data[i][5] === 'Concluído') {
        return { valid: false, reason: 'already_submitted', message: 'Este formulário já foi preenchido.' };
      }
      return { valid: true, nome: data[i][1], vaga: data[i][3], token: token };
    }
  }
  return { valid: false, reason: 'not_found', message: 'Link inválido ou expirado.' };
}

// ============================================================
// FUNÇÃO: Listar Candidatos (admin)
// ============================================================
function getCandidates(secret) {
  if (secret !== CONFIG.ADMIN_SECRET) return { error: 'Não autorizado' };
  ensureSheets();

  var ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEET_CANDIDATES);
  var data  = sheet.getDataRange().getValues();
  var candidates = [];

  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    candidates.push({
      id:                data[i][0] || '',
      nome:              data[i][1] || '',
      email:             data[i][2] || '',
      vaga:              data[i][3] || '',
      token:             data[i][4] || '',
      status:            data[i][5] || 'Pendente',
      dataEnvio:         data[i][6] || '',
      dataPreenchimento: data[i][7] || ''
    });
  }

  return { success: true, candidates: candidates.reverse() };
}

// ============================================================
// FUNÇÃO: Obter Resposta de um Candidato (admin)
// ============================================================
function getResponse(token, secret) {
  if (secret !== CONFIG.ADMIN_SECRET) return { error: 'Não autorizado' };
  if (!token) return { error: 'Token obrigatório' };
  ensureSheets();

  var ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEET_RESPONSES);
  var data  = sheet.getDataRange().getValues();
  if (data.length <= 1) return { error: 'Nenhuma resposta encontrada' };

  var headers = data[0];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === token) {
      var response = {};
      headers.forEach(function(h, j) { response[h] = data[i][j] || ''; });
      return { success: true, response: response };
    }
  }
  return { error: 'Resposta não encontrada para este candidato' };
}

// ============================================================
// FUNÇÃO: Todas as Respostas (admin — export Excel)
// ============================================================
function getAllResponses(secret) {
  if (secret !== CONFIG.ADMIN_SECRET) return { error: 'Não autorizado' };
  ensureSheets();

  var ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEET_RESPONSES);
  var data  = sheet.getDataRange().getValues();
  if (data.length <= 1) return { success: true, responses: [] };

  var headers   = data[0];
  var responses = [];
  for (var i = 1; i < data.length; i++) {
    var row = {};
    headers.forEach(function(h, j) { row[h] = data[i][j] || ''; });
    responses.push(row);
  }
  return { success: true, responses: responses };
}

// ============================================================
// FUNÇÃO: Salvar Formulário (candidato)
// ============================================================
function submitForm(data) {
  var token = data.token;
  if (!token) return { error: 'Token não fornecido' };
  ensureSheets();

  var ss      = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var cSheet  = ss.getSheetByName(SHEET_CANDIDATES);
  var cData   = cSheet.getDataRange().getValues();
  var now     = formatDate(new Date());
  var candidateName = '';

  for (var i = 1; i < cData.length; i++) {
    if (cData[i][4] === token) {
      if (cData[i][5] === 'Concluído') return { error: 'Formulário já preenchido anteriormente.' };
      cSheet.getRange(i + 1, 6).setValue('Concluído');
      cSheet.getRange(i + 1, 8).setValue(now);
      candidateName = cData[i][1];
      break;
    }
  }

  var rSheet = ss.getSheetByName(SHEET_RESPONSES);
  var fd     = data.formData || {};
  var depIR  = (fd.dependentesIR || []).map(function(d) {
    return (d.nome || '') + ' | CPF: ' + (d.cpf || '') + ' | Parentesco: ' + (d.parentesco || '');
  }).join(' ;; ');

  rSheet.appendRow([
    token, now, candidateName,
    fd.nomeCompleto || '', fd.nomeSocial || '',
    fd.cpf || '', fd.rg || '',
    fd.endereco || '', fd.bairroCidade || '', fd.cep || '',
    fd.whatsapp || '', fd.email || '',
    fd.contatoEmergenciaNome || '', fd.contatoEmergenciaTelefone || '',
    fd.tituloEleitor || '', fd.grauInstrucao || '',
    fd.possuiFilhos || '', fd.quantosFilhos || '',
    fd.declararDependenteIR || '', fd.quantosDependentesIR || '', depIR,
    fd.estadoCivil || '', fd.estadoCivilOutro || '',
    fd.botaNumero || '', fd.camisaTamanho || '', fd.calcaTamanho || '',
    fd.optanteVT || '', fd.planoSaudeOpcao || '',
    fd.dependente1Nome || '', fd.dependente1Cpf || '',
    fd.dependente2Nome || '', fd.dependente2Cpf || '',
    fd.tipoAssinatura || ''
  ]);

  return { success: true, message: 'Formulário recebido com sucesso!' };
}

// ============================================================
// FUNÇÃO: Excluir Candidato (admin)
// ============================================================
function deleteCandidate(token, secret) {
  if (secret !== CONFIG.ADMIN_SECRET) return { error: 'Não autorizado' };
  var ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEET_CANDIDATES);
  var data  = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (data[i][4] === token) { sheet.deleteRow(i + 1); return { success: true }; }
  }
  return { error: 'Candidato não encontrado' };
}

// ============================================================
// AUXILIARES
// ============================================================
function generateToken() {
  var chars = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghjkmnpqrstuvwxyz23456789';
  var t = '';
  for (var i = 0; i < 24; i++) t += chars[Math.floor(Math.random() * chars.length)];
  return t;
}

function formatDate(d) {
  return Utilities.formatDate(d, 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm');
}

function sendCandidateEmail(nome, email, token, vaga, formUrl) {
  var html = '<!DOCTYPE html><html><body style="margin:0;padding:0;background:#f1f5f9;font-family:Arial,sans-serif;">' +
    '<table width="100%" cellpadding="0" cellspacing="0"><tr><td align="center" style="padding:40px 16px;">' +
    '<table width="600" style="background:white;border-radius:12px;overflow:hidden;box-shadow:0 4px 24px rgba(0,0,0,.10);">' +
    '<tr><td style="background:#211551;padding:32px;text-align:center;">' +
    '<h1 style="color:#D4AF37;margin:0;font-size:24px;letter-spacing:1px;">PREMIER LOGISTICS</h1>' +
    '<p style="color:rgba(255,255,255,.75);margin:6px 0 0;font-size:13px;">Gestão Empresarial Ltda</p></td></tr>' +
    '<tr><td style="padding:36px 40px;">' +
    '<h2 style="color:#211551;margin:0 0 16px;font-size:20px;">Ficha de Cadastro de Admissão</h2>' +
    '<p style="color:#475569;line-height:1.8;margin:0 0 12px;">Olá, <strong>' + nome + '</strong>!</p>' +
    '<p style="color:#475569;line-height:1.8;margin:0 0 28px;">Você foi convidado(a) a preencher a Ficha de Cadastro para a vaga de <strong>' + vaga + '</strong> na Premier Logistics.</p>' +
    '<div style="text-align:center;margin:32px 0;">' +
    '<a href="' + formUrl + '" style="background:#E8761A;color:white;padding:15px 36px;border-radius:8px;text-decoration:none;font-weight:bold;font-size:16px;display:inline-block;">Preencher Meu Formulário</a></div>' +
    '<p style="color:#94a3b8;font-size:12px;margin:0;line-height:1.7;">⚠️ Este link é pessoal e intransferível.<br>Link direto: <a href="' + formUrl + '" style="color:#3b82f6;word-break:break-all;">' + formUrl + '</a></p>' +
    '</td></tr><tr><td style="background:#f8fafc;padding:20px 40px;border-top:1px solid #e2e8f0;">' +
    '<p style="color:#94a3b8;font-size:12px;margin:0;text-align:center;">Equipe de Recursos Humanos · Premier Logistics Gestão Empresarial Ltda</p></td></tr>' +
    '</table></td></tr></table></body></html>';

  var plain = 'Olá ' + nome + ',\n\nVocê foi convidado(a) a preencher a Ficha de Cadastro para a vaga de ' + vaga + '.\n\nAcesse: ' + formUrl + '\n\nAtenciosamente,\nRH — Premier Logistics';
  GmailApp.sendEmail(email, CONFIG.EMAIL_SUBJECT, plain, { name: CONFIG.EMAIL_FROM_NAME, htmlBody: html });
}

function ensureSheets() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  var cs = ss.getSheetByName(SHEET_CANDIDATES);
  if (!cs) cs = ss.insertSheet(SHEET_CANDIDATES);
  if (cs.getLastRow() === 0) {
    cs.appendRow(['ID','Nome','Email','Vaga','Token','Status','DataEnvio','DataPreenchimento']);
    cs.setFrozenRows(1);
    cs.getRange('1:1').setFontWeight('bold').setBackground('#211551').setFontColor('white');
  }

  var rs = ss.getSheetByName(SHEET_RESPONSES);
  if (!rs) rs = ss.insertSheet(SHEET_RESPONSES);
  if (rs.getLastRow() === 0) {
    rs.appendRow(['Token','DataPreenchimento','NomeCandidato','NomeCompleto','NomeSocial','CPF','RG',
      'Endereço','BairroCidade','CEP','WhatsApp','Email','EmergenciaNome','EmergenciaTel',
      'TítuloEleitor','GrauInstrução','PossuiFilhos','QtdFilhos','DeclararIR','QtdDepIR',
      'DependentesIR','EstadoCivil','EstadoCivilOutro','NúmeroBota','TamanhoCamisa',
      'TamanhoCalça','OptanteVT','PlanoSaúde','Dep1Nome','Dep1CPF','Dep2Nome','Dep2CPF','TipoAssinatura']);
    rs.setFrozenRows(1);
    rs.getRange('1:1').setFontWeight('bold').setBackground('#211551').setFontColor('white');
  }
}

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
