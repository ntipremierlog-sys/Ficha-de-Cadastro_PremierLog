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
  EMAIL_SUBJECT:   'Premier Logistics — Ficha de Cadastro de Admissão',
  DRIVE_FOLDER_ID: '1D4_gClsf_HTG57LKGDhqZbFFrKUqjQqX'  // Pasta Google Drive para os PDFs
};

const SHEET_CANDIDATES = 'Candidatos';
const SHEET_RESPONSES  = 'Respostas';

// ============================================================
// DIAGNÓSTICO: Execute esta função no editor para autorizar o Drive
// Vá em: Executar → testarAcessoDrive
// ============================================================
function testarAcessoDrive() {
  try {
    var folder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);
    Logger.log('✅ Pasta encontrada: ' + folder.getName());
    // Cria e remove um arquivo de teste para confirmar permissão de escrita
    var blob = Utilities.newBlob('teste', 'text/plain', '_teste_permissao.txt');
    var file = folder.createFile(blob);
    Logger.log('✅ Arquivo de teste criado: ' + file.getId());
    file.setTrashed(true);
    Logger.log('✅ DriveApp autorizado com sucesso! PDFs serão salvos normalmente.');
    return 'Sucesso — Drive autorizado!';
  } catch(e) {
    Logger.log('❌ Erro: ' + e.message);
    return 'Erro: ' + e.message;
  }
}

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
      case 'getDebugInfo':     return jsonResponse(getDebugInfo(e.parameter.secret));
      case 'resetCandidateStatus': return jsonResponse(resetCandidateStatus(e.parameter.token, e.parameter.secret));
      case 'cleanTestResponses':   return jsonResponse(cleanTestResponses(e.parameter.secret));
      case 'syncStatuses':           return jsonResponse(syncStatuses(e.parameter.secret));
      case 'bulkConcluirByDate':     return jsonResponse(bulkConcluirByDate(e.parameter.secret, e.parameter.dataLimite));
      default:                return jsonResponse({ error: 'Ação inválida: ' + action });
    }
  } catch (err) {
    console.error('doGet error:', err);
    return jsonResponse({ error: err.message });
  }
}

// ============================================================
// FUNÇÃO: Obter Informações de Diagnóstico
// ============================================================
function getDebugInfo(secret) {
  if (secret !== CONFIG.ADMIN_SECRET) return { error: 'Não autorizado' };
  ensureSheets();
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var cs = ss.getSheetByName(SHEET_CANDIDATES);
  var rs = ss.getSheetByName(SHEET_RESPONSES);
  
  var csRows = cs.getDataRange().getValues();
  var rsRows = rs.getDataRange().getValues();
  
  var folderStatus = '';
  try {
    var folder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);
    folderStatus = 'OK - Nome: ' + folder.getName();
  } catch(e) {
    folderStatus = 'ERRO: ' + e.message;
  }
  
  return {
    folderStatus: folderStatus,
    candidatesHeaders: csRows[0],
    lastCandidates: csRows.slice(-3).map(function(row) {
      return {
        id: row[0],
        nome: row[1],
        status: row[5],
        dateEnvio: row[6],
        datePreenchimento: row[7],
        pdfUrl: row[8]
      };
    }),
    responsesHeaders: rsRows[0],
    lastResponses: rsRows.slice(-3).map(function(row) {
      return {
        token: row[0],
        datePreenchimento: row[1],
        nome: row[2],
        pdfUrl: row[row.length - 1]
      };
    })
  };
}

// ============================================================
// ROTEADOR POST (form submission)
// ============================================================
function doPost(e) {
  try {
    var bodySize = e.postData ? e.postData.contents.length : 0;
    Logger.log('POST recebido — tamanho do body: ' + bodySize + ' chars');
    const data = JSON.parse(e.postData.contents);
    Logger.log('action: ' + data.action);
    Logger.log('pdfBase64 presente: ' + (data.pdfBase64 ? 'SIM (' + data.pdfBase64.length + ' chars)' : 'NÃO'));
    if (data.action === 'submitForm') return jsonResponse(submitForm(data));
    return jsonResponse({ error: 'Ação POST inválida' });
  } catch (err) {
    Logger.log('doPost ERRO: ' + err.message);
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
      dataPreenchimento: data[i][7] || '',
      pdfUrl:            data[i][8] || ''
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

  // Buscar pdfUrl na aba Candidatos
  var pdfUrl = '';
  try {
    var cSheet2 = ss.getSheetByName(SHEET_CANDIDATES);
    var cData2  = cSheet2.getDataRange().getValues();
    for (var k = 1; k < cData2.length; k++) {
      if (cData2[k][4] === token) { pdfUrl = cData2[k][8] || ''; break; }
    }
  } catch(e) {}

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === token) {
      var response = {};
      headers.forEach(function(h, j) { response[h] = data[i][j] || ''; });
      response.pdfUrl = pdfUrl;
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
  var candidateRow  = -1;

  for (var i = 1; i < cData.length; i++) {
    var sheetToken  = String(cData[i][4] || '').trim();
    var sheetStatus = String(cData[i][5] || '').trim();
    var reqToken    = String(token || '').trim();

    if (sheetToken === reqToken) {
      if (sheetStatus === 'Concluído') return { error: 'Formulário já preenchido anteriormente.' };
      
      // Atualizar status para Concluído e gravar imediatamente no Google Sheets
      cSheet.getRange(i + 1, 6).setValue('Concluído');
      cSheet.getRange(i + 1, 8).setValue(now);
      SpreadsheetApp.flush(); // Persiste alteração de status imediatamente
      
      candidateName = cData[i][1];
      candidateRow  = i + 1;
      break;
    }
  }

  // ⚠️ Token não encontrado — retorna erro imediatamente
  if (candidateRow === -1) {
    return { error: 'Token não encontrado. O link pode ser inválido ou expirado.' };
  }

  // ============================================================
  // SALVAR PDF NO GOOGLE DRIVE (Isolado em try/catch para não reverter status)
  // ============================================================
  var pdfUrl = '';
  var pdfError = '';
  if (data.pdfBase64 && candidateName) {
    try {
      var base64Data = data.pdfBase64.split(';base64,')[1] || data.pdfBase64;
      var pdfBytes   = Utilities.base64Decode(base64Data);
      var pdfBlob    = Utilities.newBlob(pdfBytes, 'application/pdf',
        'Ficha_Cadastro_' + candidateName.replace(/\s+/g, '_').substring(0, 30) + '.pdf');

      var folder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);
      var file   = folder.createFile(pdfBlob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      pdfUrl = 'https://drive.google.com/file/d/' + file.getId() + '/view';

      // Registrar link na aba Candidatos (coluna 9 = índice 8)
      if (candidateRow > 0) {
        cSheet.getRange(candidateRow, 9).setValue(pdfUrl);
        SpreadsheetApp.flush();
      }
    } catch (err) {
      console.error('Erro ao salvar PDF no Drive:', err);
      pdfError = err.message;
    }
  }

  // ============================================================
  // REGISTRAR RESPOSTA NA ABA RESPOSTAS
  // ============================================================
  try {
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
      fd.tipoAssinatura || '', pdfUrl
    ]);
    SpreadsheetApp.flush();
  } catch (errResp) {
    console.error('Erro ao registrar resposta:', errResp);
  }

  return { success: true, message: 'Formulário recebido com sucesso!', pdfUrl: pdfUrl, pdfError: pdfError };
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
    if (data[i][4] === token) { 
      sheet.deleteRow(i + 1); 
      
      // Também remover da aba Respostas se existir
      var rSheet = ss.getSheetByName(SHEET_RESPONSES);
      if (rSheet) {
        var rData = rSheet.getDataRange().getValues();
        for (var j = 1; j < rData.length; j++) {
          if (rData[j][0] === token) {
            rSheet.deleteRow(j + 1);
            break;
          }
        }
      }
      return { success: true }; 
    }
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
    cs.appendRow(['ID','Nome','Email','Vaga','Token','Status','DataEnvio','DataPreenchimento','PDFUrl']);
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
      'TamanhoCalça','OptanteVT','PlanoSaúde','Dep1Nome','Dep1CPF','Dep2Nome','Dep2CPF','TipoAssinatura','PDFUrl']);
    rs.setFrozenRows(1);
    rs.getRange('1:1').setFontWeight('bold').setBackground('#211551').setFontColor('white');
  }
}

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function resetCandidateStatus(token, secret) {
  if (secret !== CONFIG.ADMIN_SECRET) return { error: 'Não autorizado' };
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var cs = ss.getSheetByName(SHEET_CANDIDATES);
  var cData = cs.getDataRange().getValues();
  for (var i = 1; i < cData.length; i++) {
    if (cData[i][4] === token) {
      cs.getRange(i + 1, 6).setValue('Pendente');
      cs.getRange(i + 1, 8).setValue('');
      
      // Também remover da aba Respostas se existir
      var rs = ss.getSheetByName(SHEET_RESPONSES);
      var rData = rs.getDataRange().getValues();
      for (var j = 1; j < rData.length; j++) {
        if (rData[j][0] === token) {
          rs.deleteRow(j + 1);
          break;
        }
      }
      return { success: true };
    }
  }
  return { error: 'Candidato não encontrado' };
}

function cleanTestResponses(secret) {
  if (secret !== CONFIG.ADMIN_SECRET) return { error: 'Não autorizado' };
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var rs = ss.getSheetByName(SHEET_RESPONSES);
  if (!rs) return { success: true, message: 'Aba Respostas não encontrada' };
  
  var data = rs.getDataRange().getValues();
  var deletedCount = 0;
  var namesToRemove = [
    'geovanna correia andrade',
    'ademir do nascimento santana',
    'geovanna c andrade',
    'geovanna andrade',
    'gabrielly mendes dos santos',
    'gabrielly mendes',
    'rafael araujo',
    'priscila facchinetti',
    'teste pdf sistema',
    'teste drive pdf',
    'teste log pdfteste automatizado pdf'
  ];
  
  for (var i = data.length - 1; i >= 1; i--) {
    var nameInRow = String(data[i][2] || '').toLowerCase().trim(); // NomeCandidato
    var fullNameInRow = String(data[i][3] || '').toLowerCase().trim(); // NomeCompleto
    
    var shouldDelete = namesToRemove.some(function(name) {
      return nameInRow.indexOf(name) !== -1 || fullNameInRow.indexOf(name) !== -1;
    });
    
    if (shouldDelete) {
      rs.deleteRow(i + 1);
      deletedCount++;
    }
  }
  return { success: true, deletedCount: deletedCount };
}

// ============================================================
// FUNÇÃO: Sincronizar Status — cruza Respostas × Candidatos
// Para cada token presente na aba Respostas, se o candidato
// ainda estiver como "Pendente" na aba Candidatos, atualiza
// para "Concluído" e preenche a data de preenchimento.
// ============================================================
function syncStatuses(secret) {
  if (secret !== CONFIG.ADMIN_SECRET) return { error: 'Não autorizado' };
  ensureSheets();

  var ss      = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var cSheet  = ss.getSheetByName(SHEET_CANDIDATES);
  var rSheet  = ss.getSheetByName(SHEET_RESPONSES);

  var cData   = cSheet.getDataRange().getValues();
  var rData   = rSheet.getDataRange().getValues();

  // Construir um mapa: token → { dataPreenchimento, pdfUrl } a partir das Respostas
  // Colunas Respostas: [0]=Token [1]=DataPreenchimento [2]=NomeCandidato ... [last]=PDFUrl
  var responseMap = {};
  for (var r = 1; r < rData.length; r++) {
    var tok  = String(rData[r][0] || '').trim();
    var date = String(rData[r][1] || '').trim();
    var pdf  = String(rData[r][rData[r].length - 1] || '').trim();
    if (tok) responseMap[tok] = { date: date, pdf: pdf };
  }

  var updated = 0;
  var skipped = 0;

  // Percorrer candidatos e corrigir os que têm resposta mas estão Pendente
  for (var i = 1; i < cData.length; i++) {
    var candidateToken  = String(cData[i][4] || '').trim();
    var candidateStatus = String(cData[i][5] || '').trim();

    if (!candidateToken) continue;

    if (responseMap[candidateToken]) {
      // Há resposta para este token
      if (candidateStatus !== 'Concluído') {
        // Marcar como Concluído
        cSheet.getRange(i + 1, 6).setValue('Concluído');

        // Preencher data de preenchimento se estiver vazia
        if (!cData[i][7] && responseMap[candidateToken].date) {
          cSheet.getRange(i + 1, 8).setValue(responseMap[candidateToken].date);
        }

        // Preencher PDF URL se estiver vazia
        if (!cData[i][8] && responseMap[candidateToken].pdf) {
          cSheet.getRange(i + 1, 9).setValue(responseMap[candidateToken].pdf);
        }

        updated++;
        Logger.log('✅ Corrigido: ' + cData[i][1] + ' | Token: ' + candidateToken);
      } else {
        skipped++; // Já estava Concluído — OK
      }
    }
  }

  return {
    success: true,
    updated: updated,
    alreadyCorrect: skipped,
    message: updated + ' candidato(s) tiveram o status corrigido para Concluído. ' + skipped + ' já estavam corretos.'
  };
}

// ============================================================
// FUNÇÃO: Marcar em massa como Concluído até uma data limite
// Percorre todos os candidatos com status "Pendente" e
// dataEnvio <= dataLimite, atualizando para "Concluído".
// Parâmetro dataLimite: formato 'YYYY-MM-DD' (ex: '2026-07-22')
// ============================================================
function bulkConcluirByDate(secret, dataLimite) {
  if (secret !== CONFIG.ADMIN_SECRET) return { error: 'Não autorizado' };
  if (!dataLimite) return { error: 'Parâmetro dataLimite é obrigatório (formato: YYYY-MM-DD)' };

  // Converter a data limite para timestamp (fim do dia: 23:59:59)
  var limiteParts = dataLimite.split('-');
  if (limiteParts.length !== 3) return { error: 'Formato de data inválido. Use YYYY-MM-DD' };
  var limiteDate = new Date(parseInt(limiteParts[0]), parseInt(limiteParts[1]) - 1, parseInt(limiteParts[2]), 23, 59, 59, 999);

  ensureSheets();
  var ss     = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var cSheet = ss.getSheetByName(SHEET_CANDIDATES);
  var cData  = cSheet.getDataRange().getValues();

  var updated  = 0;
  var skipped  = 0;
  var errors   = [];
  var updatedNames = [];

  for (var i = 1; i < cData.length; i++) {
    var status    = String(cData[i][5] || '').trim();
    var dataEnvio = cData[i][6]; // Coluna DataEnvio (índice 6)

    // Só processa candidatos Pendentes
    if (status !== 'Pendente') { skipped++; continue; }

    // Converter dataEnvio para Date
    var envioDate = null;
    if (dataEnvio instanceof Date) {
      envioDate = dataEnvio;
    } else if (typeof dataEnvio === 'string' && dataEnvio.trim() !== '') {
      // Tentar parsear formatos comuns: 'dd/MM/yyyy HH:mm' ou ISO
      var iso = dataEnvio.trim();
      envioDate = new Date(iso);
      if (isNaN(envioDate.getTime())) {
        // Tentar formato 'dd/MM/yyyy HH:mm'
        var match = iso.match(/^(\d{2})\/(\d{2})\/(\d{4})\s+(\d{2}):(\d{2})/);
        if (match) {
          envioDate = new Date(parseInt(match[3]), parseInt(match[2]) - 1, parseInt(match[1]), parseInt(match[4]), parseInt(match[5]));
        }
      }
    } else if (typeof dataEnvio === 'number') {
      // Número serial do Google Sheets (dias desde 30/12/1899)
      envioDate = new Date((dataEnvio - 25569) * 86400 * 1000);
    }

    if (!envioDate || isNaN(envioDate.getTime())) {
      errors.push('Linha ' + (i + 1) + ': data de envio inválida (' + dataEnvio + ')');
      continue;
    }

    // Verificar se dataEnvio <= dataLimite
    if (envioDate <= limiteDate) {
      try {
        cSheet.getRange(i + 1, 6).setValue('Concluído');
        updated++;
        updatedNames.push(cData[i][1] + ' (envio: ' + envioDate.toISOString().substring(0, 10) + ')');
        Logger.log('✅ Marcado como Concluído: ' + cData[i][1] + ' | DataEnvio: ' + envioDate.toISOString());
      } catch (err) {
        errors.push('Linha ' + (i + 1) + ': ' + err.message);
      }
    } else {
      skipped++;
    }
  }

  return {
    success: true,
    updated: updated,
    skipped: skipped,
    errors: errors,
    updatedNames: updatedNames,
    message: updated + ' candidato(s) marcados como Concluído. ' + skipped + ' ignorados (já Concluído ou após a data limite). ' + errors.length + ' erro(s).'
  };
}
