const CONFIG = {
  CONTROL_TITLE: 'Portal de Planilhas — Controle de Acessos',
  AUDIT_TITLE: 'Portal de Planilhas — Auditoria de Acessos',
  TIMEZONE: 'America/Fortaleza',
  LOCALE: 'pt_BR',
  SESSION_TIMEOUT_SECONDS: 150,
  HEARTBEAT_SECONDS: 60,
  SHEETS: {
    USERS: 'Usuarios',
    SPREADSHEETS: 'Planilhas',
    SESSIONS: 'Sessoes',
    ACCESS: 'AcessosPlanilhas',
    LOGINS: 'Logins'
  }
};

const CONTROL_SPREADSHEET_PROPERTY = 'CONTROL_SPREADSHEET_ID';
const AUDIT_SPREADSHEET_PROPERTY = 'AUDIT_SPREADSHEET_ID';

// Planilha nativa do Google Drive / Google Sheets.
const MAIN_SPREADSHEET_ID = '1jabSU0G5IYoKP6raGtZszQuV8Ti92MJu9E44zkOWrCc';
const MAIN_SHEET_GID = 349757095;

// Planilha consolidada que recebe uma linha por sessão do portal.
const SESSION_LOG_SPREADSHEET_ID = '1Q4QHcoPY27fboJEfq73pYgRrntG50N6c8rAx2VdqtHw';
const SESSION_LOG_SHEET_NAME = 'Registros';

// Usuários iniciais: apenas salt + SHA-256; senha não fica em texto puro.
const INITIAL_USERS = [
  { usuario: 'alefim', nome: 'Alefim', salt: '0350c5d10b0835c01e5dd13db0229712', senhaHash: '742d32131837abd2901134ce4dd47d64e213106e8f82f902acbde9ef3c4ef1cb' },
  { usuario: 'joao', nome: 'João', salt: '4f28753e17b1c890d8ee66e720962aa1', senhaHash: '649062a8edd68f8448a07fd355db37455d6e8b28418a0cab1184c58a83f66975' },
  { usuario: 'robson', nome: 'Robson', salt: '4b353979857331a20428172c91fd5e12', senhaHash: 'c74d4f1ea85a338ba16066bc6842ec1255cc1e76cb3ced449540d1cd29193947' },
  { usuario: 'julenio', nome: 'Julenio', salt: '2988c4fcd2a4ecd3535364269f141865', senhaHash: 'ee08ae5a10766fe6c28f86b06d47dd35bce7a6058da2ff41665166ed76f8cb13' },
  { usuario: 'mateus', nome: 'Mateus', salt: '245326e5b7956d241c56ecd260122aab', senhaHash: '377ad683f0ca5911f5791262ed420aa3db86174f02d23faef6e5a6ed7b4b29fa' },
  { usuario: 'eliane', nome: 'Eliane', salt: 'b49fff0ab8b61b1b724d4970f7f0e2c5', senhaHash: 'd826ec870095b9316967d4f2bc3778ec63bb29fa311997137ef97169e03873f7' },
  { usuario: 'clairton', nome: 'Clairton', salt: '775cbc11ae1f3705d957c51040b6e0f7', senhaHash: 'dfcb472023dfd6f8603209ed12b0f5061e9a1a796aaee914d87026bdffb4b584' }
];

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Portal de Planilhas')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getMainSpreadsheetConfig_() {
  let ss;
  try {
    ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
  } catch (error) {
    throw new Error('A planilha principal não está acessível no Google Drive desta conta. Confirme que a conta do Apps Script possui acesso ao arquivo.');
  }

  const targetSheet = ss.getSheets().find(function(sheet) {
    return Number(sheet.getSheetId()) === Number(MAIN_SHEET_GID);
  });

  if (!targetSheet) {
    throw new Error('A aba gid=' + MAIN_SHEET_GID + ' não foi encontrada na planilha do Google Drive.');
  }

  return {
    id: ss.getId(),
    nome: ss.getName(),
    aba: targetSheet.getName(),
    url: ss.getUrl() + '?gid=' + MAIN_SHEET_GID + '#gid=' + MAIN_SHEET_GID,
    ordem: 1
  };
}

function setupSistema() {
  const principal = getMainSpreadsheetConfig_();
  ensureSessionLogSpreadsheet_();
  const controle = getOrCreateSpreadsheet_(CONTROL_SPREADSHEET_PROPERTY, CONFIG.CONTROL_TITLE);
  const auditoria = getOrCreateSpreadsheet_(AUDIT_SPREADSHEET_PROPERTY, CONFIG.AUDIT_TITLE);

  ensureControlSheets_(controle);
  ensureAuditSheets_(auditoria);
  ensureInitialUsers_(controle);
  syncMainSpreadsheet_(controle, principal);
  cleanupBlankDefaultSheets_(controle);
  cleanupBlankDefaultSheets_(auditoria);

  return [
    'Sistema configurado com sucesso.',
    'Planilha do Google Drive: ' + principal.nome,
    'Aba inicial: ' + principal.aba,
    'Controle: ' + controle.getUrl(),
    'Auditoria: ' + auditoria.getUrl()
  ].join('\n');
}

function configurarNovaConta() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty(CONTROL_SPREADSHEET_PROPERTY);
  props.deleteProperty(AUDIT_SPREADSHEET_PROPERTY);
  return setupSistema();
}

function diagnosticoSistema() {
  const principal = getMainSpreadsheetConfig_();
  const controle = getDb_();
  const auditoria = getAuditDb_();
  return {
    ok: true,
    planilhaPrincipal: principal.nome,
    planilhaPrincipalId: principal.id,
    abaPrincipal: principal.aba,
    planilhaPrincipalUrl: principal.url,
    controle: controle.getUrl(),
    auditoria: auditoria.getUrl(),
    usuarios: countConfiguredUsers_(controle),
    timezone: Session.getScriptTimeZone()
  };
}

function login(usuario, senha, userAgent) {
  cleanupStaleSessions_();
  usuario = normalizeUser_(usuario);
  senha = String(senha || '');

  if (!usuario || !senha) return { ok: false, message: 'Informe usuário e senha.' };

  const cache = CacheService.getScriptCache();
  const attemptKey = 'login:' + usuario;
  const attempts = Number(cache.get(attemptKey) || 0);
  if (attempts >= 8) {
    return { ok: false, message: 'Muitas tentativas. Aguarde alguns minutos e tente novamente.' };
  }

  const user = findUser_(usuario);
  if (!user || !user.ativo || hashPassword_(senha, user.salt) !== user.senhaHash) {
    registerLoginAttempt_(usuario, 'FALHA', '', userAgent);
    cache.put(attemptKey, String(attempts + 1), 300);
    Utilities.sleep(250);
    return { ok: false, message: 'Usuário ou senha inválidos.' };
  }

  getMainSpreadsheetConfig_();
  cache.remove(attemptKey);
  const sessionId = Utilities.getUuid();
  const now = new Date();
  const row = [sessionId, user.usuario, user.nome, now, now, '', 0, '00:00:00', 'ATIVA', String(userAgent || '').substring(0, 500)];

  const sessionsSheet = getDb_().getSheetByName(CONFIG.SHEETS.SESSIONS);
  sessionsSheet.appendRow(row);
  syncSessionRegister_(sessionsSheet.getLastRow());
  appendAuditRow_(CONFIG.SHEETS.SESSIONS, row);
  registerLoginAttempt_(user.usuario, 'SUCESSO', sessionId, userAgent);
  return buildLoginResponse_(user, sessionId, 0);
}

function restoreSession(sessionId) {
  cleanupStaleSessions_();
  const session = getSession_(sessionId);
  if (!session || session.status !== 'ATIVA') return { ok: false };
  getMainSpreadsheetConfig_();
  touchSession_(session.row);
  const seconds = Math.max(0, Math.floor((Date.now() - session.entrada.getTime()) / 1000));
  return buildLoginResponse_({ usuario: session.usuario, nome: session.nome }, sessionId, seconds);
}

function heartbeat(sessionId) {
  const session = getSession_(sessionId);
  if (!session || session.status !== 'ATIVA') return { ok: false };
  const elapsed = (Date.now() - session.ultimaAtividade.getTime()) / 1000;
  if (elapsed > CONFIG.SESSION_TIMEOUT_SECONDS) {
    closeSessionRow_(session.row, 'ENCERRADA_AUTOMATICAMENTE', session.ultimaAtividade);
    return { ok: false };
  }
  touchSession_(session.row);
  return { ok: true };
}

function registerSpreadsheetOpen(sessionId, spreadsheetId) {
  const session = getSession_(sessionId);
  if (!session || session.status !== 'ATIVA') return { ok: false, message: 'Sessão expirada.' };
  if (String(spreadsheetId) !== MAIN_SPREADSHEET_ID) return { ok: false, message: 'Planilha não autorizada.' };

  const principal = getMainSpreadsheetConfig_();
  touchSession_(session.row);
  const row = [Utilities.getUuid(), sessionId, session.usuario, principal.id, principal.nome, new Date()];
  getDb_().getSheetByName(CONFIG.SHEETS.ACCESS).appendRow(row);
  appendAuditRow_(CONFIG.SHEETS.ACCESS, row);
  syncSessionRegister_(session.row);
  return { ok: true, url: principal.url };
}

function logout(sessionId) {
  const session = getSession_(sessionId);
  if (session && session.status === 'ATIVA') closeSessionRow_(session.row, 'ENCERRADA', new Date());
  return { ok: true };
}

function cleanupStaleSessions() {
  return cleanupStaleSessions_();
}

function addUser(usuario, nome, senha) {
  usuario = normalizeUser_(usuario);
  nome = String(nome || '').trim();
  senha = String(senha || '');
  if (!usuario || usuario.length < 3) throw new Error('Usuário inválido.');
  if (!nome) throw new Error('Informe o nome.');
  if (senha.length < 8) throw new Error('A senha precisa ter pelo menos 8 caracteres.');

  const sheet = getDb_().getSheetByName(CONFIG.SHEETS.USERS);
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (normalizeUser_(values[i][0]) === usuario) throw new Error('Esse usuário já existe.');
  }
  const salt = Utilities.getUuid().replace(/-/g, '');
  sheet.appendRow([usuario, nome, salt, hashPassword_(senha, salt), true, new Date()]);
  return 'Usuário criado com sucesso.';
}

function trocarSenha(usuario, novaSenha) {
  usuario = normalizeUser_(usuario);
  novaSenha = String(novaSenha || '');
  if (!usuario) throw new Error('Informe o usuário.');
  if (novaSenha.length < 8) throw new Error('A nova senha precisa ter pelo menos 8 caracteres.');

  const sheet = getDb_().getSheetByName(CONFIG.SHEETS.USERS);
  if (!sheet || sheet.getLastRow() < 2) throw new Error('Usuário não encontrado.');
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
  for (let i = 0; i < values.length; i++) {
    if (normalizeUser_(values[i][0]) === usuario) {
      const salt = Utilities.getUuid().replace(/-/g, '');
      const row = i + 2;
      sheet.getRange(row, 3).setValue(salt);
      sheet.getRange(row, 4).setValue(hashPassword_(novaSenha, salt));
      sheet.getRange(row, 5).setValue(true);
      return 'Senha alterada com sucesso para ' + usuario + '.';
    }
  }
  throw new Error('Usuário não encontrado.');
}

function buildLoginResponse_(user, sessionId, currentSeconds) {
  const dashboard = getUserDashboardData_(user.usuario, sessionId);
  return {
    ok: true,
    sessionId: sessionId,
    user: { usuario: user.usuario, nome: user.nome },
    spreadsheets: dashboard.spreadsheets,
    recentSessions: dashboard.recentSessions,
    totalSessions: dashboard.totalSessions,
    totalAccesses: dashboard.totalAccesses,
    currentSessionDurationSeconds: currentSeconds,
    heartbeatSeconds: CONFIG.HEARTBEAT_SECONDS
  };
}

function getUserDashboardData_(usuario, currentSessionId) {
  return {
    spreadsheets: getActiveSpreadsheets_(),
    recentSessions: getRecentSessions_(usuario, currentSessionId),
    totalSessions: countRowsForUser_(CONFIG.SHEETS.SESSIONS, 2, usuario),
    totalAccesses: countRowsForUser_(CONFIG.SHEETS.ACCESS, 3, usuario)
  };
}

function getActiveSpreadsheets_() {
  const principal = getMainSpreadsheetConfig_();
  return [{ id: principal.id, nome: principal.nome, url: principal.url, ordem: 1 }];
}

function countRowsForUser_(sheetName, column, usuario) {
  const sheet = getDb_().getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return 0;
  return sheet.getRange(2, column, sheet.getLastRow() - 1, 1).getValues().reduce(function(sum, row) {
    return sum + (normalizeUser_(row[0]) === usuario ? 1 : 0);
  }, 0);
}

function getRecentSessions_(usuario, currentSessionId) {
  const sheet = getDb_().getSheetByName(CONFIG.SHEETS.SESSIONS);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const tz = Session.getScriptTimeZone() || CONFIG.TIMEZONE;
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).getValues()
    .filter(function(row) { return normalizeUser_(row[1]) === usuario; })
    .map(function(row) {
      const entrada = toDate_(row[3]);
      return {
        current: String(row[0]) === String(currentSessionId),
        entradaLabel: entrada && !isNaN(entrada.getTime()) ? Utilities.formatDate(entrada, tz, "dd/MM/yyyy 'às' HH:mm") : 'Data indisponível',
        status: String(row[8] || ''),
        duration: String(row[7] || '00:00:00'),
        browser: browserName_(String(row[9] || ''))
      };
    }).reverse().slice(0, 6);
}

function findUser_(usuario) {
  const sheet = getDb_().getSheetByName(CONFIG.SHEETS.USERS);
  if (!sheet || sheet.getLastRow() < 2) return null;
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
  for (let i = 0; i < values.length; i++) {
    if (normalizeUser_(values[i][0]) === usuario) {
      return { usuario: normalizeUser_(values[i][0]), nome: String(values[i][1] || ''), salt: String(values[i][2] || ''), senhaHash: String(values[i][3] || ''), ativo: toBoolean_(values[i][4]) };
    }
  }
  return null;
}

function getSession_(sessionId) {
  sessionId = String(sessionId || '').trim();
  if (!sessionId) return null;
  const sheet = getDb_().getSheetByName(CONFIG.SHEETS.SESSIONS);
  if (!sheet || sheet.getLastRow() < 2) return null;
  const finder = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).createTextFinder(sessionId).matchEntireCell(true).findNext();
  if (!finder) return null;
  const rowNumber = finder.getRow();
  const row = sheet.getRange(rowNumber, 1, 1, 10).getValues()[0];
  return { row: rowNumber, sessionId: String(row[0]), usuario: String(row[1]), nome: String(row[2]), entrada: toDate_(row[3]), ultimaAtividade: toDate_(row[4]), status: String(row[8] || '') };
}

function touchSession_(rowNumber) {
  const sheet = getDb_().getSheetByName(CONFIG.SHEETS.SESSIONS);
  const entrada = toDate_(sheet.getRange(rowNumber, 4).getValue());
  const now = new Date();
  const seconds = Math.max(0, Math.floor((now.getTime() - entrada.getTime()) / 1000));
  sheet.getRange(rowNumber, 5).setValue(now);
  sheet.getRange(rowNumber, 7).setValue(seconds);
  sheet.getRange(rowNumber, 8).setValue(formatDuration_(seconds));
  upsertAuditSession_(rowNumber);
  syncSessionRegister_(rowNumber);
}

function closeSessionRow_(rowNumber, status, closeTime) {
  const sheet = getDb_().getSheetByName(CONFIG.SHEETS.SESSIONS);
  const entrada = toDate_(sheet.getRange(rowNumber, 4).getValue());
  const end = closeTime instanceof Date ? closeTime : new Date();
  const seconds = Math.max(0, Math.floor((end.getTime() - entrada.getTime()) / 1000));
  sheet.getRange(rowNumber, 5).setValue(end);
  sheet.getRange(rowNumber, 6).setValue(end);
  sheet.getRange(rowNumber, 7).setValue(seconds);
  sheet.getRange(rowNumber, 8).setValue(formatDuration_(seconds));
  sheet.getRange(rowNumber, 9).setValue(status);
  upsertAuditSession_(rowNumber);
  syncSessionRegister_(rowNumber);
}

function cleanupStaleSessions_() {
  const sheet = getDb_().getSheetByName(CONFIG.SHEETS.SESSIONS);
  if (!sheet || sheet.getLastRow() < 2) return 0;
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).getValues();
  const now = new Date();
  let count = 0;
  rows.forEach(function(row, index) {
    const lastSeen = toDate_(row[4]);
    if (String(row[8] || '') === 'ATIVA' && lastSeen && !isNaN(lastSeen.getTime()) && (now.getTime() - lastSeen.getTime()) / 1000 > CONFIG.SESSION_TIMEOUT_SECONDS) {
      closeSessionRow_(index + 2, 'ENCERRADA_AUTOMATICAMENTE', lastSeen);
      count++;
    }
  });
  return count;
}

function registerLoginAttempt_(usuario, resultado, sessionId, userAgent) {
  const row = [new Date(), normalizeUser_(usuario), String(resultado || ''), String(sessionId || ''), String(userAgent || '').substring(0, 500)];
  const sheet = getDb_().getSheetByName(CONFIG.SHEETS.LOGINS);
  if (sheet) sheet.appendRow(row);
  appendAuditRow_(CONFIG.SHEETS.LOGINS, row);
}

function appendAuditRow_(sheetName, row) {
  try {
    const sheet = getAuditDb_().getSheetByName(sheetName);
    if (sheet) sheet.appendRow(row);
  } catch (error) {
    console.error('Falha ao registrar auditoria:', error);
  }
}

function upsertAuditSession_(controlRowNumber) {
  try {
    const control = getDb_().getSheetByName(CONFIG.SHEETS.SESSIONS);
    const row = control.getRange(controlRowNumber, 1, 1, 10).getValues()[0];
    const sessionId = String(row[0] || '');
    if (!sessionId) return;
    const audit = getAuditDb_().getSheetByName(CONFIG.SHEETS.SESSIONS);
    let finder = null;
    if (audit.getLastRow() >= 2) finder = audit.getRange(2, 1, audit.getLastRow() - 1, 1).createTextFinder(sessionId).matchEntireCell(true).findNext();
    if (finder) audit.getRange(finder.getRow(), 1, 1, 10).setValues([row]);
    else audit.appendRow(row);
  } catch (error) {
    console.error('Falha ao sincronizar sessão na auditoria:', error);
  }
}

function ensureSessionLogSpreadsheet_() {
  const ss = SpreadsheetApp.openById(SESSION_LOG_SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SESSION_LOG_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SESSION_LOG_SHEET_NAME);
  const headers = ['ID da sessão', 'Usuário', 'Entrada', 'Saída', 'Duração (segundos)', 'Duração', 'Aberturas de planilha', 'Status', 'Navegador/Dispositivo', 'Página/Origem', 'Observações'];
  if (sheet.getLastRow() === 0) sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  return sheet;
}

function syncSessionRegister_(controlRowNumber) {
  try {
    const control = getDb_().getSheetByName(CONFIG.SHEETS.SESSIONS);
    const row = control.getRange(controlRowNumber, 1, 1, 10).getValues()[0];
    const sessionId = String(row[0] || '');
    if (!sessionId) return;

    const accessSheet = getDb_().getSheetByName(CONFIG.SHEETS.ACCESS);
    let openingCount = 0;
    if (accessSheet && accessSheet.getLastRow() >= 2) {
      openingCount = accessSheet.getRange(2, 2, accessSheet.getLastRow() - 1, 1).getValues().reduce(function(total, item) {
        return total + (String(item[0] || '') === sessionId ? 1 : 0);
      }, 0);
    }

    const statusMap = {
      'ATIVA': 'Ativa',
      'ENCERRADA': 'Encerrada',
      'ENCERRADA_AUTOMATICAMENTE': 'Expirada'
    };
    const seconds = Math.max(0, Number(row[6] || 0));
    const values = [[
      sessionId,
      String(row[1] || ''),
      row[3] || '',
      row[5] || '',
      seconds,
      seconds / 86400,
      openingCount,
      statusMap[String(row[8] || '')] || String(row[8] || 'Ativa'),
      String(row[9] || '').substring(0, 500),
      'Portal de Planilhas',
      String(row[2] || '')
    ]];

    const target = ensureSessionLogSpreadsheet_();
    let finder = null;
    if (target.getLastRow() >= 2) finder = target.getRange(2, 1, target.getLastRow() - 1, 1).createTextFinder(sessionId).matchEntireCell(true).findNext();
    if (finder) target.getRange(finder.getRow(), 1, 1, 11).setValues(values);
    else target.appendRow(values[0]);
  } catch (error) {
    console.error('Falha ao atualizar a planilha consolidada de sessões:', error);
  }
}

function ensureControlSheets_(ss) {
  ensureSheet_(ss, CONFIG.SHEETS.USERS, ['usuario', 'nome', 'salt', 'senhaHash', 'ativo', 'criadoEm']);
  ensureSheet_(ss, CONFIG.SHEETS.SPREADSHEETS, ['id', 'nome', 'url', 'ativo', 'ordem']);
  ensureSheet_(ss, CONFIG.SHEETS.SESSIONS, ['sessaoId', 'usuario', 'nome', 'entrada', 'ultimaAtividade', 'saida', 'duracaoSegundos', 'duracao', 'status', 'userAgent']);
  ensureSheet_(ss, CONFIG.SHEETS.ACCESS, ['acessoId', 'sessaoId', 'usuario', 'planilhaId', 'planilhaNome', 'abertaEm']);
  ensureSheet_(ss, CONFIG.SHEETS.LOGINS, ['dataHora', 'usuario', 'resultado', 'sessaoId', 'userAgent']);
}

function ensureAuditSheets_(ss) {
  ensureSheet_(ss, CONFIG.SHEETS.LOGINS, ['dataHora', 'usuario', 'resultado', 'sessaoId', 'userAgent']);
  ensureSheet_(ss, CONFIG.SHEETS.SESSIONS, ['sessaoId', 'usuario', 'nome', 'entrada', 'ultimaAtividade', 'saida', 'duracaoSegundos', 'duracao', 'status', 'userAgent']);
  ensureSheet_(ss, CONFIG.SHEETS.ACCESS, ['acessoId', 'sessaoId', 'usuario', 'planilhaId', 'planilhaNome', 'abertaEm']);
}

function ensureInitialUsers_(ss) {
  const sheet = ss.getSheetByName(CONFIG.SHEETS.USERS);
  const existing = {};
  if (sheet.getLastRow() >= 2) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues().forEach(function(row, index) {
      const key = normalizeUser_(row[0]);
      if (key) existing[key] = index + 2;
    });
  }
  INITIAL_USERS.forEach(function(user) {
    const key = normalizeUser_(user.usuario);
    const rowNumber = existing[key];
    if (rowNumber) {
      sheet.getRange(rowNumber, 1).setValue(key);
      sheet.getRange(rowNumber, 2).setValue(user.nome);
      sheet.getRange(rowNumber, 5).setValue(true);
    } else {
      sheet.appendRow([key, user.nome, user.salt, user.senhaHash, true, new Date()]);
    }
  });
}

function syncMainSpreadsheet_(ss, principal) {
  const sheet = ss.getSheetByName(CONFIG.SHEETS.SPREADSHEETS);
  if (sheet.getLastRow() >= 2) sheet.getRange(2, 4, sheet.getLastRow() - 1, 1).setValue(false);
  let rowNumber = null;
  if (sheet.getLastRow() >= 2) {
    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
    for (let i = 0; i < rows.length; i++) {
      if (String(rows[i][0]) === MAIN_SPREADSHEET_ID || extractSpreadsheetId_(rows[i][2]) === MAIN_SPREADSHEET_ID) { rowNumber = i + 2; break; }
    }
  }
  const values = [[MAIN_SPREADSHEET_ID, principal.nome, principal.url, true, 1]];
  if (rowNumber) sheet.getRange(rowNumber, 1, 1, 5).setValues(values);
  else sheet.appendRow(values[0]);
}

function getDb_() { return openConfiguredSpreadsheet_(CONTROL_SPREADSHEET_PROPERTY, 'controle'); }
function getAuditDb_() { return openConfiguredSpreadsheet_(AUDIT_SPREADSHEET_PROPERTY, 'auditoria'); }

function openConfiguredSpreadsheet_(propertyName, label) {
  const id = PropertiesService.getScriptProperties().getProperty(propertyName);
  if (!id) throw new Error('Sistema não configurado. Execute configurarNovaConta() ou setupSistema().');
  try { return SpreadsheetApp.openById(id); }
  catch (error) { throw new Error('A planilha de ' + label + ' não está acessível. Execute configurarNovaConta().'); }
}

function getOrCreateSpreadsheet_(propertyName, title) {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty(propertyName);
  if (id) {
    try { return SpreadsheetApp.openById(id); }
    catch (error) { props.deleteProperty(propertyName); }
  }
  const ss = SpreadsheetApp.create(title);
  try { ss.setSpreadsheetTimeZone(CONFIG.TIMEZONE); } catch (error) {}
  try { ss.setSpreadsheetLocale(CONFIG.LOCALE); } catch (error) {}
  props.setProperty(propertyName, ss.getId());
  return ss;
}

function ensureSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
  return sheet;
}

function cleanupBlankDefaultSheets_(ss) {
  const protectedNames = Object.keys(CONFIG.SHEETS).map(function(key) { return CONFIG.SHEETS[key]; });
  ss.getSheets().slice().forEach(function(sheet) {
    if (protectedNames.indexOf(sheet.getName()) !== -1 || ss.getSheets().length <= 1) return;
    const value = String(sheet.getRange(1, 1).getValue() || '').trim();
    if (sheet.getLastRow() <= 1 && sheet.getLastColumn() <= 1 && value === '') ss.deleteSheet(sheet);
  });
}

function countConfiguredUsers_(ss) {
  const sheet = ss.getSheetByName(CONFIG.SHEETS.USERS);
  return !sheet || sheet.getLastRow() < 2 ? 0 : sheet.getLastRow() - 1;
}

function extractSpreadsheetId_(url) {
  const match = String(url || '').match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  return match ? match[1] : String(url || '').trim();
}

function hashPassword_(password, salt) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(salt) + '|' + String(password), Utilities.Charset.UTF_8);
  return bytes.map(function(b) { return (b < 0 ? b + 256 : b).toString(16).padStart(2, '0'); }).join('');
}

function formatDuration_(totalSeconds) {
  totalSeconds = Math.max(0, Math.floor(Number(totalSeconds || 0)));
  const h = String(Math.floor(totalSeconds / 3600)).padStart(2, '0');
  const m = String(Math.floor((totalSeconds % 3600) / 60)).padStart(2, '0');
  const s = String(totalSeconds % 60).padStart(2, '0');
  return h + ':' + m + ':' + s;
}

function browserName_(userAgent) {
  const ua = String(userAgent || '');
  if (/Edg\//.test(ua)) return 'Microsoft Edge';
  if (/OPR\//.test(ua)) return 'Opera';
  if (/Chrome\//.test(ua)) return 'Google Chrome';
  if (/Firefox\//.test(ua)) return 'Firefox';
  if (/Safari\//.test(ua) && !/Chrome\//.test(ua)) return 'Safari';
  return 'Navegador';
}

function normalizeUser_(value) { return String(value || '').trim().toLowerCase().replace(/\s+/g, ''); }
function toBoolean_(value) { return value === true || String(value).toLowerCase() === 'true' || String(value) === '1'; }
function toDate_(value) { return value instanceof Date ? value : new Date(value); }
