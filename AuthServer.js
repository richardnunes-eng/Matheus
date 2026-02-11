/**
 * Backend de autenticacao para Google Apps Script (HTMLService).
 * Fluxo: login -> getMe -> refresh -> logout
 */
var AUTH_CONFIG = {
  spreadsheetId: "", // opcional: deixe vazio para usar planilha ativa
  spreadsheetPropertyKey: "AUTH_SPREADSHEET_ID",
  simpleUsersSheet: "SIMPLE_USERS",
  usersSheet: "USERS",
  sessionsSheet: "SESSIONS",
  attemptsSheet: "LOGIN_ATTEMPTS",
  pbkdf2Iterations: 35000,
  keyLength: 32,
  sessionTtlMs: 2 * 60 * 60 * 1000,
  refreshTtlMs: 30 * 24 * 60 * 60 * 1000,
  rateWindowMs: 10 * 60 * 1000,
  maxAttempts: 5,
  blockMs: 10 * 60 * 1000,
  secretPropertyKey: "AUTH_PEPPER_SECRET",
  defaultPlan: "FREE"
};

var AUTH_EXECUTION_USER_ID = "";

function authCreateAccount(username, password, displayName, clientMeta) {
  try {
    ensureAuthSheets();
    var normalized = normalizeUserIdentifier_(username);
    var safeDisplayName = String(displayName || "").trim();
    var meta = normalizeClientMeta_(clientMeta);

    if (!isValidUsername_(normalized)) {
      return authError_("Usuario invalido. Use email valido ou username (3+ caracteres).", "INVALID_USERNAME");
    }
    if (!isValidPassword_(password)) {
      return authError_("Senha fraca. Use pelo menos 8 caracteres.", "WEAK_PASSWORD");
    }

    var existing = findSimpleUserByIdentifier_(normalized);
    if (existing) {
      return authError_("Usuario ja cadastrado.", "USER_ALREADY_EXISTS");
    }

    var usersSheet = getSheet_(AUTH_CONFIG.simpleUsersSheet, simpleUserHeaders_());
    usersSheet.appendRow([
      Utilities.getUuid(),
      normalized,
      password,
      "USER",
      safeDisplayName || normalized,
      "ACTIVE",
      AUTH_CONFIG.defaultPlan,
      isoNow_(),
      ""
    ]);

    logAuthEvent_("ACCOUNT_CREATED", "", normalized, meta.ip);
    return { ok: true, message: "Conta criada com sucesso." };
  } catch (err) {
    logAuthEvent_("ACCOUNT_CREATE_ERROR", "", String(username || ""), "", String(err));
    return authError_("Erro interno ao criar conta.", "ACCOUNT_CREATE_INTERNAL_ERROR", {
      detail: String(err).substring(0, 220)
    });
  }
}

function login(username, password, rememberMe, clientMeta) {
  try {
    var normalized = normalizeUserIdentifier_(username);
    if (!normalized || !password) {
      return authError_("Credenciais invalidas.", "INVALID_INPUT");
    }

    var meta = normalizeClientMeta_(clientMeta);
    var rateState = checkRateLimit_(normalized, meta.ip);
    if (rateState.blocked) {
      return authError_(
        "Muitas tentativas. Tente novamente em alguns minutos.",
        "TEMP_BLOCKED",
        { blockedUntil: rateState.blockUntil }
      );
    }

    var user = findSimpleUserByIdentifier_(normalized);
    var isSimpleUser = !!user;
    if (!user) {
      user = findUserByIdentifier_(normalized);
    }
    if (!user) {
      registerFailedAttempt_(normalized, meta.ip);
      logAuthEvent_("LOGIN_USER_NOT_FOUND", "", normalized, meta.ip);
      return authError_("Usuario inexistente.", "USER_NOT_FOUND");
    }

    if (String(user.status || "ACTIVE").toUpperCase() !== "ACTIVE") {
      logAuthEvent_("LOGIN_USER_BLOCKED", user.userId, normalized, meta.ip);
      return authError_("Usuario bloqueado. Contate o administrador.", "USER_BLOCKED");
    }

    var isValidPassword = false;
    if (isSimpleUser) {
      isValidPassword = String(user.password || "") === String(password || "");
    } else {
      var userIterations = Number(user.iterations || AUTH_CONFIG.pbkdf2Iterations);
      isValidPassword = verifyPassword_(password, user.salt, user.passwordHash, userIterations);
    }
    if (!isValidPassword) {
      registerFailedAttempt_(normalized, meta.ip);
      logAuthEvent_("LOGIN_INVALID_PASSWORD", user.userId, normalized, meta.ip);
      return authError_("Credenciais invalidas.", "INVALID_CREDENTIALS");
    }

    clearFailedAttempts_(normalized, meta.ip);
    var tokens = createSession_(user, !!rememberMe, meta);
    if (isSimpleUser) {
      updateSimpleUserLastLogin_(user.rowIndex);
    } else {
      updateUserLastLogin_(user.rowIndex);
    }
    logAuthEvent_("LOGIN_SUCCESS", user.userId, normalized, meta.ip);

    return {
      ok: true,
      sessionToken: tokens.sessionToken,
      refreshToken: tokens.refreshToken || "",
      userProfile: publicUserProfile_(user)
    };
  } catch (err) {
    logAuthEvent_("LOGIN_ERROR", "", String(username || ""), "", String(err));
    return authError_("Erro interno ao autenticar.", "LOGIN_INTERNAL_ERROR");
  }
}

function refresh(refreshToken, clientMeta) {
  try {
    if (!refreshToken) {
      return authError_("Refresh token ausente.", "INVALID_REFRESH_TOKEN");
    }

    var meta = normalizeClientMeta_(clientMeta);
    var now = Date.now();
    var sessionsSheet = getSheet_(AUTH_CONFIG.sessionsSheet, sessionHeaders_());
    var sessions = mapRowsByHeader_(sessionsSheet);
    var refreshTokenHash = hashToken_(refreshToken);
    var matched = null;

    for (var i = 0; i < sessions.length; i++) {
      var row = sessions[i];
      if (!row.refreshTokenHash) {
        continue;
      }
      if (constantTimeEqual_(row.refreshTokenHash, refreshTokenHash)) {
        matched = row;
        break;
      }
    }

    if (!matched) {
      return authError_("Sessao invalida.", "SESSION_INVALID");
    }

    if (Number(matched.revokedAtEpochMs || 0) > 0) {
      return authError_("Sessao invalida.", "SESSION_REVOKED");
    }

    if (Number(matched.refreshExpiresAtEpochMs || 0) < now) {
      revokeSessionByRow_(sessionsSheet, matched.rowIndex);
      return authError_("Sessao expirada. Faca login novamente.", "SESSION_EXPIRED");
    }

    var user = findUserById_(matched.userId);
    if (!user || String(user.status || "ACTIVE").toUpperCase() !== "ACTIVE") {
      revokeSessionByRow_(sessionsSheet, matched.rowIndex);
      return authError_("Usuario indisponivel.", "USER_UNAVAILABLE");
    }

    var rotated = rotateSessionTokens_(sessionsSheet, matched, user, meta);
    logAuthEvent_("REFRESH_SUCCESS", user.userId, user.username, meta.ip);

    return {
      ok: true,
      sessionToken: rotated.sessionToken,
      userProfile: publicUserProfile_(user)
    };
  } catch (err) {
    logAuthEvent_("REFRESH_ERROR", "", "", "", String(err));
    return authError_("Falha no refresh de sessao.", "REFRESH_INTERNAL_ERROR");
  }
}

function logout(sessionToken) {
  try {
    if (!sessionToken) {
      return authError_("Token de sessao ausente.", "INVALID_SESSION_TOKEN");
    }

    var session = findSessionBySessionToken_(sessionToken);
    if (!session) {
      return { ok: true };
    }

    var sessionsSheet = getSheet_(AUTH_CONFIG.sessionsSheet, sessionHeaders_());
    revokeSessionByRow_(sessionsSheet, session.rowIndex);
    logAuthEvent_("LOGOUT_SUCCESS", session.userId, "", "");
    return { ok: true };
  } catch (err) {
    logAuthEvent_("LOGOUT_ERROR", "", "", "", String(err));
    return authError_("Erro ao encerrar sessao.", "LOGOUT_INTERNAL_ERROR");
  }
}

function getMe(sessionToken) {
  try {
    if (!sessionToken) {
      return authError_("Token de sessao ausente.", "INVALID_SESSION_TOKEN");
    }

    var session = findSessionBySessionToken_(sessionToken);
    if (!session) {
      return authError_("Sessao invalida.", "SESSION_INVALID");
    }

    var now = Date.now();
    if (Number(session.revokedAtEpochMs || 0) > 0) {
      return authError_("Sessao revogada.", "SESSION_REVOKED");
    }
    if (Number(session.expiresAtEpochMs || 0) < now) {
      var sheet = getSheet_(AUTH_CONFIG.sessionsSheet, sessionHeaders_());
      revokeSessionByRow_(sheet, session.rowIndex);
      return authError_("Sessao expirada.", "SESSION_EXPIRED");
    }

    var user = findUserById_(session.userId);
    if (!user || String(user.status || "ACTIVE").toUpperCase() !== "ACTIVE") {
      return authError_("Usuario indisponivel.", "USER_UNAVAILABLE");
    }

    return { ok: true, userProfile: publicUserProfile_(user) };
  } catch (err) {
    logAuthEvent_("GET_ME_ERROR", "", "", "", String(err));
    return authError_("Erro ao validar sessao.", "GET_ME_INTERNAL_ERROR");
  }
}

function seedAdmin() {
  return seedSimpleUser("admin", "admin123", "ADMIN", "Administrador", AUTH_CONFIG.defaultPlan);
}

function seedSimpleUser(username, password, role, displayName, plan) {
  ensureAuthSheets();
  var normalized = normalizeUserIdentifier_(username);
  if (!isValidUsername_(normalized)) {
    return authError_("Usuario invalido.", "INVALID_USERNAME");
  }
  if (!isValidPassword_(password)) {
    return authError_("Senha invalida (minimo 8 caracteres).", "INVALID_PASSWORD");
  }

  var usersSheet = getSheet_(AUTH_CONFIG.simpleUsersSheet, simpleUserHeaders_());
  var users = mapRowsByHeader_(usersSheet);
  var existingRow = null;
  for (var i = 0; i < users.length; i++) {
    if (normalizeUserIdentifier_(users[i].username) === normalized) {
      existingRow = users[i];
      break;
    }
  }

  var safeRole = String(role || "USER").toUpperCase();
  var safeName = String(displayName || normalized).trim();
  var safePlan = String(plan || AUTH_CONFIG.defaultPlan).trim();

  if (!existingRow) {
    usersSheet.appendRow([
      Utilities.getUuid(),
      normalized,
      String(password),
      safeRole,
      safeName,
      "ACTIVE",
      safePlan,
      isoNow_(),
      ""
    ]);
  } else {
    var map = buildHeaderMap_(simpleUserHeaders_());
    usersSheet.getRange(existingRow.rowIndex, map.password + 1).setValue(String(password));
    usersSheet.getRange(existingRow.rowIndex, map.role + 1).setValue(safeRole);
    usersSheet.getRange(existingRow.rowIndex, map.displayName + 1).setValue(safeName);
    usersSheet.getRange(existingRow.rowIndex, map.status + 1).setValue("ACTIVE");
    usersSheet.getRange(existingRow.rowIndex, map.plan + 1).setValue(safePlan);
  }

  logAuthEvent_("SIMPLE_USER_SAVED", "", normalized, "");
  return { ok: true, username: normalized, password: String(password), role: safeRole };
}

function ensureAuthSheets() {
  try {
    var ss = getAuthSpreadsheet_();
    getSheet_(AUTH_CONFIG.simpleUsersSheet, simpleUserHeaders_());
    getSheet_(AUTH_CONFIG.usersSheet, userHeaders_());
    getSheet_(AUTH_CONFIG.sessionsSheet, sessionHeaders_());
    getSheet_(AUTH_CONFIG.attemptsSheet, attemptsHeaders_());
    ensurePepperSecret_();
    return {
      ok: true,
      spreadsheetId: ss.getId(),
      spreadsheetName: ss.getName(),
      spreadsheetUrl: ss.getUrl()
    };
  } catch (err) {
    return {
      ok: false,
      message: "Falha ao preparar abas de autenticacao.",
      detail: String(err)
    };
  }
}

function getAuthStatus() {
  var simpleUsers = mapRowsByHeader_(getSheet_(AUTH_CONFIG.simpleUsersSheet, simpleUserHeaders_()));
  var users = mapRowsByHeader_(getSheet_(AUTH_CONFIG.usersSheet, userHeaders_()));
  var sessions = mapRowsByHeader_(getSheet_(AUTH_CONFIG.sessionsSheet, sessionHeaders_()));
  return {
    ok: true,
    simpleUsers: simpleUsers.length,
    users: users.length,
    sessions: sessions.length
  };
}

function createSession_(user, rememberMe, meta) {
  var sessionsSheet = getSheet_(AUTH_CONFIG.sessionsSheet, sessionHeaders_());
  var now = Date.now();
  var sessionToken = generateToken_();
  var refreshToken = rememberMe ? generateToken_() : "";

  sessionsSheet.appendRow([
    Utilities.getUuid(),
    user.userId,
    hashToken_(sessionToken),
    hashToken_(refreshToken),
    String(now + AUTH_CONFIG.sessionTtlMs),
    String(now + AUTH_CONFIG.refreshTtlMs),
    isoNow_(),
    String(meta.userAgent || ""),
    String(meta.ip || ""),
    String(rememberMe ? 1 : 0),
    "",
    "",
    String(now + AUTH_CONFIG.sessionTtlMs)
  ]);

  return {
    sessionToken: sessionToken,
    refreshToken: refreshToken
  };
}

function rotateSessionTokens_(sheet, sessionRow, user, meta) {
  var now = Date.now();
  var newSessionToken = generateToken_();
  var expiresAtMs = now + AUTH_CONFIG.sessionTtlMs;
  var idx = sessionRow.rowIndex;
  var headers = sessionHeaders_();
  var headerMap = buildHeaderMap_(headers);

  sheet.getRange(idx, headerMap.sessionTokenHash + 1).setValue(hashToken_(newSessionToken));
  sheet.getRange(idx, headerMap.expiresAtEpochMs + 1).setValue(String(expiresAtMs));
  sheet.getRange(idx, headerMap.lastSeenAt + 1).setValue(isoNow_());
  if (meta.userAgent) {
    sheet.getRange(idx, headerMap.userAgent + 1).setValue(meta.userAgent);
  }
  if (meta.ip) {
    sheet.getRange(idx, headerMap.ip + 1).setValue(meta.ip);
  }

  return { sessionToken: newSessionToken };
}

function findSessionBySessionToken_(sessionToken) {
  var sessionsSheet = getSheet_(AUTH_CONFIG.sessionsSheet, sessionHeaders_());
  var sessions = mapRowsByHeader_(sessionsSheet);
  var sessionTokenHash = hashToken_(sessionToken);
  for (var i = 0; i < sessions.length; i++) {
    if (constantTimeEqual_(sessions[i].sessionTokenHash, sessionTokenHash)) {
      return sessions[i];
    }
  }
  return null;
}

function revokeSessionByRow_(sheet, rowIndex) {
  var headers = sessionHeaders_();
  var map = buildHeaderMap_(headers);
  var nowIso = isoNow_();
  var nowMs = String(Date.now());
  sheet.getRange(rowIndex, map.revokedAt + 1).setValue(nowIso);
  sheet.getRange(rowIndex, map.revokedAtEpochMs + 1).setValue(nowMs);
}

function findUserByIdentifier_(username) {
  var usersSheet = getSheet_(AUTH_CONFIG.usersSheet, userHeaders_());
  var users = mapRowsByHeader_(usersSheet);
  for (var i = 0; i < users.length; i++) {
    if (normalizeUserIdentifier_(users[i].username) === username) {
      return users[i];
    }
  }
  return null;
}

function findUserById_(userId) {
  var simpleUser = findSimpleUserById_(userId);
  if (simpleUser) {
    return simpleUser;
  }

  var usersSheet = getSheet_(AUTH_CONFIG.usersSheet, userHeaders_());
  var users = mapRowsByHeader_(usersSheet);
  for (var i = 0; i < users.length; i++) {
    if (String(users[i].userId) === String(userId)) {
      return users[i];
    }
  }
  return null;
}

function updateUserLastLogin_(rowIndex) {
  var map = buildHeaderMap_(userHeaders_());
  var usersSheet = getSheet_(AUTH_CONFIG.usersSheet, userHeaders_());
  usersSheet.getRange(rowIndex, map.lastLoginAt + 1).setValue(isoNow_());
}

function findSimpleUserByIdentifier_(username) {
  var usersSheet = getSheet_(AUTH_CONFIG.simpleUsersSheet, simpleUserHeaders_());
  var users = mapRowsByHeader_(usersSheet);
  for (var i = 0; i < users.length; i++) {
    if (normalizeUserIdentifier_(users[i].username) === username) {
      return users[i];
    }
  }
  return null;
}

function findSimpleUserById_(userId) {
  var usersSheet = getSheet_(AUTH_CONFIG.simpleUsersSheet, simpleUserHeaders_());
  var users = mapRowsByHeader_(usersSheet);
  for (var i = 0; i < users.length; i++) {
    if (String(users[i].userId) === String(userId)) {
      return users[i];
    }
  }
  return null;
}

function updateSimpleUserLastLogin_(rowIndex) {
  var map = buildHeaderMap_(simpleUserHeaders_());
  var usersSheet = getSheet_(AUTH_CONFIG.simpleUsersSheet, simpleUserHeaders_());
  usersSheet.getRange(rowIndex, map.lastLoginAt + 1).setValue(isoNow_());
}

function checkRateLimit_(username, ip) {
  var attemptsSheet = getSheet_(AUTH_CONFIG.attemptsSheet, attemptsHeaders_());
  var attempts = mapRowsByHeader_(attemptsSheet);
  var key = rateKey_(username, ip);
  var now = Date.now();
  var found = null;

  for (var i = 0; i < attempts.length; i++) {
    if (attempts[i].key === key) {
      found = attempts[i];
      break;
    }
  }

  if (!found) {
    return { blocked: false };
  }

  var blockUntil = Number(found.blockUntilEpochMs || 0);
  if (blockUntil > now) {
    return { blocked: true, blockUntil: new Date(blockUntil).toISOString() };
  }

  return { blocked: false };
}

function registerFailedAttempt_(username, ip) {
  var attemptsSheet = getSheet_(AUTH_CONFIG.attemptsSheet, attemptsHeaders_());
  var attempts = mapRowsByHeader_(attemptsSheet);
  var key = rateKey_(username, ip);
  var now = Date.now();
  var map = buildHeaderMap_(attemptsHeaders_());
  var found = null;

  for (var i = 0; i < attempts.length; i++) {
    if (attempts[i].key === key) {
      found = attempts[i];
      break;
    }
  }

  if (!found) {
    attemptsSheet.appendRow([
      key,
      username,
      ip,
      "1",
      String(now),
      "0",
      String(now)
    ]);
    return;
  }

  var firstAttempt = Number(found.firstAttemptAtEpochMs || 0);
  var failedCount = Number(found.failedCount || 0);
  if (firstAttempt === 0 || now - firstAttempt > AUTH_CONFIG.rateWindowMs) {
    failedCount = 0;
    firstAttempt = now;
  }
  failedCount += 1;

  var blockUntil = 0;
  if (failedCount >= AUTH_CONFIG.maxAttempts) {
    blockUntil = now + AUTH_CONFIG.blockMs;
  }

  attemptsSheet.getRange(found.rowIndex, map.failedCount + 1).setValue(String(failedCount));
  attemptsSheet.getRange(found.rowIndex, map.firstAttemptAtEpochMs + 1).setValue(String(firstAttempt));
  attemptsSheet.getRange(found.rowIndex, map.blockUntilEpochMs + 1).setValue(String(blockUntil));
  attemptsSheet.getRange(found.rowIndex, map.lastAttemptAtEpochMs + 1).setValue(String(now));
}

function clearFailedAttempts_(username, ip) {
  var attemptsSheet = getSheet_(AUTH_CONFIG.attemptsSheet, attemptsHeaders_());
  var attempts = mapRowsByHeader_(attemptsSheet);
  var map = buildHeaderMap_(attemptsHeaders_());
  var key = rateKey_(username, ip);

  for (var i = 0; i < attempts.length; i++) {
    if (attempts[i].key !== key) {
      continue;
    }
    attemptsSheet.getRange(attempts[i].rowIndex, map.failedCount + 1).setValue("0");
    attemptsSheet.getRange(attempts[i].rowIndex, map.firstAttemptAtEpochMs + 1).setValue("0");
    attemptsSheet.getRange(attempts[i].rowIndex, map.blockUntilEpochMs + 1).setValue("0");
    attemptsSheet.getRange(attempts[i].rowIndex, map.lastAttemptAtEpochMs + 1).setValue(String(Date.now()));
    break;
  }
}

function hashPassword_(password) {
  var salt = base64UrlFromBytes_(generateRandomBytes_(16));
  var hash = pbkdf2Sha256_(password, salt, AUTH_CONFIG.pbkdf2Iterations, AUTH_CONFIG.keyLength);
  return { salt: salt, hash: hash };
}

function verifyPassword_(password, salt, expectedHash, iterations) {
  var safeIterations = Number(iterations || AUTH_CONFIG.pbkdf2Iterations);
  var hash = pbkdf2Sha256_(password, salt, safeIterations, AUTH_CONFIG.keyLength);
  return constantTimeEqual_(hash, expectedHash);
}

function pbkdf2Sha256_(password, salt, iterations, keyLength) {
  var dk = [];
  var blockIndex = 1;
  var saltBytes = Utilities.newBlob(salt).getBytes();

  while (dk.length < keyLength) {
    var u = hmacSha256_(password, saltBytes.concat(int32be_(blockIndex)));
    var t = u.slice(0);
    for (var i = 1; i < iterations; i++) {
      u = hmacSha256_(password, u);
      t = xorBytes_(t, u);
    }
    dk = dk.concat(t);
    blockIndex++;
  }

  return base64UrlFromBytes_(dk.slice(0, keyLength));
}

function hmacSha256_(key, messageBytes) {
  var keyBytes = toSignedByteArray_(Utilities.newBlob(String(key)).getBytes());
  var valueBytes = toSignedByteArray_(messageBytes);
  return Utilities.computeHmacSha256Signature(valueBytes, keyBytes);
}

function int32be_(value) {
  return [
    (value >>> 24) & 0xff,
    (value >>> 16) & 0xff,
    (value >>> 8) & 0xff,
    value & 0xff
  ];
}

function xorBytes_(a, b) {
  var out = [];
  for (var i = 0; i < a.length; i++) {
    out.push((a[i] ^ b[i]) & 0xff);
  }
  return out;
}

function hashToken_(token) {
  if (!token) {
    return "";
  }
  var pepper = ensurePepperSecret_();
  var input = token + "." + pepper;
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, input);
  return base64UrlFromBytes_(digest);
}

function generateToken_() {
  return base64UrlFromBytes_(generateRandomBytes_(48));
}

function ensurePepperSecret_() {
  var props = PropertiesService.getScriptProperties();
  var secret = props.getProperty(AUTH_CONFIG.secretPropertyKey);
  if (secret) {
    return secret;
  }
  secret = base64UrlFromBytes_(generateRandomBytes_(32));
  props.setProperty(AUTH_CONFIG.secretPropertyKey, secret);
  return secret;
}

function generateRandomBytes_(length) {
  var out = [];
  var counter = 0;
  while (out.length < length) {
    var seed = [
      Utilities.getUuid(),
      String(new Date().getTime()),
      String(Math.random()),
      String(counter++)
    ].join("|");
    var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, seed);
    for (var i = 0; i < digest.length && out.length < length; i++) {
      out.push((digest[i] + 256) % 256);
    }
  }
  return out;
}

function toSignedByteArray_(bytes) {
  var out = [];
  for (var i = 0; i < bytes.length; i++) {
    var n = Number(bytes[i]);
    var unsigned = ((n % 256) + 256) % 256;
    out.push(unsigned > 127 ? unsigned - 256 : unsigned);
  }
  return out;
}

function publicUserProfile_(user) {
  return {
    userId: user.userId,
    username: user.username,
    name: user.displayName || user.username,
    role: user.role || "USER",
    plan: user.plan || AUTH_CONFIG.defaultPlan
  };
}

function constantTimeEqual_(a, b) {
  var aa = String(a || "");
  var bb = String(b || "");
  if (aa.length !== bb.length) {
    return false;
  }
  var out = 0;
  for (var i = 0; i < aa.length; i++) {
    out |= aa.charCodeAt(i) ^ bb.charCodeAt(i);
  }
  return out === 0;
}

function normalizeUserIdentifier_(value) {
  return String(value || "").trim().toLowerCase();
}

function isValidUsername_(username) {
  var value = String(username || "").trim();
  if (!value || value.length < 3 || value.length > 120) {
    return false;
  }
  return true;
}

function isValidPassword_(password) {
  var value = String(password || "");
  if (value.length < 8 || value.length > 128) {
    return false;
  }
  return true;
}

function normalizeClientMeta_(meta) {
  meta = meta || {};
  return {
    userAgent: String(meta.userAgent || "").substring(0, 300),
    ip: String(meta.ip || "unknown").substring(0, 80)
  };
}

function rateKey_(username, ip) {
  return username + "|" + (ip || "unknown");
}

function authError_(message, code, extra) {
  var payload = { ok: false, message: message, code: code };
  if (extra) {
    payload.extra = extra;
  }
  return payload;
}

function logAuthEvent_(eventName, userId, username, ip, detail) {
  var base = "[AUTH] " + eventName +
    " userId=" + String(userId || "-") +
    " login=" + String(username || "-") +
    " ip=" + String(ip || "-");
  if (detail) {
    base += " detail=" + String(detail).substring(0, 160);
  }
  Logger.log(base);
}

function getSheet_(sheetName, headers) {
  var ss = getAuthSpreadsheet_();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
  } else {
    var headerRange = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
    var mismatch = false;
    for (var i = 0; i < headers.length; i++) {
      if (String(headerRange[i] || "") !== headers[i]) {
        mismatch = true;
        break;
      }
    }
    if (mismatch) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
  }
  return sheet;
}

function getAuthSpreadsheet_() {
  var props = PropertiesService.getScriptProperties();
  var configuredId = String(AUTH_CONFIG.spreadsheetId || "").trim();
  var savedId = String(props.getProperty(AUTH_CONFIG.spreadsheetPropertyKey) || "").trim();
  var targetId = configuredId || savedId;

  if (targetId) {
    try {
      return SpreadsheetApp.openById(targetId);
    } catch (err) {
      props.deleteProperty(AUTH_CONFIG.spreadsheetPropertyKey);
      Logger.log("[AUTH] AUTH_SPREADSHEET_ID invalido, criando nova planilha. detail=" + String(err));
    }
  }

  var active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) {
    props.setProperty(AUTH_CONFIG.spreadsheetPropertyKey, active.getId());
    return active;
  }

  var created = SpreadsheetApp.create("Sistema Matheus - Auth");
  props.setProperty(AUTH_CONFIG.spreadsheetPropertyKey, created.getId());
  return created;
}

function setAuthSpreadsheetId(spreadsheetId) {
  var id = String(spreadsheetId || "").trim();
  if (!id) {
    return authError_("Informe um spreadsheetId valido.", "INVALID_SPREADSHEET_ID");
  }
  var ss = SpreadsheetApp.openById(id);
  PropertiesService.getScriptProperties().setProperty(AUTH_CONFIG.spreadsheetPropertyKey, ss.getId());
  return { ok: true, spreadsheetId: ss.getId(), spreadsheetName: ss.getName(), spreadsheetUrl: ss.getUrl() };
}

function authRpc(functionName, arg1, arg2, sessionToken) {
  try {
    if (!functionName) {
      return authError_("Funcao nao informada.", "INVALID_FUNCTION");
    }

    // Compatibilidade:
    // - legado: authRpc(fn, arg1, arg2, sessionToken)
    // - novo:   authRpc(fn, ...args, sessionToken)
    var rawArgs = Array.prototype.slice.call(arguments);
    var token = rawArgs.length ? rawArgs[rawArgs.length - 1] : "";
    var fnArgs = rawArgs.length > 2 ? rawArgs.slice(1, rawArgs.length - 1) : [];

    if (!token) {
      return authError_("Sessao ausente.", "AUTH_REQUIRED");
    }

    var session = findSessionBySessionToken_(token);
    if (!session) {
      return authError_("Sessao invalida.", "AUTH_REQUIRED");
    }

    var now = Date.now();
    if (Number(session.revokedAtEpochMs || 0) > 0 || Number(session.expiresAtEpochMs || 0) < now) {
      return authError_("Sessao expirada.", "AUTH_REQUIRED");
    }

    var user = findUserById_(session.userId);
    if (!user || String(user.status || "ACTIVE").toUpperCase() !== "ACTIVE") {
      return authError_("Usuario invalido.", "AUTH_REQUIRED");
    }

    // Escopo da execucao atual: Sheets.js usa esse userId para separar os dados.
    AUTH_EXECUTION_USER_ID = String(user.userId);

    var fn = this[functionName];
    if (typeof fn !== "function") {
      return authError_("Endpoint nao encontrado: " + functionName, "FUNCTION_NOT_FOUND");
    }
    if (functionName === "authRpc") {
      return authError_("Chamada invalida.", "INVALID_FUNCTION");
    }

    var result = fn.apply(this, fnArgs);
    return result;
  } catch (err) {
    return authError_("Falha de autenticacao no RPC.", "AUTH_RPC_ERROR", {
      detail: String(err).substring(0, 220)
    });
  } finally {
    AUTH_EXECUTION_USER_ID = "";
  }
}

function debugUserScope(sessionToken) {
  try {
    if (!sessionToken) {
      return authError_("Sessao ausente.", "AUTH_REQUIRED");
    }

    var session = findSessionBySessionToken_(sessionToken);
    if (!session) {
      return authError_("Sessao invalida.", "AUTH_REQUIRED");
    }

    var user = findUserById_(session.userId);
    if (!user) {
      return authError_("Usuario nao encontrado.", "AUTH_REQUIRED");
    }

    AUTH_EXECUTION_USER_ID = String(user.userId);
    var db = getDbSpreadsheet();
    ensureAllSheets();

    var contas = getSheetDataAsObjects(SHEET_CONTAS);
    var tx = getSheetDataAsObjects(SHEET_LANCAMENTOS);
    var categoriasPriv = getSheetDataAsObjects(SHEET_CATEGORIAS).filter(function (c) {
      return String(c.OwnerUserId || '') === String(user.userId);
    });
    var categoriasList = getCategoriesCached(true);
    var categoriasPub = categoriasList.filter(function (c) { return !!c.isPublic; });

    return {
      ok: true,
      userId: user.userId,
      username: user.username,
      dbSpreadsheetId: db.getId(),
      dbSpreadsheetName: db.getName(),
      counts: {
        contas: contas.length,
        lancamentos: tx.length,
        categoriasPublicasVisiveis: categoriasPub.length,
        categoriasPrivadas: categoriasPriv.length
      },
      contasSample: contas.slice(0, 10).map(function (a) {
        return { id: a.ID, nome: a.Nome, tipo: a.Tipo, instituicao: a.Instituicao || 'Outro', ativo: a.Ativo };
      })
    };
  } catch (err) {
    return authError_("Falha no debug.", "DEBUG_ERROR", { detail: String(err).substring(0, 240) });
  } finally {
    AUTH_EXECUTION_USER_ID = "";
  }
}

function syncUserBankAccounts(sessionToken) {
  try {
    if (!sessionToken) return authError_("Sessao ausente.", "AUTH_REQUIRED");
    var session = findSessionBySessionToken_(sessionToken);
    if (!session) return authError_("Sessao invalida.", "AUTH_REQUIRED");
    var user = findUserById_(session.userId);
    if (!user) return authError_("Usuario nao encontrado.", "AUTH_REQUIRED");

    AUTH_EXECUTION_USER_ID = String(user.userId);
    var scoped = getDbSpreadsheet();
    ensureAllSheets();
    seedScopedAccountsFromGlobal_(scoped);
    deactivateGenericBankAccounts_(scoped.getSheetByName(SHEET_CONTAS));

    var contas = getSheetDataAsObjects(SHEET_CONTAS);
    return { ok: true, userId: user.userId, totalContas: contas.length };
  } catch (err) {
    return authError_("Falha ao sincronizar contas.", "SYNC_ACCOUNTS_ERROR", {
      detail: String(err).substring(0, 220)
    });
  } finally {
    AUTH_EXECUTION_USER_ID = "";
  }
}

function syncAllUsersBankAccounts() {
  var report = [];
  try {
    ensureAuthSheets();
    var users = mapRowsByHeader_(getSheet_(AUTH_CONFIG.simpleUsersSheet, simpleUserHeaders_()));
    for (var i = 0; i < users.length; i++) {
      var user = users[i];
      if (String(user.status || "ACTIVE").toUpperCase() !== "ACTIVE") continue;

      try {
        AUTH_EXECUTION_USER_ID = String(user.userId);
        var scoped = getDbSpreadsheet();
        ensureAllSheets();
        seedScopedAccountsFromGlobal_(scoped);
        deactivateGenericBankAccounts_(scoped.getSheetByName(SHEET_CONTAS));
        var contas = getSheetDataAsObjects(SHEET_CONTAS);
        report.push({
          userId: user.userId,
          username: user.username,
          totalContas: contas.length,
          dbSpreadsheetId: scoped.getId()
        });
      } finally {
        AUTH_EXECUTION_USER_ID = "";
      }
    }
    return { ok: true, synced: report.length, report: report };
  } catch (err) {
    AUTH_EXECUTION_USER_ID = "";
    return authError_("Falha na sincronizacao global de contas.", "SYNC_ALL_ACCOUNTS_ERROR", {
      detail: String(err).substring(0, 220),
      partialReport: report
    });
  }
}

function mapRowsByHeader_(sheet) {
  var values = sheet.getDataRange().getValues();
  if (values.length <= 1) {
    return [];
  }
  var headers = values[0];
  var rows = [];
  for (var i = 1; i < values.length; i++) {
    var row = {};
    for (var j = 0; j < headers.length; j++) {
      row[String(headers[j] || "")] = values[i][j];
    }
    row.rowIndex = i + 1;
    rows.push(row);
  }
  return rows;
}

function buildHeaderMap_(headers) {
  var map = {};
  for (var i = 0; i < headers.length; i++) {
    map[headers[i]] = i;
  }
  return map;
}

function userHeaders_() {
  return [
    "userId",
    "username",
    "passwordHash",
    "salt",
    "iterations",
    "role",
    "displayName",
    "status",
    "plan",
    "createdAt",
    "lastLoginAt"
  ];
}

function simpleUserHeaders_() {
  return [
    "userId",
    "username",
    "password",
    "role",
    "displayName",
    "status",
    "plan",
    "createdAt",
    "lastLoginAt"
  ];
}

function sessionHeaders_() {
  return [
    "sessionId",
    "userId",
    "sessionTokenHash",
    "refreshTokenHash",
    "expiresAtEpochMs",
    "refreshExpiresAtEpochMs",
    "createdAt",
    "userAgent",
    "ip",
    "rememberMe",
    "revokedAt",
    "revokedAtEpochMs",
    "lastSeenAt"
  ];
}

function attemptsHeaders_() {
  return [
    "key",
    "username",
    "ip",
    "failedCount",
    "firstAttemptAtEpochMs",
    "blockUntilEpochMs",
    "lastAttemptAtEpochMs"
  ];
}

function isoNow_() {
  return new Date().toISOString();
}

function base64UrlFromBytes_(bytes) {
  var safeBytes = [];
  for (var i = 0; i < bytes.length; i++) {
    safeBytes.push((bytes[i] + 256) % 256);
  }
  return Utilities.base64EncodeWebSafe(safeBytes).replace(/=+$/, "");
}
