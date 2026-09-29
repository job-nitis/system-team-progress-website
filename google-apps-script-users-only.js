// Use this script only for System Website User permissions.
// It reads the "Users" sheet tab and returns one active user's role by email.
const USERS_SHEET_NAME = "Users";
const AUDIT_SHEET_NAME = "Audit Log";
const USERS_HEADERS = [
  "Email",
  "Name",
  "Role",
  "Active",
  "Password",
  "Last Login"
];
const AUDIT_HEADERS = [
  "Time",
  "Email",
  "Role",
  "Action",
  "Module",
  "Detail"
];
const SESSION_TTL_SECONDS = 15 * 60;
const SESSION_PROPERTY_PREFIX = "SYSTEM_TEAM_SESSION_";

function authorizeUserPermissions() {
  const sheet = getUsersSheet();
  sheet.getName();
  const auditSheet = getAuditSheet();
  auditSheet.getName();

  return "User permissions are authorized";
}

function doGet(e) {
  return handleUserRequest(e);
}

function doPost(e) {
  return handleUserRequest(e);
}

function handleUserRequest(e) {
  const callback = e && e.parameter && e.parameter.callback ? String(e.parameter.callback) : "";
  const action = e && e.parameter && e.parameter.action ? String(e.parameter.action) : "";
  const email = e && e.parameter && e.parameter.email ? String(e.parameter.email).trim().toLowerCase() : "";
  const password = e && e.parameter && e.parameter.password ? String(e.parameter.password) : "";
  const token = e && e.parameter && e.parameter.token ? String(e.parameter.token) : "";

  try {
    if (action === "session") {
      const session = validateSession(token, true);
      if (!session) return callbackResponse(callback, authenticationRequired());
      return callbackResponse(callback, {
        ok: true,
        user: publicUser(session.user),
        expiresAt: session.expiresAt,
        idleTimeoutMinutes: SESSION_TTL_SECONDS / 60
      });
    }

    if (action === "logout") {
      const session = validateSession(token, false);
      if (session) {
        appendAudit({
          email: session.user.email,
          role: session.user.role,
          action: "Logout",
          module: "User",
          detail: e && e.parameter && e.parameter.detail ? String(e.parameter.detail) : "Manual logout"
        });
      }
      deleteSession(token);
      return callbackResponse(callback, { ok: true });
    }

    if (action === "audit") {
      const session = validateSession(token, true);
      if (!session) return callbackResponse(callback, authenticationRequired());
      appendAudit({
        email: session.user.email,
        role: session.user.role,
        action: e && e.parameter && e.parameter.event ? String(e.parameter.event) : "",
        module: e && e.parameter && e.parameter.module ? String(e.parameter.module) : "",
        detail: e && e.parameter && e.parameter.detail ? String(e.parameter.detail) : ""
      });
      return callbackResponse(callback, { ok: true });
    }

    if (!email) {
      return callbackResponse(callback, { ok: false, error: "Email is required" });
    }

    const result = findUserByEmail(email);
    if (!result) {
      return callbackResponse(callback, { ok: false, error: "Email is not approved" });
    }

    const user = result.user;
    if (!user.active) {
      return callbackResponse(callback, { ok: false, error: "User is inactive" });
    }

    if (!user.password) {
      appendAudit({
        email: user.email,
        role: user.role,
        action: "Login failed",
        module: "User",
        detail: "Password is not set"
      });
      return callbackResponse(callback, { ok: false, error: "Password is not set for this user" });
    }

    if (user.password !== password) {
      appendAudit({
        email: user.email,
        role: user.role,
        action: "Login failed",
        module: "User",
        detail: "Wrong password"
      });
      return callbackResponse(callback, { ok: false, error: "Wrong password" });
    }

    if (action === "verify") {
      return callbackResponse(callback, { ok: true, user: publicUser(user) });
    }

    updateLastLogin(result.rowNumber);
    appendAudit({
      email: user.email,
      role: user.role,
      action: "Login",
      module: "User",
      detail: "Successful login"
    });

    const sessionToken = createSession(user);
    return callbackResponse(callback, {
      ok: true,
      user: publicUser(user),
      sessionToken,
      idleTimeoutMinutes: SESSION_TTL_SECONDS / 60
    });
  } catch (error) {
    return callbackResponse(callback, {
      ok: false,
      error: String(error && error.message ? error.message : error)
    });
  }
}

function authenticationRequired() {
  return {
    ok: false,
    code: "SESSION_EXPIRED",
    error: "Authentication required"
  };
}

function createSession(user) {
  cleanupExpiredSessions();
  const token = `${Utilities.getUuid().replace(/-/g, "")}${Utilities.getUuid().replace(/-/g, "")}`;
  const now = Date.now();
  const session = {
    email: user.email,
    createdAt: now,
    lastActivityAt: now,
    expiresAt: now + (SESSION_TTL_SECONDS * 1000)
  };

  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    PropertiesService.getScriptProperties().setProperty(sessionPropertyKey(token), JSON.stringify(session));
  } finally {
    lock.releaseLock();
  }

  return token;
}

function validateSession(token, touch) {
  const cleanToken = String(token || "").trim();
  if (!/^[a-f0-9]{64}$/i.test(cleanToken)) return null;

  const properties = PropertiesService.getScriptProperties();
  const key = sessionPropertyKey(cleanToken);
  const serialized = properties.getProperty(key);
  if (!serialized) return null;

  let session;
  try {
    session = JSON.parse(serialized);
  } catch (error) {
    properties.deleteProperty(key);
    return null;
  }

  const now = Date.now();
  if (!session.expiresAt || now >= Number(session.expiresAt)) {
    properties.deleteProperty(key);
    return null;
  }

  const result = findUserByEmail(String(session.email || "").trim().toLowerCase());
  if (!result || !result.user.active) {
    properties.deleteProperty(key);
    return null;
  }

  if (touch) {
    session.lastActivityAt = now;
    session.expiresAt = now + (SESSION_TTL_SECONDS * 1000);
    properties.setProperty(key, JSON.stringify(session));
  }

  return {
    user: result.user,
    expiresAt: Number(session.expiresAt)
  };
}

function deleteSession(token) {
  const cleanToken = String(token || "").trim();
  if (!cleanToken) return;
  PropertiesService.getScriptProperties().deleteProperty(sessionPropertyKey(cleanToken));
}

function sessionPropertyKey(token) {
  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(token || ""),
    Utilities.Charset.UTF_8
  );
  const hash = digest.map((value) => (`0${((value + 256) % 256).toString(16)}`).slice(-2)).join("");
  return `${SESSION_PROPERTY_PREFIX}${hash}`;
}

function cleanupExpiredSessions() {
  const properties = PropertiesService.getScriptProperties();
  const allProperties = properties.getProperties();
  const now = Date.now();

  Object.keys(allProperties).forEach((key) => {
    if (!key.startsWith(SESSION_PROPERTY_PREFIX)) return;
    try {
      const session = JSON.parse(allProperties[key]);
      if (!session.expiresAt || now >= Number(session.expiresAt)) properties.deleteProperty(key);
    } catch (error) {
      properties.deleteProperty(key);
    }
  });
}

function findUserByEmail(email) {
  const sheet = getUsersSheet();
  const values = sheet.getDataRange().getValues();

  for (let index = 1; index < values.length; index += 1) {
    const row = userValuesToRow(values[index]);
    if (row.email === email) return { user: row, rowNumber: index + 1 };
  }

  return null;
}

function userValuesToRow(values) {
  const role = String(values[2] || "Viewer").trim();

  return {
    email: String(values[0] || "").trim().toLowerCase(),
    name: String(values[1] || "").trim(),
    role: ["Admin", "Editor", "Viewer"].includes(role) ? role : "Viewer",
    active: values[3] === true || String(values[3]).toUpperCase() === "TRUE",
    password: String(values[4] || "")
  };
}

function publicUser(user) {
  return {
    email: user.email,
    name: user.name,
    role: user.role,
    active: user.active
  };
}

function updateLastLogin(rowNumber) {
  getUsersSheet().getRange(rowNumber, 6).setValue(new Date());
}

function appendAudit(entry) {
  const email = String(entry && entry.email ? entry.email : "").trim().toLowerCase();
  const action = String(entry && entry.action ? entry.action : "").trim();
  const module = String(entry && entry.module ? entry.module : "").trim();
  if (!email || !action || !module) return;

  getAuditSheet().appendRow([
    new Date(),
    safeAuditValue(email),
    safeAuditValue(entry.role),
    safeAuditValue(action),
    safeAuditValue(module),
    safeAuditValue(entry.detail)
  ]);
}

function safeAuditValue(value) {
  const text = String(value || "").slice(0, 1000);
  return /^[=+\-@]/.test(text) ? `'${text}` : text;
}

function callbackResponse(callback, payload) {
  const safeCallback = String(callback || "");

  if (safeCallback) {
    if (!/^[A-Za-z_$][\w$]*(\.[A-Za-z_$][\w$]*)*$/.test(safeCallback)) {
      return ContentService
        .createTextOutput(JSON.stringify({ ok: false, error: "Invalid callback" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService
      .createTextOutput(`${safeCallback}(${JSON.stringify(payload)});`)
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function getUsersSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(USERS_SHEET_NAME) || spreadsheet.insertSheet(USERS_SHEET_NAME);

  if (sheet.getLastRow() === 0 || sheet.getRange(1, 1).getValue() === "") {
    sheet.getRange(1, 1, 1, USERS_HEADERS.length).setValues([USERS_HEADERS]);
  }

  return sheet;
}

function getAuditSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(AUDIT_SHEET_NAME) || spreadsheet.insertSheet(AUDIT_SHEET_NAME);

  if (sheet.getLastRow() === 0 || sheet.getRange(1, 1).getValue() === "") {
    sheet.getRange(1, 1, 1, AUDIT_HEADERS.length).setValues([AUDIT_HEADERS]);
  }

  return sheet;
}
