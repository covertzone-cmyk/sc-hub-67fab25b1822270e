/**
 * Command Center board backend. Source for the Apps Script web app.
 *
 * Deploy: Execute as Me. Who has access: Anyone.
 * Set Script property SCRIPT_SECRET in Project Settings. Never put the
 * real value in this file or in a comment.
 *
 * The live board (index.html) talks to the /exec URL with:
 *   Content-Type: text/plain
 *   Auth: shared secret in the JSON body (POST) or ?secret= (GET)
 *
 * Contract
 *   GET  ?secret=                         -> { ok, state }
 *   POST { secret, op: "saveState", state } -> { ok }
 *   POST { secret, op: "action", kind, itemId, label, to, subject, text, title, notes }
 *        kind: draft | task | note | done | snooze | drop
 *
 * Drafts use GmailApp.createDraft only. This script never sends email.
 *
 * Paste the published /exec URL and the same secret into PIPE_URL and
 * PIPE_SECRET at the top of CFG in index.html. Leave those constants
 * empty until you have a deployed web app.
 */

var STATE_FILE_PROP = "STATE_FILE_ID";
var NOTES_KEY = "CC_NOTES";
var PROP_SECRET = "SCRIPT_SECRET";

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function expectedSecret() {
  return PropertiesService.getScriptProperties().getProperty(PROP_SECRET) || "";
}

function checkSecret(got) {
  var want = expectedSecret();
  return !!(want && got && got === want);
}

function stateFile() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty(STATE_FILE_PROP);
  if (id) {
    try { return DriveApp.getFileById(id); } catch (e) {}
  }
  var f = DriveApp.createFile("cc-board-state.json", "{}", MimeType.PLAIN_TEXT);
  props.setProperty(STATE_FILE_PROP, f.getId());
  return f;
}

function loadState() {
  try {
    return JSON.parse(stateFile().getBlob().getDataAsString() || "{}") || {};
  } catch (e) {
    return {};
  }
}

function saveState(state) {
  stateFile().setContent(JSON.stringify(state || {}));
}

function loadNotes() {
  var raw = PropertiesService.getScriptProperties().getProperty(NOTES_KEY);
  if (!raw) return [];
  try { return JSON.parse(raw); } catch (e) { return []; }
}

function saveNotes(notes) {
  PropertiesService.getScriptProperties().setProperty(NOTES_KEY, JSON.stringify(notes || []));
}

function doGet(e) {
  var secret = (e && e.parameter && e.parameter.secret) || "";
  if (!checkSecret(secret)) return jsonOut({ ok: false, error: "unauthorized" });
  return jsonOut({ ok: true, state: loadState() });
}

function doPost(e) {
  var body = {};
  try {
    body = JSON.parse((e && e.postData && e.postData.contents) || "{}");
  } catch (err) {
    return jsonOut({ ok: false, error: "bad json" });
  }
  if (!checkSecret(body.secret)) return jsonOut({ ok: false, error: "unauthorized" });

  if (body.op === "saveState") {
    saveState(body.state || {});
    return jsonOut({ ok: true });
  }
  if (body.op === "action") return handleAction(body);
  return jsonOut({ ok: false, error: "unknown op" });
}

function handleAction(body) {
  var kind = body.kind || "";

  if (kind === "draft") {
    GmailApp.createDraft(body.to || "", body.subject || "", body.text || "");
    return jsonOut({ ok: true, message: "Draft created in SC@. Nothing was sent." });
  }

  if (kind === "task") {
    var title = body.title || body.label || "Task from the board";
    var notes = body.notes || body.text || "";
    var tasks = loadNotes();
    tasks.unshift({ kind: "task", title: title, notes: notes, itemId: body.itemId || "", at: new Date().toISOString() });
    saveNotes(tasks.slice(0, 80));
    try {
      Tasks.Tasks.insert({ title: title, notes: notes }, "@default");
    } catch (err) {
      /* Tasks API is optional. The note is stored either way. */
    }
    return jsonOut({ ok: true, message: "Task added." });
  }

  if (kind === "note") {
    var notes = loadNotes();
    notes.unshift({
      kind: "note",
      itemId: body.itemId || "",
      label: body.label || "",
      text: body.text || "",
      title: body.title || "",
      notes: body.notes || "",
      at: new Date().toISOString()
    });
    saveNotes(notes.slice(0, 80));
    return jsonOut({ ok: true, message: "Noted." });
  }

  if (kind === "done" || kind === "drop" || kind === "snooze") {
    return jsonOut({ ok: true });
  }

  return jsonOut({ ok: false, error: "unknown kind" });
}
