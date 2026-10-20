// Rob's Class App: Google Sheets backend (version 3)
// Paste this whole script into Extensions > Apps Script (from inside the Google Sheet you want to use).

// Same password you use for the Teacher button. The app's "Copy" button fills this in for you.
var TEACHER_PASSWORD = 'YOUR_PASSWORD_HERE';
var VERSION = 3;

var H_RESULTS = ['Timestamp', 'Name', 'Class', 'Homework', 'Score', 'Total', 'Percent', 'Late', 'Flags', 'Seconds', 'HW ID', 'ID'];
var H_ANSWERS = ['ID', 'Name', 'Quiz', '#', 'Question', 'Student answer', 'Correct answer', 'Correct?', 'Points'];
var H_LOG = ['ID', 'Name', 'Quiz', 'Seconds into quiz', 'Event', 'Detail'];
var H_CLASSES = ['Class', 'Created'];
var H_HW = ['ID', 'Title', 'Classes', 'Deadline', 'Allow late', 'Minutes', 'Shuffle questions', 'Shuffle options', 'Show answers', 'Block copy', 'Status', 'Issued', 'Source', 'Questions'];
var H_Q = ['HW ID', '#', 'Question', 'Options', 'Answer', 'Points'];

// Actions that change the Sheet. They run one at a time so two students submitting together can't collide.
var WRITE_ACTIONS = { submit: 1, addClass: 1, deleteClass: 1, issue: 1, updateHw: 1, deleteHw: 1 };

// ---------- entry points ----------
function doGet(e) {
  var p = (e && e.parameter) || {};
  if (!p.action) return ContentService.createTextOutput("Rob's Class App webhook is live (v" + VERSION + ")");
  try {
    if (p.action === 'ping') return json_({ ok: true, version: VERSION });
    if (p.action === 'classes') return json_({ ok: true, classes: publicClasses_() });
    if (p.action === 'quiz') return json_(getQuiz_(p.id));
    return json_({ ok: false, error: 'Unknown action' });
  } catch (err) {
    return json_({ ok: false, error: msg_(err) });
  }
}

function doPost(e) {
  var d;
  try {
    d = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: 'Bad request' });
  }
  var a = (d && d.action) || 'submit';
  try {
    // Students need no password. Everything else does. Checked before taking the lock, so wrong guesses can't hold up students.
    if (a !== 'submit') checkPassword_(d.pw);

    if (!WRITE_ACTIONS[a]) {
      switch (a) {
        case 'adminPing': return json_({ ok: true, version: VERSION });
        case 'adminList': return json_(adminList_());
        case 'readSheet': return json_(readSheet_(d.url));
        case 'classResults': return json_(classResults_(d.name));
      }
      return json_({ ok: false, error: 'Unknown action' });
    }

    var lock = LockService.getScriptLock();
    try {
      lock.waitLock(30000);
    } catch (err) {
      return json_({ ok: false, error: 'The Google script is busy. Try again in a moment.' });
    }
    try {
      switch (a) {
        case 'submit': return json_(saveResult_(d));
        case 'addClass': return json_(addClass_(d.name));
        case 'deleteClass': return json_(deleteClass_(d.name));
        case 'issue': return json_(issue_(d));
        case 'updateHw': return json_(updateHw_(d));
        case 'deleteHw': return json_(deleteHw_(d.id));
      }
    } finally {
      lock.releaseLock();
    }
    return json_({ ok: false, error: 'Unknown action' });
  } catch (err) {
    return json_({ ok: false, error: msg_(err) });
  }
}

// Run this once from the editor (choose "setup" and press Run) to grant permission and create the tabs.
function setup() {
  sheet_('Results', H_RESULTS);
  sheet_('Answers', H_ANSWERS);
  sheet_('Activity log', H_LOG);
  sheet_('Classes', H_CLASSES, true);
  sheet_('Homework', H_HW, true);
  sheet_('Questions', H_Q, true);
  Logger.log('Ready. Your tabs are in: ' + ss_().getUrl());
}

// ---------- helpers ----------
function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function msg_(err) {
  return String((err && err.message) || err);
}

function ss_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

// Finds a tab (creating it with a header row if needed). If an older tab with the same name has
// different columns, it is renamed "(old ...)" and a fresh one is made, so nothing is lost.
function sheet_(name, headers, asText) {
  var ss = ss_();
  var sh = ss.getSheetByName(name);
  if (sh && sh.getLastRow() > 0) {
    var cur = sh.getRange(1, 1, 1, headers.length).getValues()[0].join('|');
    if (cur !== headers.join('|')) {
      sh.setName(name + ' (old ' + Utilities.formatDate(new Date(), 'UTC', 'yyyy-MM-dd HHmmss') + ')');
      sh = null;
    }
  }
  if (!sh) {
    sh = ss.insertSheet(name);
    if (asText) sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).setNumberFormat('@');
  }
  if (sh.getLastRow() === 0) {
    sh.appendRow(headers);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
  return sh;
}

function append_(sh, rows) {
  if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
}

// Stops names/answers that start with = + - @ from running as spreadsheet formulas
function safe_(v) {
  v = (v === null || v === undefined) ? '' : String(v);
  return ['=', '+', '-', '@'].indexOf(v.charAt(0)) !== -1 ? "'" + v : v;
}

// For teacher-typed names and titles: tidy spaces and drop leading = + - @
function clean_(s) {
  return String(s === null || s === undefined ? '' : s).replace(/\s+/g, ' ').replace(/^[=+\-@\s]+/, '').trim();
}

function cap_(v, n) {
  return String(v === null || v === undefined ? '' : v).slice(0, n);
}

function num_(v) {
  v = Number(v);
  return isFinite(v) ? v : 0;
}

// Wrong guesses are slowed down (not locked out), so a student can't lock the teacher out by guessing.
function checkPassword_(pw) {
  if (!TEACHER_PASSWORD || TEACHER_PASSWORD.indexOf('YOUR_PASSWORD') === 0) {
    throw new Error('Set TEACHER_PASSWORD at the top of the Google script (use the Copy button in the app, which fills it in).');
  }
  if (typeof pw !== 'string' || pw !== TEACHER_PASSWORD) {
    Utilities.sleep(2000);
    throw new Error('The password in the Google script does not match the one you typed. Copy the script from the app again and redeploy.');
  }
}

// ---------- results from students ----------
function saveResult_(d) {
  var name = safe_(cap_(d.name, 60).replace(/\s+/g, ' ').trim());
  var id = cap_(d.id, 40);
  if (!name || !id) throw new Error('Missing name or id.');

  var res = sheet_('Results', H_RESULTS);
  var n = res.getLastRow();
  var ids = n > 1 ? res.getRange(2, 12, n - 1, 1).getValues().map(function (r) { return r[0]; }) : [];
  if (ids.indexOf(id) !== -1) return { ok: true, duplicate: true };

  var ts = new Date(d.timestamp);
  if (isNaN(ts.getTime())) ts = new Date();
  var integ = d.integrity || {};
  var flags = integ.flagged ? cap_(integ.summary, 500) : '';
  var quiz = safe_(cap_(d.quiz, 100));
  res.appendRow([ts, name, safe_(cap_(d['class'], 60)), quiz, num_(d.score), num_(d.total), num_(d.percent),
    d.late ? 'LATE' : '', safe_(flags), num_(d.seconds), cap_(d.hwId, 40), id]);
  var r = res.getLastRow();
  if (flags) res.getRange(r, 1, 1, H_RESULTS.length).setBackground('#fde2e2');
  else if (d.late) res.getRange(r, 1, 1, H_RESULTS.length).setBackground('#fff4cc');

  var rows = (d.answers || []).slice(0, 300).map(function (a) {
    return [id, name, quiz, num_(a.n), safe_(cap_(a.q, 500)), safe_(cap_(a.chosen, 300)), safe_(cap_(a.correctAnswer, 300)), a.ok ? 'YES' : 'NO', num_(a.earned)];
  });
  append_(sheet_('Answers', H_ANSWERS), rows);

  var log = (d.log || []).slice(0, 300).map(function (l) {
    return [id, name, quiz, num_(l.t), safe_(cap_(l.type, 60)), safe_(cap_(l.detail, 200))];
  });
  append_(sheet_('Activity log', H_LOG), log);
  return { ok: true };
}

// ---------- classes ----------
function classNames_() {
  var sh = sheet_('Classes', H_CLASSES, true);
  var n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, 1).getValues().map(function (r) { return String(r[0]); }).filter(function (s) { return s; });
}

function addClass_(name) {
  name = clean_(name);
  if (!name) throw new Error('Type a class name.');
  if (name.length > 40) throw new Error('Class name is too long (40 characters max).');
  var list = classNames_();
  for (var i = 0; i < list.length; i++) {
    if (list[i].toLowerCase() === name.toLowerCase()) throw new Error('That class already exists.');
  }
  sheet_('Classes', H_CLASSES, true).appendRow([name, new Date().toISOString()]);
  return { ok: true, classes: classNames_() };
}

function deleteClass_(name) {
  var sh = sheet_('Classes', H_CLASSES, true);
  var n = sh.getLastRow();
  if (n > 1) {
    var vals = sh.getRange(2, 1, n - 1, 1).getValues();
    for (var i = vals.length - 1; i >= 0; i--) {
      if (String(vals[i][0]) === String(name)) sh.deleteRow(i + 2);
    }
  }
  return { ok: true, classes: classNames_() };
}

// ---------- homework ----------
function rowToHw_(r) {
  var classes = [];
  try { classes = JSON.parse(r[2]) || []; } catch (e) { classes = []; }
  return {
    id: String(r[0]), title: String(r[1]), classes: classes, deadline: String(r[3] || ''),
    allowLate: String(r[4]) !== 'no', minutes: Number(r[5]) || 0,
    shuffleQ: String(r[6]) === 'yes', shuffleO: String(r[7]) === 'yes',
    showReview: String(r[8]) === 'yes', blockCopy: String(r[9]) !== 'no',
    status: String(r[10]) || 'open', issued: String(r[11] || ''), source: String(r[12] || ''), count: Number(r[13]) || 0
  };
}

function readHw_() {
  var sh = sheet_('Homework', H_HW, true);
  var n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, H_HW.length).getValues().map(rowToHw_).filter(function (h) { return h.id; });
}

function publicHw_(h) {
  return { id: h.id, title: h.title, classes: h.classes, deadline: h.deadline, allowLate: h.allowLate, minutes: h.minutes,
    shuffleQ: h.shuffleQ, shuffleO: h.shuffleO, showReview: h.showReview, blockCopy: h.blockCopy, count: h.count };
}

function findRow_(sh, id) {
  var n = sh.getLastRow();
  if (n < 2) return 0;
  var ids = sh.getRange(2, 1, n - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2;
  }
  return 0;
}

function publicClasses_() {
  var open = readHw_().filter(function (h) { return h.status === 'open'; });
  return classNames_().map(function (name) {
    return {
      name: name,
      homework: open.filter(function (h) { return h.classes.indexOf(name) !== -1; }).map(publicHw_)
    };
  });
}

function getQuiz_(id) {
  var hw = null;
  readHw_().forEach(function (h) { if (h.id === String(id)) hw = h; });
  if (!hw) throw new Error('That homework was not found.');
  if (hw.status !== 'open') throw new Error('This homework has been closed by your teacher.');
  var sh = sheet_('Questions', H_Q, true);
  var n = sh.getLastRow();
  var qs = [];
  if (n > 1) {
    sh.getRange(2, 1, n - 1, H_Q.length).getValues().forEach(function (r) {
      if (String(r[0]) === hw.id) {
        var raw = String(r[4]);
        var item = { n: Number(r[1]), q: String(r[2]), options: JSON.parse(r[3]), points: Number(r[5]) || 1 };
        if (raw.charAt(0) === '[') { item.type = 'text'; item.answers = JSON.parse(raw); }
        else { item.type = 'mc'; item.answer = Number(raw); }
        qs.push(item);
      }
    });
  }
  qs.sort(function (a, b) { return a.n - b.n; });
  return { ok: true, hw: publicHw_(hw), questions: qs };
}

function adminList_() {
  var res = sheet_('Results', H_RESULTS);
  var n = res.getLastRow();
  var counts = {};
  if (n > 1) {
    res.getRange(2, 11, n - 1, 1).getValues().forEach(function (r) {
      var k = String(r[0]);
      if (k) counts[k] = (counts[k] || 0) + 1;
    });
  }
  var hws = readHw_().reverse().map(function (h) { h.submissions = counts[h.id] || 0; return h; });
  return { ok: true, sheetUrl: ss_().getUrl(), classes: classNames_(), homework: hws };
}

// Latest submissions for one class (newest first), for the class page in the app.
function classResults_(name) {
  var res = sheet_('Results', H_RESULTS);
  var n = res.getLastRow();
  var out = [];
  if (n > 1) {
    var vals = res.getRange(2, 1, n - 1, H_RESULTS.length).getValues();
    for (var i = vals.length - 1; i >= 0 && out.length < 300; i--) {
      var r = vals[i];
      if (String(r[2]) !== String(name)) continue;
      out.push({
        ts: r[0] instanceof Date ? r[0].toISOString() : String(r[0]), name: String(r[1]), hw: String(r[3]),
        score: Number(r[4]), total: Number(r[5]), percent: Number(r[6]), late: String(r[7]) === 'LATE',
        flags: String(r[8]), seconds: Number(r[9]), hwId: String(r[10])
      });
    }
  }
  return { ok: true, results: out };
}

// Reads the question rows out of a Google Sheet in your Drive. The app turns them into questions.
function readSheet_(url) {
  url = String(url || '').trim();
  if (!url) throw new Error('Paste the link to your Google Sheet.');
  var src;
  try {
    src = /^https?:/i.test(url) ? SpreadsheetApp.openByUrl(url) : SpreadsheetApp.openById(url);
  } catch (err) {
    throw new Error("Couldn't open that Google Sheet. Check the link, and that it is a real Google Sheet (an Excel file in Drive must first be opened and saved as a Google Sheet) owned by the same Google account as this script.");
  }
  var sheets = src.getSheets();
  var sh = sheets[0];
  var m = url.match(/[#&?]gid=(\d+)/);
  if (m) {
    sheets.forEach(function (s) { if (String(s.getSheetId()) === m[1]) sh = s; });
  }
  var rows = sh.getDataRange().getDisplayValues();
  if (rows.length > 1001) rows = rows.slice(0, 1001);
  return { ok: true, title: src.getName(), tab: sh.getName(), rows: rows };
}

function issue_(d) {
  var title = clean_(d.title);
  if (!title) throw new Error('Give the homework a title.');
  var known = classNames_();
  var classes = (d.classes || []).map(String).filter(function (c) { return known.indexOf(c) !== -1; });
  if (!classes.length) throw new Error('Pick at least one class.');
  var qs = d.questions || [];
  if (!qs.length) throw new Error('There are no questions to issue.');
  if (qs.length > 300) throw new Error('Too many questions (300 max).');
  qs.forEach(function (q, i) {
    var bad;
    if (q && q.type === 'text') {
      // typed answer: no options, one or more accepted answers
      bad = typeof q.q !== 'string' || !q.q.trim() || q.q.length > 2000 || !(Number(q.points) > 0) ||
        !q.answers || !q.answers.length || q.answers.length > 10 ||
        q.answers.some(function (a) { return typeof a !== 'string' || !a.trim() || a.length > 200; });
    } else {
      bad = !q || typeof q.q !== 'string' || !q.q.trim() || q.q.length > 2000 ||
        !q.options || q.options.length < 2 || q.options.length > 8 ||
        q.options.some(function (o) { return String(o).length > 500; }) ||
        !(q.answer >= 0 && q.answer < q.options.length) || !(Number(q.points) > 0);
    }
    if (bad) throw new Error('Question ' + (i + 1) + ' is not valid.');
  });
  var id = 'hw-' + Utilities.getUuid().slice(0, 8);
  sheet_('Homework', H_HW, true).appendRow([
    id, title, JSON.stringify(classes), String(d.deadline || ''), d.allowLate === false ? 'no' : 'yes', Number(d.minutes) || 0,
    d.shuffleQ ? 'yes' : 'no', d.shuffleO ? 'yes' : 'no', d.showReview ? 'yes' : 'no', d.blockCopy === false ? 'no' : 'yes',
    'open', new Date().toISOString(), clean_(d.source || ''), qs.length
  ]);
  var rows = qs.map(function (q, i) {
    // A typed-answer question has no options, and its Answer cell holds the list of accepted answers as JSON.
    if (q.type === 'text') return [id, i + 1, String(q.q), '[]', JSON.stringify(q.answers.map(String)), Number(q.points)];
    return [id, i + 1, String(q.q), JSON.stringify(q.options.map(String)), q.answer, Number(q.points)];
  });
  append_(sheet_('Questions', H_Q, true), rows);
  return { ok: true, id: id };
}

function updateHw_(d) {
  var sh = sheet_('Homework', H_HW, true);
  var row = findRow_(sh, d.id);
  if (!row) throw new Error('That homework was not found.');
  var set = function (col, v) { sh.getRange(row, col).setValue(v); };
  if (d.title !== undefined) {
    var t = clean_(d.title);
    if (!t) throw new Error('Title cannot be empty.');
    set(2, t);
  }
  if (d.deadline !== undefined) set(4, String(d.deadline || ''));
  if (d.allowLate !== undefined) set(5, d.allowLate ? 'yes' : 'no');
  if (d.minutes !== undefined) set(6, Number(d.minutes) || 0);
  if (d.status !== undefined) set(11, d.status === 'closed' ? 'closed' : 'open');
  return { ok: true };
}

function deleteHw_(id) {
  var sh = sheet_('Homework', H_HW, true);
  var row = findRow_(sh, id);
  if (row) sh.deleteRow(row);
  var q = sheet_('Questions', H_Q, true);
  var n = q.getLastRow();
  if (n > 1) {
    var ids = q.getRange(2, 1, n - 1, 1).getValues();
    for (var i = ids.length - 1; i >= 0; i--) {
      if (String(ids[i][0]) === String(id)) q.deleteRow(i + 2);
    }
  }
  return { ok: true };
}
