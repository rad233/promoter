function selAns(tn, q, letter, el) {
  var key = "t" + tn;
  if (!S.answers[key]) S.answers[key] = {};
  S.answers[key][q] = letter;
  // Update visual selection
  var row = el.closest(".q-block");
  if (row) {
    var opts = row.querySelectorAll(".opt");
    for (var i = 0; i < opts.length; i++) opts[i].className = "opt";
  }
  el.className = "opt sel";
  // Sync sidebar
  syncSidebar(tn, q, letter);
}

function safeCopy(text, successMsg) {
  try {
    var ta = document.createElement("textarea");
    ta.value = text;
    ta.style.cssText =
      "position:fixed;top:0;left:0;width:2em;height:2em;opacity:0;font-size:16px;border:none;outline:none;";
    ta.setAttribute("readonly", "");
    document.body.appendChild(ta);
    ta.focus();
    ta.setSelectionRange(0, 99999);
    try {
      document.execCommand("copy");
    } catch (e2) {}
    document.body.removeChild(ta);
    showToast(successMsg || "Copied!", "success");
  } catch (e) {
    showToast("Copy failed", "");
  }
}

async function postAnnouncement(pin) {
  var t = document.getElementById("ann-t");
  var b = document.getElementById("ann-b");
  if (!t || !b || !t.value.trim() || !b.value.trim()) {
    showToast("Please fill in both title and message", "");
    return;
  }
  var title = t.value.trim(),
    body = b.value.trim();
  // Update student-facing announcement
  var annTitle = document.getElementById("ann-title"),
    annBody = document.getElementById("ann-body");
  if (annTitle) annTitle.textContent = title;
  if (annBody) annBody.textContent = body;
  // Pin on student dashboard if requested
  if (pin) {
    var pt = document.getElementById("pin-title"),
      pb = document.getElementById("pin-body");
    if (pt) pt.textContent = title;
    if (pb) pb.textContent = body;
    var pa = document.getElementById("pinned-ann");
    if (pa) pa.classList.add("show");
  }
  // Add to history
  var h = document.getElementById("ann-history");
  if (h) {
    var d = document.createElement("div");
    d.style.cssText =
      "padding:12px;background:var(--gray-50);border-radius:8px;margin-bottom:8px;";
    d.innerHTML =
      '<div style="font-weight:600;font-size:13px;">' +
      title +
      "</div>" +
      '<div style="font-size:12px;color:var(--gray-600);margin-top:4px;">' +
      body +
      "</div>" +
      (pin
        ? '<div style="margin-top:6px;"><span style="font-size:11px;color:var(--pink);font-weight:600;">PINNED</span></div>'
        : "");
    h.insertBefore(d, h.firstChild);
  }
  // Save to Supabase
  await dbInsert("announcements", {
    title: title,
    body: body,
    pinned: pin ? true : false,
  });
  t.value = "";
  b.value = "";
  showToast(pin ? "Posted and pinned!" : "Posted!", "success");
}
// ── Supabase ─────────────────────────────────────────────────────────────────
const SUPA_URL = "https://ntqvzwztgcorgssxfvbf.supabase.co";
const SUPA_KEY =
  "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Im50cXZ6d3p0Z2Nvcmdzc3hmdmJmIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzM2MDI0MDQsImV4cCI6MjA4OTE3ODQwNH0.5dd3ObqsBf7h49IWWaNuzKcFeOiKUtTfK0iH-TmuA1o";
const sb = supabase.createClient(SUPA_URL, SUPA_KEY);

async function dbGet(table, filters) {
  try {
    var q = sb.from(table).select("*");
    if (filters) {
      for (var k in filters) q = q.eq(k, filters[k]);
    }
    var { data, error } = await q;
    if (error) {
      console.error("dbGet", table, error);
      return [];
    }
    return data || [];
  } catch (e) {
    console.error("dbGet err", e);
    return [];
  }
}
async function dbUpsert(table, row, conflict) {
  try {
    var q = sb
      .from(table)
      .upsert(row, conflict ? { onConflict: conflict } : {});
    var { data, error } = await q;
    if (error) {
      console.error(
        "dbUpsert ERROR",
        table,
        error.message,
        error.details,
        error.hint,
      );
      return false;
    }
    return true;
  } catch (e) {
    console.error("dbUpsert EXCEPTION", table, e);
    return false;
  }
}
async function dbInsert(table, row) {
  try {
    var { error } = await sb.from(table).insert(row);
    if (error) {
      console.error("dbInsert", table, error);
      return false;
    }
    return true;
  } catch (e) {
    console.error("dbInsert err", e);
    return false;
  }
}
async function dbUpdate(table, row, filters) {
  try {
    var q = sb.from(table).update(row);
    for (var k in filters) q = q.eq(k, filters[k]);
    var { error } = await q;
    if (error) {
      console.error("dbUpdate", table, error);
      return false;
    }
    return true;
  } catch (e) {
    console.error("dbUpdate err", e);
    return false;
  }
}
async function dbDelete(table, filters) {
  try {
    var q = sb.from(table).delete();
    for (var k in filters) q = q.eq(k, filters[k]);
    var { error } = await q;
    if (error) {
      console.error("dbDelete", table, error);
      return false;
    }
    return true;
  } catch (e) {
    console.error("dbDelete err", e);
    return false;
  }
}

var S = {
  regCodes: {},
  examCodes: {},
  examIDs: {},
  usedIDs: {},
  adminAuth: null,
  student: null,
  answers: { t1: {}, t2: {}, t3: {}, t4: {}, t5: {}, t6: {}, t7: "" },
  timeLeft: 150 * 60,
  timerInt: null,
  timerPaused: false,
  liveInt: null,
  curTask: 1,
  exams: {},
  schedule: {},
  parsedExam: null,
  parsedTasks: null,
  grades: {},
  completed: {},
  activeExamCode: null,
  objectHasValue: function (obj, val) {
    var ks = Object.keys(obj);
    for (var i = 0; i < ks.length; i++) if (obj[ks[i]] === val) return true;
    return false;
  },
};

function showPage(id) {
  var pages = document.querySelectorAll(".page");
  for (var i = 0; i < pages.length; i++) {
    pages[i].classList.remove("active");
    pages[i].style.display = "none";
  }
  var el = document.getElementById(id);
  if (el) {
    el.classList.add("active");
    el.style.display = "flex";
    if (id === "page-admin" || id === "page-student") {
      el.style.background = "#f9f9f9";
    }
  }
  var ao = document.getElementById("ann-overlay");
  if (ao) {
    ao.classList.remove("open");
  }
  var overlays = document.querySelectorAll(".overlay");
  for (var j = 0; j < overlays.length; j++) {
    overlays[j].classList.remove("open");
    overlays[j].style.display = "";
  }
  // Enable exam protection when entering exam page, disable otherwise
  if (id === "page-exam") {
    document.body.classList.add("exam-active");
    enableExamProtection();
  } else {
    document.body.classList.remove("exam-active");
    disableExamProtection();
  }
}

function enableExamProtection() {
  document.addEventListener("copy", _blockExamEvent, true);
  document.addEventListener("cut", _blockExamEvent, true);
  document.addEventListener("contextmenu", _blockExamEvent, true);
  document.addEventListener("keydown", _blockExamKeys, true);
  // Block paste everywhere except the essay textarea
  document.addEventListener("paste", _blockPasteExceptEssay, true);
}

function disableExamProtection() {
  document.removeEventListener("copy", _blockExamEvent, true);
  document.removeEventListener("cut", _blockExamEvent, true);
  document.removeEventListener("contextmenu", _blockExamEvent, true);
  document.removeEventListener("keydown", _blockExamKeys, true);
  document.removeEventListener("paste", _blockPasteExceptEssay, true);
}

function _blockExamEvent(e) {
  e.preventDefault();
  e.stopPropagation();
  return false;
}

function _blockPasteExceptEssay(e) {
  // Allow paste only inside the essay textarea
  var target = e.target;
  if (
    target &&
    (target.id === "essay-ta" || target.classList.contains("essay-ta"))
  )
    return;
  e.preventDefault();
  e.stopPropagation();
  return false;
}

function _blockExamKeys(e) {
  // Block Ctrl/Cmd + C, X, A, V (copy, cut, select all, paste)
  if (e.ctrlKey || e.metaKey) {
    var key = e.key ? e.key.toLowerCase() : "";
    if (key === "c" || key === "x" || key === "a") {
      // Allow Ctrl+A inside essay textarea
      var target = e.target;
      if (
        key === "a" &&
        target &&
        (target.id === "essay-ta" || target.classList.contains("essay-ta"))
      )
        return;
      e.preventDefault();
      e.stopPropagation();
      return false;
    }
    if (key === "v") {
      // Allow Ctrl+V only inside essay textarea
      var target2 = e.target;
      if (
        target2 &&
        (target2.id === "essay-ta" || target2.classList.contains("essay-ta"))
      )
        return;
      e.preventDefault();
      e.stopPropagation();
      return false;
    }
  }
}
function showToast(msg, type) {
  var t = document.getElementById("toast");
  t.textContent = msg;
  t.className = "toast show" + (type ? " " + type : "");
  setTimeout(function () {
    t.classList.remove("show");
  }, 3000);
}
function closeOverlay(id) {
  document.getElementById(id).classList.remove("open");
}
function pad(n) {
  return n < 10 ? "0" + n : "" + n;
}

function togglePw(inputId, btn) {
  var inp = document.getElementById(inputId);
  inp.type = inp.type === "password" ? "text" : "password";
  var svg = btn.querySelector("svg");
  if (svg) svg.style.opacity = inp.type === "text" ? "0.4" : "1";
}

// AUTH
function switchTab(tab) {
  document.getElementById("tab-login").className =
    "auth-tab" + (tab === "login" ? " active" : "");
  document.getElementById("tab-signup").className =
    "auth-tab" + (tab === "signup" ? " active" : "");
  document.getElementById("form-login").style.display =
    tab === "login" ? "block" : "none";
  document.getElementById("form-signup").style.display =
    tab === "signup" ? "block" : "none";
}

async function doLogin() {
  var email = document.getElementById("li-code").value.trim().toLowerCase();
  var pw = document.getElementById("li-pw").value;
  var err = document.getElementById("login-err");
  err.style.display = "none";
  if (!email) {
    err.textContent = "Please enter your email address.";
    err.style.display = "block";
    return;
  }
  var rows = await dbGet("students", { email: email });
  if (!rows.length || !rows[0].activated) {
    err.textContent = "No account found, or account not activated.";
    err.style.display = "block";
    return;
  }
  if (rows[0].password !== pw) {
    err.textContent = "Incorrect password.";
    err.style.display = "block";
    return;
  }
  S.student = {
    name: rows[0].name,
    email: rows[0].email,
    phone: rows[0].phone || "",
    pw: rows[0].password,
    activated: true,
    id: rows[0].student_id || "",
  };
  // Reload fresh data before showing dashboard so grades are current
  await loadAllData();
  enterDash();
}

async function doSignup() {
  var code = document.getElementById("su-code").value.trim();
  var name = document.getElementById("su-name").value.trim();
  var email = document.getElementById("su-email").value.trim();
  var phone = document.getElementById("su-phone")
    ? document.getElementById("su-phone").value.trim()
    : "";
  var pw = document.getElementById("su-pw").value;
  var err = document.getElementById("signup-err");
  err.style.display = "none";
  if (!name || !email || !pw || !code) {
    err.textContent = "Please fill in all required fields.";
    err.style.display = "block";
    return;
  }
  var rows = await dbGet("students", { reg_code: code });
  if (!rows.length) {
    err.textContent = "Invalid activation code.";
    err.style.display = "block";
    return;
  }
  if (rows[0].activated) {
    err.textContent = "This code has already been used.";
    err.style.display = "block";
    return;
  }
  var ok = await dbUpdate(
    "students",
    { name: name, email: email, phone: phone, password: pw, activated: true },
    { reg_code: code },
  );
  if (!ok) {
    err.textContent = "Registration failed. Please try again.";
    err.style.display = "block";
    return;
  }
  S.student = {
    name: name,
    email: email,
    phone: phone,
    pw: pw,
    activated: true,
    id: rows[0].student_id || "",
  };
  enterDash();
}

async function doAdminLogin() {
  var u = document.getElementById("al-user").value.trim();
  var p = document.getElementById("al-pw").value;
  var err = document.getElementById("al-err");
  if (err) err.style.display = "none";

  const { data, error } = await sb.auth.signInWithPassword({
    email: u,
    password: p,
  });

  if (data.user && !error) {
    showPage("page-admin");
    showSec("dashboard");
    buildEssays();
    startLiveCountdown();
    refreshStudentsTable();
    refreshCodesTable();
    renderExamList();
    renderSchedTable();
    updateAdminDashboard();
    clearInterval(window._dashInt);
    window._dashInt = setInterval(updateAdminDashboard, 5000);
  } else {
    if (err) {
      err.textContent = error ? error.message : "Incorrect credentials.";
      err.style.display = "block";
    }
  }
}

async function doLogout() {
  await sb.auth.signOut();
  S.student = null;
  S.answers = { t1: {}, t2: {}, t3: {}, t4: {}, t5: {}, t6: {}, t7: "" };
  var _activeEx = S.activeExamCode && S.exams[S.activeExamCode];
  var _examDur = _activeEx ? _activeEx.duration || 150 : 150;
  S.timeLeft = _examDur * 60;
  S.examStartTime = Date.now();
  clearInterval(S.timerInt);
  clearInterval(S.liveInt);
  showPage("page-auth");
  var liCode = document.getElementById("li-code");
  if (liCode) liCode.value = "";
  var liPw = document.getElementById("li-pw");
  if (liPw) liPw.value = "";
}

function enterDash() {
  var n = S.student.name;
  document.getElementById("s-name").textContent = n;
  document.getElementById("s-avatar").textContent = n.charAt(0).toUpperCase();
  document.getElementById("s-welcome").textContent =
    "Welcome, " + n.split(" ")[0] + "!";
  showPage("page-student");
  setTimeout(function () {
    document.getElementById("ann-overlay").classList.add("open");
  }, 700);
  startLiveCountdown();
  renderStudentExams();
  renderStudentResults();
}

function renderStudentExams() {
  var container = document.getElementById("student-exam-list");
  if (!container) return;
  var schedKeys = Object.keys(S.schedule || {});
  if (!schedKeys.length) {
    container.innerHTML =
      '<div style="color:var(--gray-400);font-size:13px;padding:16px 0;">No exams scheduled yet.</div>';
    return;
  }
  var html = "";
  schedKeys.forEach(function (code) {
    var sc = S.schedule[code];
    var ex = S.exams[code];
    if (!ex) return;
    var status = getExamStatus(sc.date, sc.time, sc.duration);
    var dur = sc.duration || 150;
    var dh = Math.floor(dur / 60),
      dm = dur % 60;
    var durStr = (dh > 0 ? dh + "h " : "") + dm + " min";
    var compKey = S.student.email + "|" + code;
    var isCompleted = S.completed[compKey];
    // Build card
    html += '<div class="exam-card">';
    html += '<div class="ec-head"><div>';
    html += '<div class="ec-title">' + ex.title + "</div>";
    html +=
      '<div class="ec-meta"><span>' +
      (sc.date || "") +
      " &middot; " +
      (sc.time || "") +
      "</span><span>" +
      durStr +
      "</span></div>";
    html += "</div>";
    // Status badge
    if (isCompleted) {
      html += '<span class="ec-status s-done">Completed</span>';
    } else if (status === "live") {
      html += '<span class="ec-status s-live">Live Now</span>';
    } else if (status === "upcoming") {
      html += '<span class="ec-status s-upcoming">Upcoming</span>';
    } else {
      html += '<span class="ec-status s-done">Ended</span>';
    }
    html += "</div>";
    // Countdown bar
    if (!isCompleted && status === "live") {
      html +=
        '<div class="countdown-bar urgent"><span>Exam is live — enter your ID to start</span></div>';
    } else if (!isCompleted && status === "upcoming") {
      html += '<div class="countdown-bar"><span>Not started yet</span></div>';
    }
    // Footer button
    html += '<div class="ec-foot">';
    if (isCompleted) {
      html +=
        '<button class="btn-sm btn-sm-outline" disabled>Submitted</button>';
    } else if (status === "live") {
      html +=
        '<button class="btn-sm btn-sm-pink" onclick="showUnlock()">Unlock &amp; Start &rarr;</button>';
    } else if (status === "upcoming") {
      var info = ex.info ? ex.info : "No info added.";
      html +=
        '<button class="btn-sm btn-sm-outline" onclick="showInfoModal(\'' +
        code +
        "')\">View Info</button>";
    } else {
      html +=
        '<button class="btn-sm btn-sm-outline" disabled>Exam Ended</button>';
    }
    html += "</div></div>";
  });
  container.innerHTML =
    html ||
    '<div style="color:var(--gray-400);font-size:13px;padding:16px 0;">No exams scheduled yet.</div>';
}

// FORGOT PASSWORD
function showForgot() {
  document.getElementById("forgot-email-field").style.display = "none";
  document.getElementById("forgot-phone-field").style.display = "none";
  document.getElementById("fp-success").classList.remove("show");
  document.getElementById("method-email").classList.remove("active");
  document.getElementById("method-phone").classList.remove("active");
  document.getElementById("fp-email").value = "";
  document.getElementById("fp-phone").value = "";
  showPage("page-forgot");
}
function selectMethod(method) {
  document
    .getElementById("method-email")
    .classList.toggle("active", method === "email");
  document
    .getElementById("method-phone")
    .classList.toggle("active", method === "phone");
  document.getElementById("forgot-email-field").style.display =
    method === "email" ? "block" : "none";
  document.getElementById("forgot-phone-field").style.display =
    method === "phone" ? "block" : "none";
  document.getElementById("fp-success").classList.remove("show");
}
async function sendReset(method) {
  var val =
    method === "email"
      ? document.getElementById("fp-email").value.trim()
      : document.getElementById("fp-phone").value.trim();
  if (!val) {
    showToast(
      "Please enter your " +
        (method === "email" ? "email address" : "phone number"),
      "",
    );
    return;
  }
  var filter =
    method === "email" ? { email: val.toLowerCase() } : { phone: val };
  var rows = await dbGet("students", filter);
  document.getElementById("forgot-email-field").style.display = "none";
  document.getElementById("forgot-phone-field").style.display = "none";
  document.getElementById("method-email").classList.remove("active");
  document.getElementById("method-phone").classList.remove("active");
  if (!rows.length || !rows[0].activated) {
    showToast(
      "No account found with that " +
        (method === "email" ? "email" : "phone number"),
      "",
    );
    return;
  }
  var msgEl = document.getElementById("fp-success-msg");
  if (msgEl)
    msgEl.innerHTML =
      'Your password is: <strong style="font-family:monospace;font-size:16px;color:var(--pink);">' +
      rows[0].password +
      "</strong>";
  document.getElementById("fp-success").classList.add("show");
}

// ANNOUNCEMENTS
function closeAnn(e) {
  if (!e || e.target === document.getElementById("ann-overlay"))
    document.getElementById("ann-overlay").classList.remove("open");
}
function postAnn(pin) {
  var t = document.getElementById("ann-t").value.trim(),
    b = document.getElementById("ann-b").value.trim();
  if (!t || !b) {
    showToast("Fill in title and message", "");
    return;
  }
  document.getElementById("ann-title").textContent = t;
  document.getElementById("ann-body").textContent = b;
  if (pin) {
    document.getElementById("pin-title").textContent = t;
    document.getElementById("pin-body").textContent = b;
    document.getElementById("pinned-ann").classList.add("show");
  }
  var hist = document.getElementById("ann-hist"),
    d = document.createElement("div");
  d.style.cssText = "padding:8px 0;border-bottom:1px solid var(--gray-100);";
  d.innerHTML =
    "<strong>" +
    t +
    "</strong>" +
    (pin
      ? '<span class="bdg bdg-pink" style="margin-left:6px;">Pinned</span>'
      : "") +
    '<p style="color:var(--gray-600);margin-top:2px;">' +
    b +
    "</p>";
  hist.insertBefore(d, hist.firstChild);
  showToast(pin ? "Posted & pinned!" : "Posted!", "success");
  document.getElementById("ann-t").value = "";
  document.getElementById("ann-b").value = "";
}
function unpin() {
  document.getElementById("pinned-ann").classList.remove("show");
  showToast("Unpinned", "");
}

// UNLOCK
function showUnlock() {
  document.getElementById("unlock-modal").classList.add("open");
  document.getElementById("unlock-err").style.display = "none";
}
function showInfoModal(examCode) {
  var modal = document.getElementById("info-modal");
  var body = document.getElementById("info-modal-body");
  if (!modal) return;
  if (body) {
    var exam =
      examCode && S.exams && S.exams[examCode] ? S.exams[examCode] : null;
    var sched =
      examCode && S.schedule && S.schedule[examCode]
        ? S.schedule[examCode]
        : null;
    var h =
      '<div style="background:var(--gray-50);border-radius:10px;padding:14px 16px;font-size:13px;line-height:2.2;">';
    if (exam) {
      var dur = exam.duration || 150;
      var dh = Math.floor(dur / 60),
        dm = dur % 60;
      var durStr =
        (dh > 0 ? dh + " hour" + (dh !== 1 ? "s" : "") + " " : "") +
        (dm > 0 ? dm + " min" : "");
      h += "<div><strong>Exam:</strong> " + exam.title + "</div>";
      if (sched)
        h +=
          "<div><strong>Date:</strong> " +
          sched.date +
          (sched.time ? " &middot; " + sched.time : "") +
          "</div>";
      h += "<div><strong>Duration:</strong> " + durStr + "</div>";
      if (exam.info)
        h +=
          '</div><div style="margin-top:10px;background:#fff;border:1.5px solid var(--gray-200);border-radius:10px;padding:12px 14px;font-size:13px;line-height:1.7;color:var(--charcoal-light);">' +
          exam.info;
    } else {
      h += "<div><strong>Subject:</strong> English Language</div>";
      h += "<div><strong>Duration:</strong> 2 hours 30 minutes</div>";
      h += "<div><strong>Total:</strong> 70 points</div>";
    }
    h += "</div>";
    body.innerHTML = h;
  }
  modal.classList.add("open");
}
async function doUnlock() {
  var code = document.getElementById("ex-code").value.trim().toUpperCase();
  var id = document.getElementById("ex-id").value.trim().toUpperCase();
  var err = document.getElementById("unlock-err");
  err.style.display = "none";
  // Check exam exists
  var examRows = await dbGet("exams", { code: code });
  if (!examRows.length) {
    err.textContent = "Invalid exam code.";
    err.style.display = "block";
    return;
  }
  // Check student ID registered
  var idRows = await dbGet("exam_ids", { exam_code: code, student_id: id });
  if (!idRows.length) {
    err.textContent = "Your ID is not registered for this exam.";
    err.style.display = "block";
    return;
  }
  // Check if this ID is already used by a DIFFERENT student
  if (idRows[0].used) {
    var usedByEmail = idRows[0].student_email || null;
    if (usedByEmail && S.student && usedByEmail !== S.student.email) {
      err.textContent = "This ID is already in use by another student.";
      err.style.display = "block";
      return;
    }
    if (!usedByEmail && !S.student) {
      // No way to verify — block for safety
      err.textContent = "This ID has already been used.";
      err.style.display = "block";
      return;
    }
  }
  // Check already submitted by THIS student
  if (S.student) {
    var subRows = await dbGet("submissions", {
      student_email: S.student.email,
      exam_code: code,
    });
    // Only block if it's a real submission (not an autosave in_progress)
    if (subRows.length && subRows[0].elapsed_str !== "in_progress") {
      err.textContent = "This exam has already been submitted.";
      err.style.display = "block";
      return;
    }
  }
  // Load exam data
  var ex = examRows[0];
  S.exams[code] = {
    title: ex.title,
    code: code,
    duration: ex.duration || 150,
    tasks: ex.tasks || [],
    info: ex.info || "",
    parsedExam: ex.parsed_exam || null,
    pts: ex.pts || 70,
    audioURL: ex.audio_url || null,
  };
  if (ex.audio_url) {
    S.audioURL = ex.audio_url;
    window.S_audioURL = ex.audio_url;
  }
  if (ex.answer_key) {
    if (!S.answerKey) S.answerKey = {};
    Object.assign(S.answerKey, ex.answer_key);
  }
  S.examCodes[code] = true;
  if (ex.parsed_exam) applyParsedExam(ex.parsed_exam, ex.title);
  // Load schedule
  var schedRows = await dbGet("schedule", { exam_code: code });
  if (schedRows.length)
    S.schedule[code] = {
      date: schedRows[0].date,
      time: schedRows[0].time,
      duration: schedRows[0].duration || 150,
      title: ex.title,
    };
  // Mark used
  var isReEntry = idRows[0].used;
  if (!isReEntry && S.student) {
    await dbUpdate(
      "exam_ids",
      { used: true, student_email: S.student.email },
      { exam_code: code, student_id: id },
    );
  }
  S.usedIDs[code + ":" + id] = true;
  S.activeExamCode = code;
  closeOverlay("unlock-modal");
  startExam(isReEntry);
}

// EXAM
function _runExamTimer(examCode, startTimeLeft) {
  clearInterval(S.timerInt);
  S.timerPaused = false;
  S.timeLeft = startTimeLeft;
  if (S.timeLeft <= 0) {
    showTimesUp();
    return;
  }
  // Check if admin has paused before we start running
  dbGet("timer_state", { exam_code: examCode }).then(function (rows) {
    if (rows.length && rows[0].paused) {
      // Admin has paused — don't start, just show current time
      S.timerPaused = true;
      var tl = rows[0].time_left || S.timeLeft;
      S.timeLeft = tl;
      var h2 = Math.floor(tl / 3600),
        m2 = Math.floor((tl % 3600) / 60),
        s2 = tl % 60;
      var el2 = document.getElementById("exam-timer");
      if (el2) {
        el2.textContent = h2 + ":" + pad(m2) + ":" + pad(s2);
        el2.className = "exam-timer warn";
      }
      return;
    }
    // Not paused — save our running state
    dbUpsert(
      "timer_state",
      {
        exam_code: examCode,
        time_left: S.timeLeft,
        paused: false,
        started_at: Date.now(),
      },
      "exam_code",
    );
  });
  S.timerInt = setInterval(function () {
    if (S.timerPaused) return;
    S.timeLeft--;
    if (S.timeLeft <= 0) {
      clearInterval(S.timerInt);
      showTimesUp();
      return;
    }
    var h = Math.floor(S.timeLeft / 3600),
      m = Math.floor((S.timeLeft % 3600) / 60),
      s = S.timeLeft % 60;
    var el = document.getElementById("exam-timer");
    if (el) {
      el.textContent = h + ":" + pad(m) + ":" + pad(s);
      el.className =
        "exam-timer" +
        (S.timeLeft < 600 ? " danger" : S.timeLeft < 1800 ? " warn" : "");
    }
    // Sync to Supabase every 15 seconds
    if (S.timeLeft % 15 === 0) {
      dbUpsert(
        "timer_state",
        {
          exam_code: examCode,
          time_left: S.timeLeft,
          paused: false,
          started_at: Date.now(),
        },
        "exam_code",
      );
    }
    // Autosave answers every 30 seconds
    if (S.timeLeft % 30 === 0 && S.student && examCode) {
      var _eta2 = document.getElementById("essay-ta-" + S.curTask);
      if (_eta2) S.answers["t" + S.curTask] = _eta2.value;
      dbUpsert(
        "submissions",
        {
          student_email: S.student.email,
          exam_code: examCode,
          answers: JSON.parse(JSON.stringify(S.answers)),
          elapsed_str: "in_progress",
          exam_start_time: S.examStartTime || 0,
          submitted_at: null,
        },
        "student_email,exam_code",
      );
    }
    // Check Supabase every 3 seconds for admin pause/adjust
    if (S.timeLeft % 3 === 0) {
      (function (curLeft) {
        dbGet("timer_state", { exam_code: examCode }).then(function (rows) {
          if (!rows.length) return;
          var ts = rows[0];
          // Admin paused — stop our timer and start polling for resume
          if (ts.paused && !S.timerPaused) {
            S.timerPaused = true;
            clearInterval(S.timerInt);
            showToast("Exam paused by invigilator", "");
            // Poll every 2 seconds to detect resume
            clearInterval(window._pausePollInt);
            window._pausePollInt = setInterval(function () {
              dbGet("timer_state", { exam_code: examCode }).then(
                function (pRows) {
                  if (!pRows.length || pRows[0].paused) return;
                  clearInterval(window._pausePollInt);
                  var pts = pRows[0];
                  var pElapsed = pts.started_at
                    ? Math.floor((Date.now() - pts.started_at) / 1000)
                    : 0;
                  var pLeft = Math.max(0, (pts.time_left || 0) - pElapsed);
                  S.timerPaused = false;
                  _runExamTimer(examCode, pLeft);
                  showToast("Exam resumed by invigilator", "success");
                },
              );
            }, 2000);
          }
          // Admin resumed (caught by poll above, but also handle if timer was running)
          if (!ts.paused && S.timerPaused) {
            clearInterval(window._pausePollInt);
            var elapsed2 = ts.started_at
              ? Math.floor((Date.now() - ts.started_at) / 1000)
              : 0;
            var newLeft = Math.max(0, (ts.time_left || curLeft) - elapsed2);
            S.timerPaused = false;
            _runExamTimer(examCode, newLeft);
            showToast("Exam resumed by invigilator", "success");
          }
          // Admin adjusted time — sync if difference > 30s
          if (!ts.paused && !S.timerPaused) {
            var elapsed3 = ts.started_at
              ? Math.floor((Date.now() - ts.started_at) / 1000)
              : 0;
            var adminLeft = Math.max(0, ts.time_left - elapsed3);
            if (Math.abs(adminLeft - curLeft) > 30) {
              S.timeLeft = adminLeft;
              if (S.timeLeft <= 0) {
                clearInterval(S.timerInt);
                showTimesUp();
              }
            }
          }
        });
      })(S.timeLeft);
    }
  }, 1000);
}

function startExam(isReEntry) {
  showPage("page-exam");
  buildSidebar();
  var _hasIntro = S.parsedExam && S.parsedExam.hasIntro;
  gotoTask(_hasIntro ? 0 : 1);
  var _activeEx = S.activeExamCode && S.exams[S.activeExamCode];
  var _examDur = _activeEx ? _activeEx.duration || 150 : 150;
  var _examCode = S.activeExamCode;
  var _totalSecs = _examDur * 60;
  clearInterval(S.timerInt);

  if (isReEntry) {
    // Restore saved answers from Supabase
    if (S.student) {
      dbGet("submissions", {
        student_email: S.student.email,
        exam_code: _examCode,
      }).then(function (subRows) {
        if (subRows.length && subRows[0].elapsed_str === "in_progress") {
          S.answers = subRows[0].answers || {
            t1: {},
            t2: {},
            t3: {},
            t4: {},
            t5: {},
            t6: {},
            t7: "",
          };
          showToast("Your previous answers have been restored", "success");
          // Re-render current task to show restored answers
          renderTask(S.curTask);
        }
      });
    }
    // Read remaining time from Supabase timer_state
    dbGet("timer_state", { exam_code: _examCode }).then(function (rows) {
      var startLeft = _totalSecs;
      if (rows.length) {
        var ts = rows[0];
        if (ts.paused) {
          // Exam is currently paused — show paused state, wait for resume
          startLeft = ts.time_left || _totalSecs;
          S.timeLeft = startLeft;
          S.timerPaused = true;
          // Show timer but don't start it
          var h = Math.floor(startLeft / 3600),
            m = Math.floor((startLeft % 3600) / 60),
            s = startLeft % 60;
          var el = document.getElementById("exam-timer");
          if (el) {
            el.textContent = h + ":" + pad(m) + ":" + pad(s);
            el.className = "exam-timer warn";
          }
          showToast("Exam is currently paused by invigilator", "");
          // Poll every 3 seconds until resumed
          var _pollInt = setInterval(function () {
            dbGet("timer_state", { exam_code: _examCode }).then(
              function (rows2) {
                if (!rows2.length || !rows2[0].paused) {
                  clearInterval(_pollInt);
                  var ts2 = rows2[0] || {};
                  var elapsed2 = ts2.started_at
                    ? Math.floor((Date.now() - ts2.started_at) / 1000)
                    : 0;
                  var resumeLeft = Math.max(
                    0,
                    (ts2.time_left || startLeft) - elapsed2,
                  );
                  S.timerPaused = false;
                  _runExamTimer(_examCode, resumeLeft);
                  showToast("Exam resumed", "success");
                }
              },
            );
          }, 3000);
        } else {
          // Running — calculate remaining time
          var elapsed = ts.started_at
            ? Math.floor((Date.now() - ts.started_at) / 1000)
            : 0;
          startLeft = Math.max(0, (ts.time_left || _totalSecs) - elapsed);
          _runExamTimer(_examCode, startLeft);
        }
      } else {
        _runExamTimer(_examCode, startLeft);
      }
    });
  } else {
    // First entry — calculate time from schedule start (timer is shared for all students)
    S.examStartTime = Date.now();
    var sc3 = S.schedule && S.schedule[_examCode];
    if (sc3 && sc3.date && sc3.time) {
      var schedStart = new Date(sc3.date + "T" + sc3.time + ":00").getTime();
      var schedElapsed = Math.floor((Date.now() - schedStart) / 1000);
      if (schedElapsed > 0 && schedElapsed < _totalSecs) {
        // Exam already started — join at correct remaining time
        _runExamTimer(_examCode, Math.max(0, _totalSecs - schedElapsed));
      } else if (schedElapsed >= _totalSecs) {
        // Time already expired
        showTimesUp();
      } else {
        // Exam hasn't started yet or just started
        _runExamTimer(_examCode, _totalSecs);
      }
    } else {
      // Also check Supabase timer_state for existing timer
      dbGet("timer_state", { exam_code: _examCode }).then(function (rows) {
        if (rows.length && !rows[0].paused) {
          var elapsed2 = rows[0].started_at
            ? Math.floor((Date.now() - rows[0].started_at) / 1000)
            : 0;
          var left2 = Math.max(0, rows[0].time_left - elapsed2);
          _runExamTimer(_examCode, left2 > 0 ? left2 : _totalSecs);
        } else {
          _runExamTimer(_examCode, _totalSecs);
        }
      });
    }
  }
  if (window.pdfjsLib)
    window.pdfjsLib.GlobalWorkerOptions.workerSrc =
      "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
  if (S.pdfFile && window.pdfjsLib) {
    var reader = new FileReader();
    reader.onload = function (e) {
      window.pdfjsLib
        .getDocument({ data: new Uint8Array(e.target.result) })
        .promise.then(function (doc) {
          S.pdfDoc = doc;
          renderTask(S.curTask);
        });
    };
    reader.readAsArrayBuffer(S.pdfFile);
  }
}

function gotoTask(n) {
  var _eta = document.getElementById("essay-ta-" + S.curTask);
  if (_eta) S.answers["t" + S.curTask] = _eta.value;
  S.curTask = n;
  for (var i = 0; i <= 7; i++) {
    var p = document.getElementById("pill-" + i);
    if (p) {
      p.classList.remove("active");
      if (i === n) p.classList.add("active");
    }
  }
  renderTask(n);
  var main = document.getElementById("exam-main");
  if (main) main.scrollTop = 0;
}

var TASKS = {};

function renderTask(n) {
  var main = document.getElementById("exam-main");
  if (!main) return;
  var _eta = document.getElementById("essay-ta-" + S.curTask);
  if (_eta) S.answers["t" + S.curTask] = _eta.value;

  if (n === 0) {
    var h = '<div class="task-card">';
    if (S.pdfDoc) {
      h +=
        '<div class="pdf-page-wrap"><canvas id="pdf-canvas-0"></canvas></div>';
    } else {
      var _iex =
        S.activeExamCode && S.exams[S.activeExamCode]
          ? S.exams[S.activeExamCode]
          : null;
      var _ititle = _iex ? _iex.title : "English Exam";
      var _iinfo = _iex && _iex.info ? _iex.info : "";
      var _idur = _iex ? _iex.duration || 150 : 150;
      var _idh = Math.floor(_idur / 60),
        _idm = _idur % 60;
      var _idurStr =
        (_idh > 0 ? _idh + " hour" + (_idh !== 1 ? "s " : " ") : "") +
        (_idm > 0 ? _idm + " min" : "");
      var _iparsedIntro =
        S.parsedExam && S.parsedExam.introText ? S.parsedExam.introText : "";
      h += '<div style="padding:24px 20px;">';
      h +=
        '<div style="font-size:22px;font-weight:700;margin-bottom:12px;">' +
        _ititle +
        "</div>";
      if (_iparsedIntro) {
        h +=
          '<div style="font-size:14px;line-height:1.9;color:#444;white-space:pre-wrap;">' +
          _iparsedIntro +
          "</div>";
      } else {
        h +=
          '<div style="background:var(--gray-50);border-radius:10px;padding:14px 16px;font-size:13px;line-height:2;margin-bottom:12px;">';
        h += "<div><strong>Duration:</strong> " + _idurStr + "</div>";
        if (_iinfo) h += '<div style="margin-top:8px;">' + _iinfo + "</div>";
        h += "</div>";
      }
      h += "</div>";
    }
    h +=
      '</div><div class="exam-nav-btns">' +
      '<button class="geo-btn geo-btn-pink" onclick="gotoTask(1)">\u10D2\u10D0\u10DB\u10DD\u10EA\u10D3\u10D8\u10E1 \u10D3\u10D0\u10EC\u10E7\u10D4\u10D1\u10D0 &rarr;</button>' +
      "</div>";
    main.innerHTML = h;
    if (S.pdfDoc) renderPdfPage(1, "pdf-canvas-0");
    return;
  }

  var task = TASKS[n];
  // If TASKS not populated yet, try to apply parsedExam
  if (!task && S.parsedExam && S.parsedExam.tasks) {
    applyParsedExam(S.parsedExam, "");
    task = TASKS[n];
  }
  // If still no task, check S.exams for tasks array
  if (
    !task &&
    S.activeExamCode &&
    S.exams[S.activeExamCode] &&
    S.exams[S.activeExamCode].tasks
  ) {
    var _exTasks = S.exams[S.activeExamCode].tasks;
    if (_exTasks[n - 1]) {
      TASKS[n] = {
        title: "Task " + n,
        type: _exTasks[n - 1].type || "reading",
        pts: _exTasks[n - 1].pts || _exTasks[n - 1].points || 0,
        instructions: _exTasks[n - 1].instructions || "",
        qs: [],
        opts: ["A", "B", "C", "D"],
      };
      task = TASKS[n];
    }
  }
  if (!task) {
    // Last resort — show submit button so student isn't stuck
    var h2 =
      '<div class="task-card"><p style="color:var(--gray-400);text-align:center;padding:40px;">Loading task...</p></div>';
    h2 += '<div class="exam-nav-btns">';
    if (n > 0)
      h2 +=
        '<button class="geo-btn geo-btn-outline" onclick="gotoTask(' +
        (n - 1) +
        ')">\u2190 \u10E3\u10D9\u10D0\u10DC \u10D3\u10D0\u10D1\u10E0\u10E3\u10DC\u10D4\u10D1\u10D0</button>';
    h2 +=
      '<button type="button" class="geo-btn geo-btn-pink" onclick="showSubmitConfirm()">\u10D2\u10D0\u10DB\u10DD\u10EA\u10D3\u10D8\u10E1 \u10D3\u10D0\u10E1\u10E0\u10E3\u10DA\u10D4\u10D1\u10D0 \u2713</button>';
    h2 += "</div>";
    main.innerHTML = h2;
    return;
  }
  if (task.type === "sheet") {
    renderSheet(main);
    return;
  }

  var h = '<div class="task-card">';
  h +=
    '<span class="task-badge">' +
    task.title +
    '</span><div class="task-pts">' +
    task.pts +
    " points</div>";
  if (S.pdfDoc)
    h +=
      '<div class="pdf-page-wrap"><canvas id="pdf-canvas-' +
      n +
      '"></canvas></div>';

  if (task.type === "listening") {
    h +=
      '<p class="task-instr">' +
      (task.instructions ||
        "Listen and mark the correct answer A, B, C or D.") +
      "</p>";
    // Get audio URL from exam data or stored URL
    var _activeExForAudio = S.activeExamCode && S.exams[S.activeExamCode];
    var _audioSrc =
      window.S_audioURL ||
      S.audioURL ||
      (_activeExForAudio && _activeExForAudio.audioURL) ||
      "";
    if (_audioSrc) {
      h +=
        '<audio id="task-audio" src="' +
        _audioSrc +
        '" preload="metadata" style="display:none;"></audio>';
      h += '<div class="audio-player">';
      h += '<button class="audio-btn" onclick="playAudio(this)">';
      h +=
        '<svg viewBox="0 0 24 24"><polygon points="5 3 19 12 5 21 5 3"/></svg>';
      h += "</button>";
      h +=
        '<div class="audio-bar"><div class="audio-prog" id="ap"></div></div>';
      h +=
        '<span class="audio-time" id="audio-time-display">0:00 / 0:00</span>';
      h += "</div>";
    } else {
      h +=
        '<div style="background:var(--gray-50);border:1.5px solid var(--gray-200);border-radius:10px;padding:14px 16px;font-size:13px;color:var(--gray-500);margin-bottom:16px;">Audio will be played during the exam. Please wait for instructions.</div>';
    }
    for (var i = 0; i < task.qs.length; i++) {
      var q = task.qs[i];
      h +=
        '<div class="q-block"><div class="q-text"><span class="q-num-inline">' +
        (i + 1) +
        ".</span> " +
        q.q +
        '</div><div class="opts">';
      for (var j = 0; j < task.opts.length; j++) {
        var l = task.opts[j],
          sel =
            S.answers["t" + n] && S.answers["t" + n][i + 1] === l ? " sel" : "";
        h +=
          '<div class="opt' +
          sel +
          '" onclick="selAns(' +
          n +
          "," +
          (i + 1) +
          ",&#39;" +
          l +
          '&#39;,this)"><div class="opt-letter">' +
          l +
          "</div><span>" +
          q.os[j] +
          "</span></div>";
      }
      h += "</div></div>";
    }
  }
  if (task.type === "match") {
    h +=
      '<p class="task-instr">' +
      (task.instructions ||
        "Read the questions and find answers in the paragraphs.") +
      "</p>";
    // Questions first
    for (var i = 0; i < task.qs.length; i++) {
      var selLet =
        S.answers["t" + n] && S.answers["t" + n][i + 1]
          ? S.answers["t" + n][i + 1]
          : "";
      h +=
        '<div class="q-block"><div class="q-text"><span class="q-num-inline">' +
        (i + 1) +
        ".</span> " +
        task.qs[i].q +
        "</div>";
      h += '<div class="opts" style="flex-wrap:wrap;gap:6px;margin-top:6px;">';
      for (var mi = 0; mi < task.opts.length; mi++) {
        var ml = task.opts[mi],
          msel = selLet === ml ? " sel" : "";
        h +=
          '<div class="opt-letter-only' +
          msel +
          '" onclick="selAns(' +
          n +
          "," +
          (i + 1) +
          ",'" +
          ml +
          '\',this)" style="width:36px;height:36px;">' +
          ml +
          "</div>";
      }
      h += "</div></div>";
    }
    // Passages after questions
    h += '<div class="reading-box" style="margin-top:16px;">';
    var pk = Object.keys(task.passages || {});
    for (var pi = 0; pi < pk.length; pi++)
      h +=
        "<p><strong>" + pk[pi] + ".</strong> " + task.passages[pk[pi]] + "</p>";
    h += "</div>";
  }
  if (task.type === "reading") {
    h +=
      '<p class="task-instr">' +
      (task.instructions ||
        "Read the text and mark the correct answer A, B, C or D.") +
      "</p>";
    if (task.text)
      h +=
        '<div class="reading-box" style="margin-bottom:16px;">' +
        task.text +
        "</div>";
    for (var i = 0; i < task.qs.length; i++) {
      var q = task.qs[i];
      h +=
        '<div class="q-block"><div class="q-text"><span class="q-num-inline">' +
        (i + 1) +
        ".</span> " +
        q.q +
        '</div><div class="opts">';
      for (var j = 0; j < task.opts.length; j++) {
        var l = task.opts[j],
          sel =
            S.answers["t" + n] && S.answers["t" + n][i + 1] === l ? " sel" : "";
        h +=
          '<div class="opt' +
          sel +
          '" onclick="selAns(' +
          n +
          "," +
          (i + 1) +
          ",\x27" +
          l +
          '\x27,this)"><div class="opt-letter">' +
          l +
          "</div><span>" +
          q.os[j] +
          "</span></div>";
      }
      h += "</div></div>";
    }
  }
  if (task.type === "gapfill4") {
    var wb = task.wordBank || {};
    var wks = Object.keys(wb);
    h +=
      '<p class="task-instr">' +
      (task.instructions ||
        "Drag each word into the correct gap, or tap a word then tap a gap.") +
      "</p>";
    // Word bank chips
    h += '<div class="wb-chips" id="wb-chips-t' + n + '">';
    for (var wi = 0; wi < wks.length; wi++) {
      var wkey = wks[wi],
        wval = wb[wkey];
      var placed =
        S.answers["t" + n] && S.objectHasValue(S.answers["t" + n], wkey);
      h +=
        '<span class="wb-chip' +
        (placed ? " wb-placed" : "") +
        '" data-key="' +
        wkey +
        '" data-val="' +
        wval +
        '" draggable="true" onclick="wbChipTap(' +
        n +
        ",'" +
        wkey +
        "',this)\">" +
        wval +
        "</span>";
    }
    h += "</div>";
    // Article text with embedded gaps OR numbered gap list
    if (task.text) {
      // Replace ......(N) markers with interactive gap spans
      var txt = task.text;
      var gapcount = task.qcount || 12;
      for (var gi = 1; gi <= gapcount; gi++) {
        var curKey =
          S.answers["t" + n] && S.answers["t" + n][gi]
            ? S.answers["t" + n][gi]
            : null;
        var curWord = curKey && wb[curKey] ? wb[curKey] : null;
        var gapSpan =
          '<span class="gap-drop' +
          (curWord ? " gap-filled" : "") +
          '" id="gap-t' +
          n +
          "-" +
          gi +
          '" data-tn="' +
          n +
          '" data-gn="' +
          gi +
          '" ondragover="event.preventDefault()" ondrop="dropOnGap(event,' +
          n +
          "," +
          gi +
          ')" onclick="gapTap(' +
          n +
          "," +
          gi +
          ',this)">' +
          (curWord || '<span class="gap-arrow">&#9660;</span>') +
          "</span>";
        if (txt.indexOf("......(" + gi + ")") !== -1) {
          txt = txt.replace("......(" + gi + ")", gapSpan);
        } else {
          txt = txt.replace("(" + gi + ")", gapSpan);
        }
      }
      h += '<div class="gapfill-text reading-box">' + txt + "</div>";
    } else {
      // Fallback: show numbered gaps without text
      h += '<div class="gapfill-text">';
      var gapcount2 = task.qcount || 12;
      for (var gi2 = 1; gi2 <= gapcount2; gi2++) {
        var curKey2 =
          S.answers["t" + n] && S.answers["t" + n][gi2]
            ? S.answers["t" + n][gi2]
            : null;
        var curWord2 = curKey2 && wb[curKey2] ? wb[curKey2] : null;
        h +=
          '<span style="margin-right:8px;">' +
          gi2 +
          '. </span><span class="gap-drop' +
          (curWord2 ? " gap-filled" : "") +
          '" id="gap-t' +
          n +
          "-" +
          gi2 +
          '" data-tn="' +
          n +
          '" data-gn="' +
          gi2 +
          '" ondragover="event.preventDefault()" ondrop="dropOnGap(event,' +
          n +
          "," +
          gi2 +
          ')" onclick="gapTap(' +
          n +
          "," +
          gi2 +
          ',this)">' +
          (curWord2 || '<span class="gap-arrow">&#9660;</span>') +
          "</span><br>";
      }
      h += "</div>";
    }
  }
  if (task.type === "gapfill5") {
    var wc = task.wordChoices || [];
    h +=
      '<p class="task-instr">' +
      (task.instructions || "Choose the correct word for each gap.") +
      "</p>";
    if (task.text) {
      // Replace ......(N) with inline dropdowns
      var txt5 = task.text;
      var gapcount5 = task.qcount || wc.length;
      for (var i5 = 1; i5 <= gapcount5; i5++) {
        var ca5 =
          S.answers["t" + n] && S.answers["t" + n][i5]
            ? S.answers["t" + n][i5]
            : "";
        var choices5 = wc[i5 - 1] || [];
        var dd5 =
          '<select class="inline-dd" data-tn="' +
          n +
          '" data-q="' +
          i5 +
          '" onchange="selDD(' +
          n +
          "," +
          i5 +
          ',this)">';
        dd5 += '<option value="">(' + i5 + ")</option>";
        for (var j5 = 0; j5 < choices5.length; j5++) {
          var opt5 = choices5[j5];
          // strip leading "A. " etc if present
          var optval5 = opt5.replace(/^[A-D]\.\s*/, "");
          var optlbl5 = String.fromCharCode(65 + j5) + ". " + optval5;
          dd5 +=
            '<option value="' +
            String.fromCharCode(65 + j5) +
            '"' +
            (ca5 === String.fromCharCode(65 + j5) ? " selected" : "") +
            ">" +
            optlbl5 +
            "</option>";
        }
        dd5 += "</select>";
        if (txt5.indexOf("......(" + i5 + ")") !== -1) {
          txt5 = txt5.replace("......(" + i5 + ")", dd5);
        } else {
          txt5 = txt5.replace("(" + i5 + ")", dd5);
        }
      }
      h += '<div class="gapfill-text reading-box">' + txt5 + "</div>";
    } else {
      // Fallback: list each gap
      h += '<div class="gapfill-text">';
      for (var i6 = 0; i6 < wc.length; i6++) {
        var ca6b =
          S.answers["t" + n] && S.answers["t" + n][i6 + 1]
            ? S.answers["t" + n][i6 + 1]
            : "";
        h += '<span style="margin-right:6px;">' + (i6 + 1) + ".</span>";
        h +=
          '<select class="inline-dd" data-tn="' +
          n +
          '" data-q="' +
          (i6 + 1) +
          '" onchange="selDD(' +
          n +
          "," +
          (i6 + 1) +
          ',this)">';
        h += '<option value="">(' + (i6 + 1) + ")</option>";
        var choices6b = wc[i6] || [];
        var ls5 = ["A", "B", "C", "D"];
        for (var j6 = 0; j6 < ls5.length; j6++) {
          var w6 = (choices6b[j6] || ls5[j6]).replace(/^[A-D]\.\s*/, "");
          h +=
            '<option value="' +
            ls5[j6] +
            '"' +
            (ca6b === ls5[j6] ? " selected" : "") +
            ">" +
            ls5[j6] +
            ". " +
            w6 +
            "</option>";
        }
        h += "</select><br>";
      }
      h += "</div>";
    }
  }
  if (task.type === "dialogue") {
    h +=
      '<p class="task-instr">' +
      (task.instructions ||
        "Complete the conversation by choosing the correct sentence for each gap.") +
      "</p>";
    var sk6 = Object.keys(task.dialogueSents || {});
    var dl6 = task.dialogue || [];
    h += '<div style="display:flex;gap:16px;flex-wrap:wrap;">';
    h += '<div style="flex:2;min-width:220px;">';
    h += '<div class="reading-box">';
    var gn6 = 0;
    for (var di6 = 0; di6 < dl6.length; di6++) {
      var line6 = dl6[di6];
      // Detect gap: ...(N) or just (N)
      var _hasGap6 = line6.indexOf("...(") !== -1 || /\(\d+\)/.test(line6);
      if (_hasGap6) {
        gn6++;
        var ca6 =
          S.answers["t" + n] && S.answers["t" + n][gn6]
            ? S.answers["t" + n][gn6]
            : "";
        var dd6 =
          '<select class="inline-dd" data-tn="' +
          n +
          '" data-q="' +
          gn6 +
          '" onchange="selDD(' +
          n +
          "," +
          gn6 +
          ',this)">';
        dd6 += '<option value="">(' + gn6 + ")</option>";
        for (var sk6i = 0; sk6i < sk6.length; sk6i++) {
          var sl6 = sk6[sk6i];
          dd6 +=
            '<option value="' +
            sl6 +
            '"' +
            (ca6 === sl6 ? " selected" : "") +
            ">" +
            sl6 +
            "</option>";
        }
        dd6 += "</select>";
        if (line6.indexOf("...(" + gn6 + ")") !== -1) {
          line6 = line6.replace("...(" + gn6 + ")", dd6);
        } else {
          line6 = line6.replace("(" + gn6 + ")", dd6);
        }
      }
      h += '<p style="font-size:13px;margin:6px 0;">' + line6 + "</p>";
    }
    h += "</div></div>";
    h +=
      '<div style="flex:1;min-width:160px;background:var(--gray-50);border-radius:10px;padding:12px;">';
    h +=
      '<div style="font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:var(--gray-400);margin-bottom:8px;">Sentences</div>';
    for (var sk6j = 0; sk6j < sk6.length; sk6j++) {
      var sk6key = sk6[sk6j];
      var used6 =
        S.answers["t" + n] && S.objectHasValue(S.answers["t" + n], sk6key);
      h +=
        '<div style="font-size:12px;margin:4px 0;padding:4px 6px;border-radius:6px;' +
        (used6 ? "opacity:.4;" : "") +
        '"><strong>' +
        sk6key +
        ".</strong> " +
        task.dialogueSents[sk6key] +
        "</div>";
    }
    h += "</div></div>";
  }
  if (task.type === "essay" || task.type === "writing") {
    var _essayInstr = task.instructions || "Write your response.";
    var _essayPrompt = task.prompt || "";
    var _essayAns = S.answers["t" + n] || "";
    h += '<p class="task-instr">' + _essayInstr + "</p>";
    if (_essayPrompt)
      h += '<div class="essay-prompt">' + _essayPrompt + "</div>";
    h +=
      '<textarea class="essay-ta" id="essay-ta-' +
      n +
      '" placeholder="Write your response here...">' +
      _essayAns +
      "</textarea>";
    h += '<div class="wcount" id="wc">0 words</div>';
  }
  h += "</div>";

  // Determine last task number dynamically
  // Determine last task number
  var _lastTask = 7;
  if (S.parsedExam && S.parsedExam.tasks && S.parsedExam.tasks.length) {
    _lastTask = S.parsedExam.tasks.length;
  } else {
    var _activeEx2 = S.activeExamCode && S.exams[S.activeExamCode];
    if (_activeEx2 && _activeEx2.tasks && _activeEx2.tasks.length) {
      _lastTask = _activeEx2.tasks.length;
    } else {
      var _tkeys2 = Object.keys(TASKS)
        .map(Number)
        .filter(function (k) {
          return k >= 1 && TASKS[k] && TASKS[k].type;
        });
      if (_tkeys2.length) _lastTask = Math.max.apply(null, _tkeys2);
    }
  }
  h += '<div class="exam-nav-btns">';
  if (n > 0)
    h +=
      '<button class="geo-btn geo-btn-outline" onclick="gotoTask(' +
      (n - 1) +
      ')">\u2190 \u10E3\u10D9\u10D0\u10DC \u10D3\u10D0\u10D1\u10E0\u10E3\u10DC\u10D4\u10D1\u10D0</button>';
  if (n < _lastTask)
    h +=
      '<button class="geo-btn geo-btn-pink" onclick="gotoTask(' +
      (n + 1) +
      ')">\u10E8\u10D4\u10DB\u10D3\u10D4\u10D2\u10D6\u10D4 \u10D2\u10D0\u10D3\u10D0\u10E1\u10D5\u10DA\u10D0 \u2192</button>';
  else
    h +=
      '<button type="button" class="geo-btn geo-btn-pink" style="min-width:160px;" onclick="console.log(\'Submit btn clicked\'); showSubmitConfirm();">\u10D2\u10D0\u10DB\u10DD\u10EA\u10D3\u10D8\u10E1 \u10D3\u10D0\u10E1\u10E0\u10E3\u10DA\u10D4\u10D1\u10D0 \u2713</button>';
  h += "</div>";

  main.innerHTML = h;
  if (S.pdfDoc) renderPdfPage(n + 1, "pdf-canvas-" + n);
  if (task && task.type === "gapfill4") {
    setTimeout(initDragDrop, 50);
  }
  if (task.type === "essay" || task.type === "writing") {
    var _taId = "essay-ta-" + n;
    var ta = document.getElementById(_taId);
    if (ta) {
      ta.addEventListener(
        "input",
        (function (taskN, taEl) {
          return function () {
            S.answers["t" + taskN] = taEl.value;
            updWC();
          };
        })(n, ta),
      );
      updWC();
    }
  }
}

function renderPdfPage(pageNum, canvasId) {
  if (!S.pdfDoc) return;
  S.pdfDoc
    .getPage(pageNum)
    .then(function (page) {
      var canvas = document.getElementById(canvasId);
      if (!canvas) return;
      var vp = page.getViewport({ scale: 1 });
      var scale = Math.min(
        (canvas.parentElement.offsetWidth || 700) / vp.width,
        1.4,
      );
      var viewport = page.getViewport({ scale: scale });
      canvas.width = viewport.width;
      canvas.height = viewport.height;
      page.render({
        canvasContext: canvas.getContext("2d"),
        viewport: viewport,
      });
    })
    .catch(function () {});
}

function checkAnswers(tn) {
  var answers = S.answers["t" + tn] || {};
  var qcount = 0;
  if (S.parsedExam && S.parsedExam.tasks)
    for (var ti = 0; ti < S.parsedExam.tasks.length; ti++) {
      if (S.parsedExam.tasks[ti].task === tn) {
        qcount = S.parsedExam.tasks[ti].question_count || 0;
        break;
      }
    }
  if (!qcount && TASKS[tn] && TASKS[tn].qs) qcount = TASKS[tn].qs.length;
  var correct = 0;
  var sc = document.getElementById("sidebar-content");
  if (sc) {
    var rows = sc.querySelectorAll('.ans-row[data-tn="' + tn + '"]');
    var opts = [
      "A",
      "B",
      "C",
      "D",
      "E",
      "F",
      "G",
      "H",
      "I",
      "J",
      "K",
      "L",
      "M",
      "N",
    ];
    for (var ri = 0; ri < rows.length; ri++) {
      var row = rows[ri],
        q = parseInt(row.getAttribute("data-q"));
      var correctLetter = S.answerKey && S.answerKey["ak-ak" + tn + "-" + q];
      var studentLetter = answers[q];
      var boxes = row.querySelectorAll(".ans-box");
      for (var bi = 0; bi < boxes.length; bi++) {
        boxes[bi].classList.remove("correct", "wrong");
        if (opts[bi] === correctLetter) boxes[bi].classList.add("correct");
        if (opts[bi] === studentLetter && studentLetter !== correctLetter)
          boxes[bi].classList.add("wrong");
      }
      if (studentLetter && studentLetter === correctLetter) correct++;
    }
  }
  var btn = document.getElementById("check-btn-" + tn);
  if (btn) {
    btn.textContent =
      correct + "/" + qcount + " \u10E1\u10EC\u10DD\u10E0\u10D8";
    btn.disabled = true;
    btn.style.background = "#16a34a";
    btn.style.color = "#fff";
    btn.style.borderColor = "#16a34a";
  }
  showToast(
    correct + " / " + qcount + " \u10E1\u10EC\u10DD\u10E0\u10D8",
    "success",
  );
}

function selFull(key, q, letter, el) {
  S.answers[key][q] = letter;
  var row = el.closest(".fs-row"),
    bs = row.querySelectorAll(".fs-box");
  for (var i = 0; i < bs.length; i++) bs[i].classList.remove("sel");
  el.classList.add("sel");
  syncSidebar(parseInt(key.replace("t", "")), q, letter);
}
function syncSidebar(tn, q, letter) {
  var row = document.querySelector(
    '#sidebar-content .ans-row[data-tn="' + tn + '"][data-q="' + q + '"]',
  );
  if (!row) return;
  var bs = row.querySelectorAll(".ans-box");
  for (var i = 0; i < bs.length; i++) bs[i].classList.remove("sel");
  var box = row.querySelector('.ans-box[data-l="' + letter + '"]');
  if (box) box.classList.add("sel");
}
function markPillDone(tn) {
  var p = document.getElementById("pill-" + tn);
  if (p && !p.classList.contains("active")) p.classList.add("done");
}
function buildSidebar() {
  var sc = document.getElementById("sidebar-content");
  if (!sc) return;
  // Build sidebar config from parsed TASKS or hardcoded defaults
  var cfg = [];
  var taskNums = Object.keys(TASKS)
    .map(Number)
    .filter(function (k) {
      return k >= 1 && k <= 8;
    })
    .sort(function (a, b) {
      return a - b;
    });
  for (var tni = 0; tni < taskNums.length; tni++) {
    var tnum = taskNums[tni];
    var tk = TASKS[tnum];
    if (!tk || tk.type === "essay" || tk.type === "sheet") continue;
    var qcount = tk.qs ? tk.qs.length : tk.qcount || 0;
    var opts_ = tk.opts || ["A", "B", "C", "D"];
    // For gapfill4/5 and dialogue the sidebar just shows gap numbers
    if (
      tk.type === "gapfill4" ||
      tk.type === "gapfill5" ||
      tk.type === "dialogue"
    ) {
      opts_ = [
        "A",
        "B",
        "C",
        "D",
        "E",
        "F",
        "G",
        "H",
        "I",
        "J",
        "K",
        "L",
        "M",
        "N",
      ].slice(0, opts_.length);
    }
    cfg.push({
      key: "t" + tnum,
      label: "Task " + tnum,
      q: qcount,
      opts: opts_,
    });
  }
  if (!cfg.length) {
    cfg = [
      { key: "t1", label: "Task 1", q: 8, opts: ["A", "B", "C", "D"] },
      {
        key: "t2",
        label: "Task 2",
        q: 8,
        opts: ["A", "B", "C", "D", "E", "F"],
      },
      { key: "t3", label: "Task 3", q: 8, opts: ["A", "B", "C", "D"] },
      {
        key: "t4",
        label: "Task 4",
        q: 12,
        opts: [
          "A",
          "B",
          "C",
          "D",
          "E",
          "F",
          "G",
          "H",
          "I",
          "J",
          "K",
          "L",
          "M",
          "N",
        ],
      },
      { key: "t5", label: "Task 5", q: 12, opts: ["A", "B", "C", "D"] },
      {
        key: "t6",
        label: "Task 6",
        q: 6,
        opts: ["A", "B", "C", "D", "E", "F", "G", "H"],
      },
    ];
  }
  var h = "";
  for (var ci = 0; ci < cfg.length; ci++) {
    var c = cfg[ci],
      tn = parseInt(c.key.replace("t", ""));
    h +=
      '<div class="asb"><div class="asb-head">' +
      c.label +
      " \u2014 " +
      c.opts.join(" ") +
      '</div><div class="asb-body">';
    for (var q = 1; q <= c.q; q++) {
      h +=
        '<div class="ans-row" data-tn="' +
        tn +
        '" data-q="' +
        q +
        '"><span class="ans-qn">' +
        q +
        "</span>";
      for (var li = 0; li < c.opts.length; li++) {
        var l = c.opts[li],
          sel = S.answers[c.key][q] === l ? " sel" : "";
        h +=
          '<div class="ans-box' +
          sel +
          '" data-l="' +
          l +
          '" onclick="sidebarClick(' +
          "'" +
          c.key +
          "'" +
          "," +
          q +
          "," +
          "'" +
          l +
          "'" +
          ',this)">' +
          l +
          "</div>";
      }
      h += "</div>";
    }
    h += "</div></div>";
  }
  h +=
    '<div class="asb"><div class="asb-head">Task 7 \u2014 Essay</div><div class="asb-body" style="font-size:11px;color:var(--gray-600);padding:8px 10px;">Written on Task 7 page</div></div>';
  sc.innerHTML = h;
}
function sidebarClick(key, q, letter, el) {
  S.answers[key][q] = letter;
  var row = el.closest(".ans-row"),
    bs = row.querySelectorAll(".ans-box");
  for (var i = 0; i < bs.length; i++) bs[i].classList.remove("sel");
  el.classList.add("sel");
  var tn = parseInt(key.replace("t", ""));
  markPillDone(tn);
  if (S.curTask === tn) {
    var blocks = document.querySelectorAll(".q-block");
    if (blocks[q - 1]) {
      var os = blocks[q - 1].querySelectorAll(".opt");
      for (var i = 0; i < os.length; i++) os[i].classList.remove("sel");
      var ols = blocks[q - 1].querySelectorAll(".opt-letter");
      for (var i = 0; i < ols.length; i++) {
        if (ols[i].textContent === letter)
          ols[i].closest(".opt").classList.add("sel");
      }
    }
  }
}
function updWC() {
  var ta = document.getElementById("essay-ta-" + S.curTask);
  if (!ta) ta = document.getElementById("essay-ta");
  if (!ta) return;
  var w = ta.value.trim() === "" ? 0 : ta.value.trim().split(/\s+/).length;
  var el = document.getElementById("wc");
  if (!el) return;
  el.className =
    "wcount" + (w >= 120 && w <= 170 ? " good" : w > 170 ? " over" : "");
  el.textContent = w + " words (120-170 required)";
}
function playAudio(btn) {
  var audio = document.getElementById("task-audio");
  if (!audio) return;
  var prog = document.getElementById("ap");
  var timeEl = btn
    ? btn.closest(".audio-player").querySelector(".audio-time")
    : null;
  if (audio.paused) {
    audio.play();
    // Switch icon to pause
    btn.innerHTML =
      '<svg viewBox="0 0 24 24" fill="#fff" width="12" height="12"><rect x="6" y="4" width="4" height="16"/><rect x="14" y="4" width="4" height="16"/></svg>';
  } else {
    audio.pause();
    btn.innerHTML =
      '<svg viewBox="0 0 24 24"><polygon points="5 3 19 12 5 21 5 3"/></svg>';
  }
  audio.ontimeupdate = function () {
    if (!audio.duration) return;
    var pct = (audio.currentTime / audio.duration) * 100;
    if (prog) prog.style.width = pct + "%";
    if (timeEl) {
      var cur = Math.floor(audio.currentTime),
        dur = Math.floor(audio.duration);
      var cm = Math.floor(cur / 60),
        cs = cur % 60,
        dm = Math.floor(dur / 60),
        ds = dur % 60;
      timeEl.textContent =
        cm +
        ":" +
        (cs < 10 ? "0" : "") +
        cs +
        " / " +
        dm +
        ":" +
        (ds < 10 ? "0" : "") +
        ds;
    }
  };
  audio.onended = function () {
    if (prog) prog.style.width = "0%";
    btn.innerHTML =
      '<svg viewBox="0 0 24 24"><polygon points="5 3 19 12 5 21 5 3"/></svg>';
  };
}

// ADMIN
function showSec(name) {
  var secs = document.querySelectorAll(".asec");
  for (var i = 0; i < secs.length; i++) secs[i].classList.remove("active");
  var navs = document.querySelectorAll(".admin-nav-item");
  for (var i = 0; i < navs.length; i++) navs[i].classList.remove("active");
  var sec = document.getElementById("sec-" + name);
  if (sec) sec.classList.add("active");
  var btn = document.querySelector('.admin-nav-item[data-sec="' + name + '"]');
  if (btn) btn.classList.add("active");
  if (name === "exams") renderExamList();
  if (name === "tasks") renderTaskList();
  if (name === "grading") buildEssays();
  if (name === "schedule") renderSchedTable();
  if (name === "students") refreshStudentsTable();
  if (name === "codes") refreshCodesTable();
  if (name === "results") renderResultsTable();
  if (name === "schedule") renderSchedTable();
}

function renderSchedTable() {
  var tbody = document.getElementById("sched-tbody");
  if (!tbody) return;
  var keys = Object.keys(S.schedule || {});
  if (!keys.length) {
    tbody.innerHTML =
      '<tr><td colspan="6" style="text-align:center;color:var(--gray-400);padding:20px;font-size:13px;">No exams scheduled yet.</td></tr>';
    return;
  }
  var rows = "";
  for (var i = 0; i < keys.length; i++) {
    var code = keys[i];
    var sc = S.schedule[code];
    var ex = S.exams[code];
    var status = getExamStatus(sc.date, sc.time, sc.duration);
    var dur = sc.duration || 150;
    var dh = Math.floor(dur / 60),
      dm = dur % 60;
    var durStr = (dh > 0 ? dh + "h " : "") + dm + "min";
    var badgeHtml =
      status === "live"
        ? '<span class="bdg bdg-green">Live</span>'
        : status === "upcoming"
          ? '<span class="bdg bdg-orange">Upcoming</span>'
          : '<span class="bdg" style="background:var(--gray-100);color:var(--gray-600);">Ended</span>';
    rows +=
      "<tr>" +
      "<td><strong>" +
      (ex ? ex.title : code) +
      "</strong></td>" +
      '<td><span class="code-tag">' +
      code +
      "</span></td>" +
      "<td>" +
      (sc.date || "-") +
      "</td>" +
      "<td>" +
      (sc.time || "-") +
      "</td>" +
      "<td>" +
      durStr +
      "</td>" +
      "<td>" +
      badgeHtml +
      "</td>" +
      '<td><button class="abtn abtn-ghost" style="padding:3px 10px;font-size:11px;" onclick="openExamEdit(\'' +
      code +
      "')\">" +
      (status === "live" ? "Manage" : status === "ended" ? "Renew" : "Edit") +
      "</button></td>" +
      "</tr>";
  }
  tbody.innerHTML = rows;
}
function showExamBuilder() {
  document.getElementById("exam-list-view").style.display = "none";
  document.getElementById("exam-builder-view").style.display = "block";
}
function hideExamBuilder() {
  document.getElementById("exam-builder-view").style.display = "none";
  document.getElementById("exam-list-view").style.display = "block";
}

function openNewExamBuilder() {
  var titleEl = document.getElementById("eb-title");
  var codeEl = document.getElementById("eb-code");
  var infoEl = document.getElementById("eb-info");
  var ptsEl = document.getElementById("eb-pts");
  var dateEl = document.getElementById("eb-date");
  var timeEl = document.getElementById("eb-time-start");

  if (titleEl) titleEl.value = "";
  if (codeEl) {
    codeEl.value = "";
    codeEl.removeAttribute("data-editing");
  }
  if (infoEl) infoEl.value = "";
  if (ptsEl) ptsEl.value = 70;
  if (dateEl) dateEl.value = "";
  // Set default time to next hour
  if (timeEl) {
    var d = new Date();
    d.setHours(d.getHours() + 1);
    timeEl.value = pad2(d.getHours()) + ":00";
  }

  if (typeof scrollDrumTo === "function") {
    scrollDrumTo("drum-hours", 2);
    scrollDrumTo("drum-minutes", 30);
  }

  S.parsedExam = null;
  S.parsedTasks = [];
  S.answerKey = {};
  S.audioURL = null;

  var titleBar = document.querySelector("#exam-builder-view .asec-title");
  if (titleBar)
    titleBar.innerHTML =
      '<button class="back-btn" onclick="hideExamBuilder()">&larr; Back</button> Create Exam';

  var summEl = document.getElementById("parsed-summary");
  var statusEl = document.getElementById("pdf-status");
  var resultEl = document.getElementById("pdf-parsed-result");
  var previewEl = document.getElementById("pdf-preview");
  var nameEl = document.getElementById("pdf-name");
  var audioPreviewEl = document.getElementById("audio-preview");

  if (summEl) summEl.textContent = "";
  if (statusEl) statusEl.textContent = "Ready";
  if (resultEl) resultEl.style.display = "none";
  if (previewEl) previewEl.style.display = "none";
  if (nameEl) nameEl.textContent = "exam.pdf";
  if (audioPreviewEl) audioPreviewEl.style.display = "none";

  var unlockPreview = document.getElementById("exam-unlock-preview");
  if (unlockPreview) unlockPreview.style.display = "none";
  var idsPreview = document.getElementById("exam-ids-preview");
  if (idsPreview) idsPreview.style.display = "none";

  showExamBuilder();
}

function switchBuilder(tab) {
  var bu = document.getElementById("btab-upload"),
    bk = document.getElementById("btab-key");
  if (bu) bu.className = "btab" + (tab === "upload" ? " active" : "");
  if (bk) bk.className = "btab" + (tab === "key" ? " active" : "");
  var du = document.getElementById("builder-upload"),
    dk = document.getElementById("builder-key");
  if (du) du.style.display = tab === "upload" ? "block" : "none";
  if (dk) dk.style.display = tab === "key" ? "block" : "none";
  if (tab === "key") buildAK();
}

function handlePdfUpload(e) {
  var file = e.target.files[0];
  if (!file) return;
  S.pdfFile = file;
  document.getElementById("pdf-name").textContent = file.name;
  document.getElementById("pdf-preview").style.display = "block";
  document.getElementById("pdf-parsing").style.display = "block";
  document.getElementById("pdf-parsed-result").style.display = "none";
  document.getElementById("pdf-status").textContent = "Parsing...";
  var prog = document.getElementById("parse-progress"),
    pct = 0;
  var progInt = setInterval(function () {
    pct += Math.random() * 8 + 2;
    if (pct > 90) pct = 90;
    if (prog) prog.style.width = pct + "%";
  }, 300);
  var reader = new FileReader();
  reader.onload = function (ev) {
    parsePdfWithAI(ev.target.result.split(",")[1], file.name, progInt, prog);
  };
  reader.readAsDataURL(file);
}

function parsePdfWithAI(b64, fname, progInt, progEl) {
  var prompt =
    'You are parsing an English language exam PDF. The PDF may contain Georgian (ქართული) and/or English text — preserve ALL text EXACTLY as written, do not translate or transliterate anything. Return ONLY a valid JSON object with no markdown, no code blocks, no extra text. Use this exact format: {"hasIntro":true,"introText":"Copy the FULL intro/instructions page text exactly as it appears. If no intro page exists set hasIntro to false and introText to empty string.","tasks":[{"task":1,"type":"listening","points":8,"question_count":8,"instructions":"Copy the EXACT instruction text for this task from the PDF","questions":["full question text"],"options":[["A. option text","B. option text","C. option text","D. option text"]]},{"task":2,"type":"match","points":8,"question_count":8,"instructions":"Copy EXACT instruction text","questions":["full statement text"],"passages":{"A":"Copy the COMPLETE paragraph A text exactly","B":"Complete paragraph B","C":"Complete paragraph C","D":"Complete paragraph D","E":"Complete paragraph E","F":"Complete paragraph F"}},{"task":3,"type":"reading","points":8,"question_count":8,"instructions":"Copy EXACT instruction text","text":"REQUIRED - Copy the COMPLETE reading passage WORD FOR WORD - this field is mandatory and must not be empty","questions":["full question text"],"options":[["A. opt","B. opt","C. opt","D. opt"]]},{"task":4,"type":"gapfill4","points":12,"question_count":12,"instructions":"Copy EXACT instruction text","text":"REQUIRED - Copy the COMPLETE article text WORD FOR WORD with gap markers as ......(1) ......(2) etc in EXACT positions - this field is mandatory","word_bank":{"A":"word","B":"word"}},{"task":5,"type":"gapfill5","points":12,"question_count":12,"instructions":"Copy EXACT instruction text","text":"Copy the COMPLETE article text WORD FOR WORD with gap markers as ......(1) ......(2) etc","choices":[["A. word","B. word","C. word","D. word"]]},{"task":6,"type":"dialogue","points":6,"question_count":6,"instructions":"Copy EXACT instruction text","dialogue":["Speaker: full line","Speaker: ...(1)"],"options":{"A":"sentence A","B":"sentence B"}},{"task":7,"type":"essay","points":16,"instructions":"Copy EXACT instruction text","prompt":"Copy the EXACT essay question from the PDF"}]}. RULES: (1) Extract EVERY task. (2) Copy ALL text verbatim. (3) Georgian text must be copied character by character. (4) Gap markers must appear at exact positions. (5) Detect task type from content. (6) Adjust tasks array for actual number of tasks in exam.';

  fetch("/api/parse", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ pdfBase64: b64, prompt: prompt }),
  })
    .then(function (r) {
      return r.json();
    })
    .then(function (data) {
      clearInterval(progInt);
      if (progEl) progEl.style.width = "100%";
      setTimeout(function () {
        document.getElementById("pdf-parsing").style.display = "none";
        try {
          var text = data.text || "";
          if (!text && data.error) throw new Error(data.error);
          if (!text) throw new Error("Empty response from AI");
          // Strip markdown code fences
          var bt = String.fromCharCode(96);
          var btRx = new RegExp(
            bt + bt + bt + "(?:json)?\\s*|" + bt + bt + bt,
            "g",
          );
          text = text.replace(btRx, "").trim();
          // Find outermost JSON object
          var jsonStart = text.indexOf("{");
          var jsonEnd = text.lastIndexOf("}");
          if (jsonStart === -1 || jsonEnd === -1)
            throw new Error("No JSON found in response");
          text = text.slice(jsonStart, jsonEnd + 1);
          var parsed = JSON.parse(text);
          if (!parsed.tasks || !parsed.tasks.length)
            throw new Error("No tasks found");
          // Log each task structure to debug
          for (var dbi = 0; dbi < parsed.tasks.length; dbi++) {
            var dbt = parsed.tasks[dbi];
          }
          applyParsedExam(parsed, fname);
        } catch (err) {
          console.error("Parse error:", err);
          document.getElementById("pdf-status").textContent = "Parse error";
          document.getElementById("pdf-parsed-result").style.display = "block";
          document.getElementById("parsed-summary").textContent =
            "Could not auto-read PDF. Try again or set answer key manually.";
        }
      }, 400);
    })
    .catch(function (err) {
      clearInterval(progInt);
      document.getElementById("pdf-parsing").style.display = "none";
      showToast("PDF parse failed — check your connection", "");
      console.error("Fetch error:", err);
    });
}

function applyParsedExam(parsed, fname) {
  S.parsedExam = parsed;
  S.parsedTasks = parsed.tasks || [];
  // Clear TASKS and rebuild from parsed data
  var oldTaskKeys = Object.keys(TASKS);
  for (var oki = 0; oki < oldTaskKeys.length; oki++)
    delete TASKS[oldTaskKeys[oki]];

  // Rebuild TASKS from parsed data so renderTask uses real content
  if (parsed.tasks) {
    for (var ti = 0; ti < parsed.tasks.length; ti++) {
      var t = parsed.tasks[ti];
      var tn = t.task;
      if (!TASKS[tn]) TASKS[tn] = { title: "Task " + tn, pts: t.points || 0 };
      TASKS[tn].type = t.type;
      TASKS[tn].pts = t.points || TASKS[tn].pts || 0;
      TASKS[tn].title = "Task " + tn;
      TASKS[tn].qcount = t.question_count || 0;

      // Save instructions for all task types
      TASKS[tn].instructions = t.instructions || "";
      if (t.type === "listening" || t.type === "reading") {
        TASKS[tn].qs = [];
        TASKS[tn].opts = t.answer_options || ["A", "B", "C", "D"];
        // Handle various field names Gemini might use for the passage
        var taskText =
          t.text || t.passage || t.reading_text || t.article || t.content || "";
        if (taskText) TASKS[tn].text = taskText;
        for (var qi = 0; qi < (t.questions || []).length; qi++) {
          TASKS[tn].qs.push({
            q: t.questions[qi],
            os: (t.options && t.options[qi]) || [],
          });
        }
      } else if (t.type === "match") {
        TASKS[tn].passages = t.passages || {};
        TASKS[tn].opts = t.answer_options || ["A", "B", "C", "D", "E", "F"];
        TASKS[tn].qs = (t.questions || []).map(function (q) {
          return { q: q };
        });
        // Some PDFs include text/intro for match tasks
        if (t.text) TASKS[tn].text = t.text;
      } else if (t.type === "gapfill4") {
        TASKS[tn].wordBank = t.word_bank || {};
        TASKS[tn].opts = t.answer_options || [
          "A",
          "B",
          "C",
          "D",
          "E",
          "F",
          "G",
          "H",
          "I",
          "J",
          "K",
          "L",
          "M",
          "N",
        ];
        if (t.text) TASKS[tn].text = t.text;
        TASKS[tn].qcount = t.question_count || 12;
      } else if (t.type === "gapfill5") {
        TASKS[tn].wordChoices = t.choices || [];
        TASKS[tn].opts = t.answer_options || ["A", "B", "C", "D"];
        if (t.text) TASKS[tn].text = t.text;
        TASKS[tn].qcount = t.question_count || 12;
      } else if (t.type === "dialogue") {
        TASKS[tn].dialogue = t.dialogue || [];
        TASKS[tn].dialogueSents =
          t.options || t.sentence_options || t.sentences || {};
        TASKS[tn].opts = t.answer_options || [
          "A",
          "B",
          "C",
          "D",
          "E",
          "F",
          "G",
          "H",
        ];
      } else if (t.type === "essay" || t.type === "writing") {
        TASKS[tn].prompt = t.prompt || "";
      }
    }
  }

  // Update task pills to match parsed task count
  var taskCount = (parsed.tasks || []).length;
  var pillsEl = document.getElementById("task-pills");
  if (pillsEl) {
    var h =
      '<div class="tpill active" id="pill-0" onclick="gotoTask(0)">იმი</div>';
    for (var pi = 1; pi <= taskCount; pi++) {
      h +=
        '<div class="tpill" id="pill-' +
        pi +
        '" onclick="gotoTask(' +
        pi +
        ')">' +
        pi +
        "</div>";
    }
    pillsEl.innerHTML = h;
  }

  var parts = [];
  for (var i = 0; i < (parsed.tasks || []).length; i++) {
    parts.push(
      "Task " +
        parsed.tasks[i].task +
        " (" +
        (parsed.tasks[i].question_count || 0) +
        "q)",
    );
  }
  document.getElementById("parsed-summary").textContent =
    "Detected: " + parts.join(", ") + (parsed.hasIntro ? " + Intro" : "");
  document.getElementById("pdf-status").textContent = "Parsed successfully";
  document.getElementById("pdf-parsed-result").style.display = "block";
  document.getElementById("pdf-name").textContent = fname;
  showToast("Exam questions extracted!", "success");
  if (
    document.getElementById("builder-key") &&
    document.getElementById("builder-key").style.display !== "none"
  )
    buildAK();
}
function removePdf() {
  S.pdfFile = null;
  S.pdfDoc = null;
  S.parsedExam = null;
  var pi = document.getElementById("pdf-preview");
  if (pi) pi.style.display = "none";
  var pr = document.getElementById("pdf-parsed-result");
  if (pr) pr.style.display = "none";
  var fi = document.getElementById("pdf-upload");
  if (fi) fi.value = "";
}

function removePdf() {
  document.getElementById("pdf-preview").style.display = "none";
  document.getElementById("pdf-upload").value = "";
}
async function saveExam() {
  var title = document.getElementById("eb-title")
    ? document.getElementById("eb-title").value.trim()
    : "";
  var code = document.getElementById("eb-code")
    ? document.getElementById("eb-code").value.trim().toUpperCase()
    : "";
  var infoText = document.getElementById("eb-info")
    ? document.getElementById("eb-info").value.trim()
    : "";
  var durEl = document.getElementById("eb-time-min");
  var dateEl = document.getElementById("eb-date");
  var timeEl = document.getElementById("eb-time-start");
  var duration = durEl ? parseInt(durEl.value) || 150 : 150;
  if (!title || !code) {
    showToast("Please fill in exam title and code", "");
    return;
  }
  var upperCode = code.toUpperCase();
  var pts = document.getElementById("eb-pts")
    ? parseInt(document.getElementById("eb-pts").value) || 70
    : 70;

  var codeElHtml = document.getElementById("eb-code");
  var editingCode = codeElHtml ? codeElHtml.getAttribute("data-editing") : null;
  if (editingCode !== upperCode && S.exams && S.exams[upperCode]) {
    showToast(
      "Exam code " + upperCode + " already exists! Please use a unique code.",
      "",
    );
    return;
  }

  // Save exam — strip parsed_exam to reduce payload size if too large
  var parsedExamToSave = S.parsedExam || null;
  var examRow = {
    code: upperCode,
    title: title,
    duration: duration,
    info: infoText,
    has_audio: !!S.audioURL,
    tasks: S.parsedTasks || [],
    answer_key: S.answerKey || {},
    pts: pts,
    audio_url: S.audioURL || null,
  };
  // Try with parsed_exam first, fall back without if it fails
  var ok = await dbUpsert(
    "exams",
    Object.assign({}, examRow, { parsed_exam: parsedExamToSave }),
    "code",
  );
  if (!ok) {
    showToast(
      "Warning: Could not save full exam data. Trying without PDF content...",
      "",
    );
    ok = await dbUpsert(
      "exams",
      Object.assign({}, examRow, { parsed_exam: null }),
      "code",
    );
    if (!ok) {
      showToast("Error saving exam to database. Please try again.", "");
      return;
    }
  }

  S.examCodes[upperCode] = true;
  S.exams[upperCode] = {
    title: title,
    code: upperCode,
    duration: duration,
    hasAudio: !!S.audioURL,
    tasks: S.parsedTasks || [],
    info: infoText,
    pts: pts,
    audioURL: S.audioURL || null,
  };

  if (window._ebGeneratedIDs && window._ebGeneratedIDs.length) {
    S.examIDs[upperCode] = window._ebGeneratedIDs.slice();
    for (var ii = 0; ii < window._ebGeneratedIDs.length; ii++) {
      await dbUpsert(
        "exam_ids",
        {
          exam_code: upperCode,
          student_id: window._ebGeneratedIDs[ii],
          used: false,
        },
        "exam_code,student_id",
      );
    }
  }

  if (dateEl && dateEl.value) {
    var t = timeEl ? timeEl.value : "10:00";
    S.schedule[upperCode] = {
      date: dateEl.value,
      time: t,
      title: title,
      duration: duration,
    };
    var schedOk = await dbUpsert(
      "schedule",
      { exam_code: upperCode, date: dateEl.value, time: t, duration: duration },
      "exam_code",
    );
    if (!schedOk) {
      showToast(
        "Warning: Exam saved but schedule may not have saved correctly.",
        "",
      );
    }
  } else {
    showToast(
      "Note: No date set — exam saved as draft without a schedule.",
      "",
    );
  }

  showToast("Exam saved!", "success");
  renderExamList();
  renderSchedTable();
  updateAdminDashboard();
  hideExamBuilder();
}
function buildAK() {
  var akContainer = document.getElementById("ak-container");
  if (!akContainer) return;
  if (!S.parsedExam || !S.parsedExam.tasks || !S.parsedExam.tasks.length) {
    akContainer.innerHTML =
      '<div style="text-align:center;padding:32px 20px;color:var(--gray-500);">' +
      '<div style="font-weight:600;margin-bottom:6px;font-size:14px;">No PDF uploaded yet</div>' +
      '<p style="font-size:12px;line-height:1.6;">Upload an exam PDF first.</p>' +
      '<button class="abtn abtn-ghost" style="margin-top:14px;font-size:12px;" onclick=\"switchBuilder(\x27upload\x27)\">Go to Upload</button>' +
      "</div>";
    return;
  }
  var optsByTask = {
    1: ["A", "B", "C", "D"],
    2: ["A", "B", "C", "D", "E", "F"],
    3: ["A", "B", "C", "D"],
    4: ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N"],
    5: ["A", "B", "C", "D"],
    6: ["A", "B", "C", "D", "E", "F", "G", "H"],
  };
  var html = "";
  for (var ti = 0; ti < S.parsedExam.tasks.length; ti++) {
    var t = S.parsedExam.tasks[ti],
      tn = t.task;
    if (tn < 1 || tn > 6) continue;
    var qcount = t.questions ? t.questions.length : t.question_count || 0;
    if (!qcount) continue;
    var opts = t.answer_options || optsByTask[tn] || ["A", "B", "C", "D"];
    html +=
      '<div class="ak-section"><div class="ak-label">Task ' +
      tn +
      " (" +
      qcount +
      " questions, " +
      (t.points || 0) +
      " pts)</div><div>";
    for (var q = 1; q <= qcount; q++) {
      var key = "ak-ak" + tn + "-" + q;
      html +=
        '<div class="ak-row"><span class="ql">Q' +
        q +
        '</span><div class="ak-opts">';
      for (var oi = 0; oi < opts.length; oi++) {
        var sel =
          S.answerKey && S.answerKey[key] === opts[oi] ? " ak-selected" : "";
        html +=
          '<button type="button" class="ak-opt-btn' +
          sel +
          '" data-key="' +
          key +
          '" data-val="' +
          opts[oi] +
          '" onclick="selectAK(this)">' +
          opts[oi] +
          "</button>";
      }
      html += "</div></div>";
    }
    html += "</div></div>";
  }
  akContainer.innerHTML =
    html ||
    '<p style="font-size:13px;color:var(--gray-500);">No tasks found.</p>';
}

function selectAK(btn) {
  var parent = btn.closest(".ak-opts");
  if (!parent) return;
  var btns = parent.querySelectorAll(".ak-opt-btn");
  for (var i = 0; i < btns.length; i++) btns[i].classList.remove("ak-selected");
  btn.classList.add("ak-selected");
  if (!S.answerKey) S.answerKey = {};
  S.answerKey[btn.getAttribute("data-key")] = btn.getAttribute("data-val");
}

function saveAnswerKey() {
  var saved = document.getElementById("ak-saved");
  if (saved) {
    saved.style.display = "inline";
    setTimeout(function () {
      saved.style.display = "none";
    }, 2000);
  }
  showToast("Answer key saved!", "success");
}

function applyTimePreset(val) {
  var inp = document.getElementById("eb-time-min");
  if (val !== "custom" && inp) inp.value = val;
}

function buildEssays() {
  var el = document.getElementById("essay-list");
  if (!el) return;
  el.innerHTML =
    '<div style="text-align:center;padding:20px;color:var(--gray-400);font-size:13px;">Loading...</div>';
  // Load all submissions fresh from Supabase
  dbGet("submissions").then(function (subs) {
    dbGet("grades").then(function (gradeRows) {
      // Rebuild grades map
      var gradeMap = {};
      gradeRows.forEach(function (r) {
        var key = r.student_email + "|" + r.exam_code;
        if (!gradeMap[key]) gradeMap[key] = {};
        var obj = {
          score: r.score,
          max: r.max_score,
          feedback: r.feedback || "",
        };
        if (r.fluency != null) {
          obj.fluency = r.fluency;
          obj.accuracy = r.accuracy;
        }
        gradeMap[key][r.task_key] = obj;
        // Also store in S.grades
        if (!S.grades[key]) S.grades[key] = {};
        S.grades[key][r.task_key] = obj;
      });

      var entries = [];
      subs.forEach(function (sub) {
        // Show all submissions that have answers
        var stEmail = sub.student_email,
          exCode = sub.exam_code;
        var ex = S.exams[exCode];
        if (!ex) return;
        var stName = "";
        var rks = Object.keys(S.regCodes);
        for (var ri = 0; ri < rks.length; ri++) {
          if (S.regCodes[rks[ri]].email === stEmail) {
            stName = S.regCodes[rks[ri]].name;
            break;
          }
        }
        var tasks = ex.tasks || [];
        var foundWriting = false;
        for (var ti = 0; ti < tasks.length; ti++) {
          if (tasks[ti].type === "essay" || tasks[ti].type === "writing") {
            foundWriting = true;
            var taskNum = ti + 1;
            var essayText = (sub.answers && sub.answers["t" + taskNum]) || "";
            var maxPts = tasks[ti].pts || 16;
            var half = Math.floor(maxPts / 2);
            var gk = stEmail + "|" + exCode;
            var eg =
              gradeMap[gk] &&
              (gradeMap[gk]["t" + taskNum] || gradeMap[gk]["wt" + taskNum]);
            entries.push({
              stEmail: stEmail,
              stName: stName || stEmail,
              exCode: exCode,
              exTitle: ex.title,
              taskNum: taskNum,
              essayText: essayText,
              existing: eg,
              maxPts: maxPts,
              half: half,
              gk: gk,
              legacy: false,
            });
          }
        }
        if (!foundWriting) {
          // Fallback: check t7 answer
          var essayText2 = (sub.answers && sub.answers["t7"]) || "";
          if (essayText2) {
            var gk2 = stEmail + "|" + exCode;
            var eg2 =
              gradeMap[gk2] && (gradeMap[gk2]["t7"] || gradeMap[gk2]["essay"]);
            entries.push({
              stEmail: stEmail,
              stName: stName || stEmail,
              exCode: exCode,
              exTitle: ex.title,
              taskNum: 7,
              essayText: essayText2,
              existing: eg2,
              maxPts: 16,
              half: 8,
              gk: gk2,
              legacy: false,
            });
          }
        }
      });

      if (!entries.length) {
        el.innerHTML =
          '<div class="acard" style="text-align:center;padding:28px;color:var(--gray-400);font-size:13px;">No writing tasks submitted yet.</div>';
        return;
      }
      window._essayEntries = entries;
      var h = "";
      for (var i = 0; i < entries.length; i++) {
        var e = entries[i];
        var flV = e.existing ? e.existing.fluency : "",
          acV = e.existing ? e.existing.accuracy : "",
          fbV = e.existing ? e.existing.feedback : "";
        h += '<div class="acard">';
        h +=
          '<div style="display:flex;justify-content:space-between;align-items:flex-start;flex-wrap:wrap;gap:6px;margin-bottom:10px;">';
        h +=
          '<div><div style="font-weight:700;font-size:14px;">' +
          e.stName +
          "</div>";
        h +=
          '<div style="font-size:11px;color:var(--gray-500);">' +
          e.exTitle +
          " &mdash; Task " +
          e.taskNum +
          "</div></div>";
        if (e.existing)
          h +=
            '<span style="background:#e8f5e9;color:#2e7d32;padding:3px 10px;border-radius:20px;font-size:11px;font-weight:700;">Graded: ' +
            e.existing.score +
            "/" +
            e.maxPts +
            "</span>";
        h += "</div>";
        h +=
          '<div style="background:var(--gray-50);border-radius:8px;padding:10px 12px;font-size:13px;line-height:1.7;color:var(--charcoal);margin-bottom:12px;min-height:60px;white-space:pre-wrap;">' +
          (e.essayText ||
            '<span style="color:var(--gray-400);">No response submitted</span>') +
          "</div>";
        h += '<div class="score-row">';
        h +=
          '<div class="score-wrap"><label>Fluency</label><input type="number" min="0" max="' +
          e.half +
          '" id="fl-' +
          i +
          '" value="' +
          flV +
          '" placeholder="0"><span class="mx">/ ' +
          e.half +
          "</span></div>";
        h +=
          '<div class="score-wrap"><label>Accuracy</label><input type="number" min="0" max="' +
          e.half +
          '" id="ac-' +
          i +
          '" value="' +
          acV +
          '" placeholder="0"><span class="mx">/ ' +
          e.half +
          "</span></div>";
        h += "</div>";
        h +=
          '<div class="afl"><label>Feedback (optional)</label><textarea id="fb-' +
          i +
          '" rows="2" style="width:100%;padding:7px 10px;border:1.5px solid var(--gray-200);border-radius:8px;font-family:Georgia,serif;font-size:12px;outline:none;resize:vertical;">' +
          fbV +
          "</textarea></div>";
        h +=
          '<button class="abtn abtn-pink" onclick="saveEssayGradeEntry(' +
          i +
          ')">Save Grade</button>';
        h += "</div>";
      }
      el.innerHTML = h;
    });
  });
}
function saveEssayGradeEntry(i) {
  var entries = window._essayEntries || [];
  var e = entries[i];
  if (!e) return;
  var f = document.getElementById("fl-" + i),
    a = document.getElementById("ac-" + i),
    fb = document.getElementById("fb-" + i);
  if (!f || !a || !f.value || !a.value) {
    showToast("Enter both scores", "");
    return;
  }
  var fluency = parseInt(f.value),
    accuracy = parseInt(a.value);
  if (fluency > e.half || accuracy > e.half) {
    showToast("Max is " + e.half + " per criterion", "");
    return;
  }
  var total = fluency + accuracy;
  if (!S.grades[e.gk]) S.grades[e.gk] = {};
  var obj = {
    score: total,
    max: e.maxPts,
    fluency: fluency,
    accuracy: accuracy,
    feedback: fb ? fb.value.trim() : "",
  };
  if (e.legacy) {
    S.grades[e.gk]["t" + e.taskNum] = obj;
    S.grades[e.gk].essay = obj;
  } else {
    S.grades[e.gk]["wt" + e.taskNum] = obj;
    S.grades[e.gk]["t" + e.taskNum] = obj;
  }
  // Save to Supabase — always use t+taskNum as key for consistency
  var _eEmail = e.gk.split("|")[0];
  var _eCode = e.gk.split("|")[1];
  var _tKey = "t" + e.taskNum;
  dbUpsert(
    "grades",
    {
      student_email: _eEmail,
      exam_code: _eCode,
      task_key: _tKey,
      score: total,
      max_score: e.maxPts,
      fluency: fluency,
      accuracy: accuracy,
      feedback: fb ? fb.value.trim() : "",
    },
    "student_email,exam_code,task_key",
  );
  showToast(
    "Grade saved for " + e.stName + "! " + total + "/" + e.maxPts,
    "success",
  );
  if (S.student) renderStudentResults();
  buildEssays();
}
function saveGrade(i, name) {
  var f = document.getElementById("fl-" + i);
  var a = document.getElementById("ac-" + i);
  var fb = document.getElementById("fb-" + i);
  if (!f || !a || !f.value || !a.value) {
    showToast("Enter both scores", "");
    return;
  }
  var fluency = parseInt(f.value),
    accuracy = parseInt(a.value);
  if (fluency > 8 || accuracy > 8) {
    showToast("Max is 8 per criterion", "");
    return;
  }
  var total = fluency + accuracy;
  // Find student by name and store grade
  var keys = Object.keys(S.regCodes);
  for (var i2 = 0; i2 < keys.length; i2++) {
    if (S.regCodes[keys[i2]].name === name) {
      var email = S.regCodes[keys[i2]].email;
      var gradeKey = email + "|ESSAY";
      if (!S.grades[gradeKey]) S.grades[gradeKey] = {};
      S.grades[gradeKey].essay = {
        score: total,
        max: 16,
        fluency: fluency,
        accuracy: accuracy,
        feedback: fb ? fb.value.trim() : "",
      };
      break;
    }
  }
  showToast("Essay grade saved for " + name + "! " + total + "/16", "success");
  // Refresh student results if student is logged in
  if (S.student) renderStudentResults();
}

async function saveTaskGrade(
  studentEmail,
  examCode,
  taskNum,
  score,
  max,
  feedback,
) {
  var key = studentEmail + "|" + examCode;
  if (!S.grades[key]) S.grades[key] = {};
  S.grades[key]["t" + taskNum] = {
    score: score,
    max: max,
    feedback: feedback || "",
  };
  await dbUpsert(
    "grades",
    {
      student_email: studentEmail,
      exam_code: examCode,
      task_key: "t" + taskNum,
      score: score,
      max_score: max,
      feedback: feedback || "",
    },
    "student_email,exam_code,task_key",
  );
  showToast("Task " + taskNum + " grade saved!", "success");
  if (S.student) renderStudentResults();
}
async function genCodes() {
  var prefix = document.getElementById("cp").value.trim() || "PROMOTER",
    count = parseInt(document.getElementById("cc").value) || 5;
  var tbody = document.getElementById("codes-tbody");
  var ph = document.getElementById("codes-placeholder");
  if (ph) ph.remove();
  for (var i = 0; i < count; i++) {
    var code =
      prefix + "-" + Math.random().toString(36).substr(2, 5).toUpperCase();
    S.regCodes[code] = {
      name: "",
      email: "",
      phone: "",
      pw: "",
      activated: false,
    };
    await dbInsert("students", {
      reg_code: code,
      name: "",
      email: code.toLowerCase() + "@pending.ep",
      phone: "",
      password: "",
      activated: false,
      student_id: "",
    });
    var tr = document.createElement("tr");
    tr.innerHTML =
      '<td><span class="code-tag">' +
      code +
      '</span></td><td><span class="bdg bdg-orange">Unused</span></td><td>&mdash;</td><td><button class="abtn abtn-ghost" style="padding:3px 8px;font-size:11px;" onclick="revokeCode(this,\'+code+\')">Revoke</button></td>';
    tbody.appendChild(tr);
  }
  showToast("Generated " + count + " codes!", "success");
}
async function revokeCode(btn, code) {
  if (confirm("Revoke this code?")) {
    btn.closest("tr").remove();
    if (code) {
      delete S.regCodes[code];
      await dbDelete("students", { reg_code: code });
    }
    showToast("Code revoked", "");
  }
}
function openAddStudent() {
  document.getElementById("add-stud-modal").classList.add("open");
  document.getElementById("as-err").style.display = "none";
}
function addStudent() {
  var name = document.getElementById("as-name").value.trim(),
    code = document.getElementById("as-code").value.trim();
  var contact = document.getElementById("as-contact").value.trim(),
    id = document.getElementById("as-id").value.trim();
  if (!name || !code || !id) {
    document.getElementById("as-err").style.display = "block";
    return;
  }
  S.regCodes[code] = {
    name: name,
    email: contact,
    phone: "",
    pw: "",
    activated: false,
  };
  var tbody = document.getElementById("stud-tbody"),
    tr = document.createElement("tr");
  tr.innerHTML =
    "<td><strong>" +
    name +
    '</strong></td><td><span class="code-tag">' +
    code +
    "</span></td><td>" +
    contact +
    '</td><td><span class="code-tag">' +
    id +
    '</span></td><td><span class="bdg bdg-orange">Pending</span></td><td><button class="abtn abtn-ghost" style="padding:3px 8px;font-size:11px;">Edit</button></td>';
  tbody.appendChild(tr);
  closeOverlay("add-stud-modal");
  showToast(name + " added!", "success");
  ["as-name", "as-code", "as-contact", "as-id"].forEach(function (id) {
    document.getElementById(id).value = "";
  });
}
async function postAnn(pin) {
  var t = document.getElementById("ann-t").value.trim(),
    b = document.getElementById("ann-b").value.trim();
  if (!t || !b) {
    showToast("Fill in title and message", "");
    return;
  }
  document.getElementById("ann-title").textContent = t;
  document.getElementById("ann-body").textContent = b;
  if (pin) {
    document.getElementById("pin-title").textContent = t;
    document.getElementById("pin-body").textContent = b;
    document.getElementById("pinned-ann").classList.add("show");
  }
  var hist = document.getElementById("ann-hist"),
    d = document.createElement("div");
  d.style.cssText = "padding:8px 0;border-bottom:1px solid var(--gray-100);";
  d.innerHTML =
    "<strong>" +
    t +
    "</strong>" +
    (pin
      ? '<span class="bdg bdg-pink" style="margin-left:6px;">Pinned</span>'
      : "") +
    '<p style="color:var(--gray-600);margin-top:2px;">' +
    b +
    "</p>";
  hist.insertBefore(d, hist.firstChild);
  await dbInsert("announcements", {
    title: t,
    body: b,
    pinned: pin ? true : false,
  });
  showToast(pin ? "Posted & pinned!" : "Posted!", "success");
  document.getElementById("ann-t").value = "";
  document.getElementById("ann-b").value = "";
}
function unpin() {
  document.getElementById("pinned-ann").classList.remove("show");
  showToast("Unpinned", "");
}

function startLiveCountdown() {
  var sec = 107 * 60 + 33;
  clearInterval(S.liveInt);
  S.liveInt = setInterval(function () {
    sec = Math.max(0, sec - 1);
    var h = Math.floor(sec / 3600),
      m = Math.floor((sec % 3600) / 60),
      s = sec % 60;
    var el = document.getElementById("live-cd");
    if (el)
      el.textContent = "Time remaining: " + h + ":" + pad(m) + ":" + pad(s);
    var ac = document.getElementById("admin-cd");
    if (ac) ac.textContent = h + ":" + pad(m) + ":" + pad(s);
  }, 1000);
}

// Safe init - wait for DOM

function getTodayStr() {
  var d = new Date();
  return (
    d.getFullYear() +
    "-" +
    (d.getMonth() < 9 ? "0" : "") +
    (d.getMonth() + 1) +
    "-" +
    (d.getDate() < 10 ? "0" : "") +
    d.getDate()
  );
}

function getDateStr(daysAhead) {
  var d = new Date();
  d.setDate(d.getDate() + (daysAhead || 0));
  return (
    d.getFullYear() +
    "-" +
    (d.getMonth() < 9 ? "0" : "") +
    (d.getMonth() + 1) +
    "-" +
    (d.getDate() < 10 ? "0" : "") +
    d.getDate()
  );
}

function pad2(n) {
  return n < 10 ? "0" + n : "" + n;
}

function getExamStatus(dateStr, timeStr, duration) {
  if (!dateStr) return "unscheduled";
  var now = new Date();
  var start = new Date(dateStr + "T" + (timeStr || "00:00") + ":00");
  var end = new Date(start.getTime() + (duration || 150) * 60000);
  if (now < start) return "upcoming";
  if (now >= start && now <= end) return "live";
  return "ended";
}

function formatSchedDate(dateStr, timeStr) {
  if (!dateStr) return "";
  var months = [
    "Jan",
    "Feb",
    "Mar",
    "Apr",
    "May",
    "Jun",
    "Jul",
    "Aug",
    "Sep",
    "Oct",
    "Nov",
    "Dec",
  ];
  var d = new Date(dateStr + "T" + (timeStr || "00:00") + ":00");
  return (
    months[d.getMonth()] +
    " " +
    d.getDate() +
    ", " +
    d.getFullYear() +
    (timeStr ? " · " + timeStr : "")
  );
}

function formatCountdown(dateStr, timeStr, duration) {
  if (!dateStr) return "";
  var now = new Date();
  var start = new Date(dateStr + "T" + (timeStr || "00:00") + ":00");
  var end = new Date(start.getTime() + (duration || 150) * 60000);
  var status = getExamStatus(dateStr, timeStr, duration);
  if (status === "live") {
    var rem = Math.max(0, Math.floor((end - now) / 1000));
    var h = Math.floor(rem / 3600),
      m = Math.floor((rem % 3600) / 60),
      s = rem % 60;
    return "Live — " + h + ":" + pad(m) + ":" + pad(s) + " left";
  }
  if (status === "upcoming") {
    var diff = Math.max(0, Math.floor((start - now) / 1000));
    var days = Math.floor(diff / 86400);
    var h2 = Math.floor((diff % 86400) / 3600);
    var m2 = Math.floor((diff % 3600) / 60);
    var s2 = diff % 60;
    if (days > 0)
      return "Starts in: " + days + "d " + h2 + ":" + pad(m2) + ":" + pad(s2);
    return "Starts in: " + h2 + ":" + pad(m2) + ":" + pad(s2);
  }
  return "Ended";
}

window.addEventListener("load", function () {
  var overlays = document.querySelectorAll(".overlay");
  for (var i = 0; i < overlays.length; i++) {
    (function (o) {
      o.addEventListener("click", function (e) {
        if (e.target === o) o.classList.remove("open");
      });
    })(overlays[i]);
  }
  var lipw = document.getElementById("li-pw");
  if (lipw)
    lipw.addEventListener("keydown", function (e) {
      if (e.key === "Enter") doLogin();
    });
  var supw = document.getElementById("su-pw");
  if (supw)
    supw.addEventListener("keydown", function (e) {
      if (e.key === "Enter") doSignup();
    });
  loadAllData().then(function () {
    startRealtimeSubscriptions();
  });
});

/* ──────────────────────────────────────────
   Exam List + Edit
────────────────────────────────────────── */

function renderExamList() {
  var tbody = document.getElementById("exam-list-tbody");
  if (!tbody) return;
  var keys = Object.keys(S.exams || {});
  if (!keys.length) {
    tbody.innerHTML =
      '<tr><td colspan="5" style="text-align:center;color:var(--gray-400);padding:20px;font-size:13px;">No exams yet \u2014 create one above</td></tr>';
    return;
  }
  var rows = "";
  for (var i = 0; i < keys.length; i++) {
    var code = keys[i];
    var ex = S.exams[code];
    var sched = S.schedule && S.schedule[code];
    var status = sched
      ? getExamStatus(sched.date, sched.time, sched.duration)
      : "unscheduled";
    var dur = ex.duration || 150;
    var hrs = Math.floor(dur / 60),
      mins = dur % 60;
    var durStr = (hrs > 0 ? hrs + "h " : "") + mins + "min";
    var badgeHtml;
    if (status === "live")
      badgeHtml = '<span class="bdg bdg-green">Live</span>';
    else if (status === "upcoming")
      badgeHtml = '<span class="bdg bdg-orange">Scheduled</span>';
    else if (status === "ended")
      badgeHtml = '<span class="bdg bdg-gray">Ended</span>';
    else badgeHtml = '<span class="bdg bdg-gray">Draft</span>';
    rows +=
      "<tr>" +
      '<td style="font-weight:600;font-size:13px;">' +
      (ex.title || code) +
      "</td>" +
      '<td><span class="code-tag">' +
      code +
      "</span></td>" +
      "<td>" +
      durStr +
      "</td>" +
      "<td>" +
      badgeHtml +
      "</td>" +
      '<td><button class="abtn abtn-ghost" style="padding:3px 10px;font-size:11px;" onclick="openExamEdit(\'' +
      code +
      "')\">" +
      (status === "live" ? "Manage" : "Edit") +
      "</button></td>" +
      "</tr>";
  }
  tbody.innerHTML = rows;
}

function openExamEdit(code) {
  var ex = S.exams && S.exams[code];
  if (!ex) return;
  var sched = S.schedule && S.schedule[code];
  var status = sched
    ? getExamStatus(sched.date, sched.time, sched.duration)
    : "unscheduled";
  if (status === "live") {
    openLiveEditModal(code);
  } else if (status === "ended") {
    renewExam(code);
  } else {
    openScheduledEditModal(code);
  }
}

function renewExam(code) {
  var ex = S.exams && S.exams[code];
  if (!ex) return;

  // Restore parsed exam content from stored data
  if (ex.parsedExam && ex.parsedExam.tasks && ex.parsedExam.tasks.length) {
    S.parsedExam = ex.parsedExam;
    S.parsedTasks = ex.parsedExam.tasks || [];
    applyParsedExam(ex.parsedExam, ex.title);
  } else if (ex.tasks && ex.tasks.length) {
    // Fallback: rebuild from tasks array stored in exam
    S.parsedTasks = ex.tasks;
    for (var rti = 0; rti < ex.tasks.length; rti++) {
      var rt = ex.tasks[rti];
      var rtn = rti + 1;
      if (!TASKS[rtn]) TASKS[rtn] = {};
      TASKS[rtn].type = rt.type || TASKS[rtn].type || "";
      TASKS[rtn].pts = rt.pts || TASKS[rtn].pts || 0;
      TASKS[rtn].title = "Task " + rtn;
    }
  }

  // Restore answer key
  if (ex.answer_key) {
    S.answerKey = JSON.parse(JSON.stringify(ex.answer_key));
  }

  // Pre-fill the builder fields
  var titleEl = document.getElementById("eb-title");
  var codeEl = document.getElementById("eb-code");
  var infoEl = document.getElementById("eb-info");
  var ptsEl = document.getElementById("eb-pts");
  var dateEl = document.getElementById("eb-date");
  var timeEl = document.getElementById("eb-time-start");

  if (titleEl) titleEl.value = ex.title || "";
  if (codeEl) {
    codeEl.value = "";
    codeEl.removeAttribute("data-editing");
  } // clear code so admin sets new one
  if (infoEl) infoEl.value = ex.info || "";
  if (ptsEl) ptsEl.value = ex.pts || 70;
  if (dateEl) dateEl.value = "";
  if (timeEl) timeEl.value = "10:00";

  var dur = ex.duration || 150;
  var dh = Math.floor(dur / 60),
    dm = dur % 60;
  scrollDrumTo("drum-hours", dh);
  scrollDrumTo("drum-minutes", Math.floor(dm / 5) * 5);
  onDrumScroll("hours");
  onDrumScroll("minutes");

  // Show parsed result section
  var summEl = document.getElementById("parsed-summary");
  var statusEl = document.getElementById("pdf-status");
  var resultEl = document.getElementById("pdf-parsed-result");
  var previewEl = document.getElementById("pdf-preview");
  var nameEl = document.getElementById("pdf-name");
  if (ex.parsedExam && ex.parsedExam.tasks && ex.parsedExam.tasks.length) {
    var parts = [];
    for (var i2 = 0; i2 < ex.parsedExam.tasks.length; i2++)
      parts.push(
        "Task " +
          ex.parsedExam.tasks[i2].task +
          " (" +
          (ex.parsedExam.tasks[i2].question_count || 0) +
          "q)",
      );
    if (summEl)
      summEl.textContent =
        "Restored: " +
        parts.join(", ") +
        (ex.parsedExam.hasIntro ? " + Intro" : "");
    if (statusEl) statusEl.textContent = "Restored from previous exam";
    if (resultEl) resultEl.style.display = "block";
    if (previewEl) previewEl.style.display = "block";
    if (nameEl) nameEl.textContent = ex.title + " (restored)";
  } else {
    if (summEl)
      summEl.textContent =
        "No PDF content saved — please upload a PDF to extract questions.";
    if (statusEl) statusEl.textContent = "No PDF";
    if (resultEl) resultEl.style.display = "block";
  }

  // Update builder title
  var titleBar = document.querySelector("#exam-builder-view .asec-title");
  if (titleBar)
    titleBar.innerHTML =
      '<button class="back-btn" onclick="hideExamBuilder()">&larr; Back</button> Renew Exam <span style="font-size:13px;color:var(--gray-400);font-weight:400;">— enter a new exam code to publish</span>';

  showToast(
    "Exam content restored! Enter a new code and date to republish.",
    "success",
  );
  showExamBuilder();
}

function openLiveEditModal(code) {
  var ex = S.exams[code];
  if (!ex) return;
  var modal = document.getElementById("live-edit-modal");
  var body = document.getElementById("live-edit-body");
  if (!modal || !body) return;
  modal.setAttribute("data-code", code);
  // Read real pause state from Supabase
  dbGet("timer_state", { exam_code: code }).then(function (timerRows) {
    var isPaused = timerRows.length && timerRows[0].paused;
    var t0 = 0;
    if (timerRows.length) {
      var ts = timerRows[0];
      if (ts.paused) {
        t0 = ts.time_left || 0;
      } else {
        var elapsed = ts.started_at
          ? Math.floor((Date.now() - ts.started_at) / 1000)
          : 0;
        t0 = Math.max(0, (ts.time_left || 0) - elapsed);
      }
    } else {
      var sc = S.schedule && S.schedule[code];
      t0 = sc ? (sc.duration || 150) * 60 : 0;
    }
    var timerCol = t0 < 600 ? "var(--pink)" : t0 < 1800 ? "#ff9800" : "#fff";
    var th = Math.floor(t0 / 3600),
      tm = Math.floor((t0 % 3600) / 60),
      ts2 = t0 % 60;
    var tStr = th + ":" + pad(tm) + ":" + pad(ts2);
    body.innerHTML =
      '<div style="margin-bottom:14px;">' +
      '<div style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:var(--gray-400);margin-bottom:4px;">Exam</div>' +
      '<div style="font-weight:700;font-size:15px;">' +
      (ex.title || code) +
      "</div>" +
      '<div style="font-size:12px;color:var(--gray-500);margin-top:2px;">Code: <span style="font-family:monospace;">' +
      code +
      "</span></div>" +
      "</div>" +
      // ── Big live timer ──
      '<div style="background:var(--charcoal);border-radius:14px;padding:20px 16px;text-align:center;margin-bottom:14px;">' +
      '<div style="font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:1.5px;color:rgba(255,255,255,.4);margin-bottom:8px;">' +
      (isPaused ? "⏸ PAUSED" : "▶ LIVE TIMER") +
      "</div>" +
      '<div id="live-modal-timer" style="font-family:\'Courier New\',monospace;font-size:46px;font-weight:700;letter-spacing:3px;color:' +
      timerCol +
      ';line-height:1;">' +
      tStr +
      "</div>" +
      "</div>" +
      // ── Pause / Continue ──
      '<div class="live-edit-section">' +
      '<div class="live-edit-label">Timer Control</div>' +
      '<button class="abtn ' +
      (isPaused ? "abtn-pink" : "abtn-ghost") +
      '" style="width:100%;justify-content:center;padding:11px;" onclick="confirmExamAction(\'' +
      code +
      "','pause')\">" +
      (isPaused ? "▶&nbsp; Continue Timer" : "⏸&nbsp; Pause Timer") +
      "</button>" +
      "</div>" +
      // ── Adjust time ──
      '<div class="live-edit-section" style="margin-top:12px;">' +
      '<div class="live-edit-label">Adjust Time</div>' +
      '<div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap;">' +
      '<select id="live-adjust-dir" style="padding:8px 10px;border:1.5px solid var(--gray-200);border-radius:8px;font-size:13px;outline:none;background:#fff;flex:1;">' +
      '<option value="add">+ Add time</option>' +
      '<option value="sub">\u2212 Remove time</option>' +
      "</select>" +
      '<select id="live-adjust-mins" style="padding:8px 10px;border:1.5px solid var(--gray-200);border-radius:8px;font-size:13px;outline:none;background:#fff;flex:1;">' +
      '<option value="5">5 min</option>' +
      '<option value="10" selected>10 min</option>' +
      '<option value="15">15 min</option>' +
      '<option value="20">20 min</option>' +
      '<option value="30">30 min</option>' +
      "</select>" +
      '<button class="abtn abtn-dark" onclick="confirmExamAction(\'' +
      code +
      "','adjust')\">Apply</button>" +
      "</div>" +
      "</div>";
    document.getElementById("live-edit-overlay").style.display = "flex";
  }); // end dbGet.then
  // Tick the modal timer — reads from Supabase every second for real sync
  clearInterval(window._liveModalInt);
  window._liveModalInt = setInterval(function () {
    var el = document.getElementById("live-modal-timer");
    if (!el) {
      clearInterval(window._liveModalInt);
      return;
    }
    dbGet("timer_state", { exam_code: code }).then(function (rows) {
      if (!el) return;
      var t = 0;
      if (rows.length) {
        var ts = rows[0];
        if (ts.paused) {
          t = ts.time_left || 0;
        } else {
          var elapsed = ts.started_at
            ? Math.floor((Date.now() - ts.started_at) / 1000)
            : 0;
          t = Math.max(0, (ts.time_left || 0) - elapsed);
        }
      } else {
        // No timer record yet — calculate from schedule start time
        var sc2 = S.schedule && S.schedule[code];
        if (sc2 && sc2.date && sc2.time) {
          var startMs = new Date(sc2.date + "T" + sc2.time + ":00").getTime();
          var totalSecs2 = (sc2.duration || 150) * 60;
          var elapsedSecs2 = Math.floor((Date.now() - startMs) / 1000);
          t = Math.max(0, totalSecs2 - elapsedSecs2);
        } else {
          t = sc2 ? (sc2.duration || 150) * 60 : 0;
        }
      }
      var h = Math.floor(t / 3600),
        m = Math.floor((t % 3600) / 60),
        s = t % 60;
      el.textContent = h + ":" + pad(m) + ":" + pad(s);
      el.style.color = t < 600 ? "var(--pink)" : t < 1800 ? "#ff9800" : "#fff";
    });
  }, 1000);
}

function closeLiveEditModal() {
  clearInterval(window._liveModalInt);
  var el = document.getElementById("live-edit-overlay");
  if (el) el.style.display = "none";
}

function formatTimeLeft() {
  var t = S.timeLeft || 0;
  var h = Math.floor(t / 3600),
    m = Math.floor((t % 3600) / 60),
    s = t % 60;
  return h + ":" + pad(m) + ":" + pad(s);
}

function confirmExamAction(code, action) {
  var overlay = document.getElementById("exam-confirm-overlay");
  var msg = document.getElementById("exam-confirm-msg");
  if (!overlay || !msg) return;
  overlay.setAttribute("data-code", code);
  overlay.setAttribute("data-action", action);
  if (action === "pause") {
    msg.textContent = S.timerPaused
      ? "Are you sure you want to continue the exam timer?"
      : "Are you sure you want to pause the exam timer?";
  } else if (action === "adjust") {
    var dir = document.getElementById("live-adjust-dir");
    var mins = document.getElementById("live-adjust-mins");
    var dirTxt = dir ? dir.options[dir.selectedIndex].text : "";
    var minsTxt = mins ? mins.options[mins.selectedIndex].text : "";
    msg.textContent =
      "Are you sure you want to " +
      dirTxt +
      " " +
      minsTxt +
      " to the exam timer?";
    overlay.setAttribute("data-dir", dir ? dir.value : "add");
    overlay.setAttribute("data-mins", mins ? mins.value : "10");
  }
  overlay.style.display = "flex";
}

function closeConfirm() {
  var el = document.getElementById("exam-confirm-overlay");
  if (el) el.style.display = "none";
}

function doExamAction() {
  var overlay = document.getElementById("exam-confirm-overlay");
  var action = overlay.getAttribute("data-action");
  var code = overlay.getAttribute("data-code");
  closeConfirm();
  if (action === "pause") {
    // Read current timer state from Supabase first
    dbGet("timer_state", { exam_code: code }).then(function (rows) {
      var isPaused = rows.length && rows[0].paused;
      if (isPaused) {
        // Resume — calculate correct remaining time
        var ts = rows[0];
        var newLeft = ts.time_left || 0;
        dbUpsert(
          "timer_state",
          {
            exam_code: code,
            time_left: newLeft,
            paused: false,
            started_at: Date.now(),
          },
          "exam_code",
        );
        showToast("Timer resumed", "success");
      } else {
        // Pause — save current remaining time
        var curLeft = 0;
        if (rows.length) {
          var ts2 = rows[0];
          var elapsed = ts2.started_at
            ? Math.floor((Date.now() - ts2.started_at) / 1000)
            : 0;
          curLeft = Math.max(0, (ts2.time_left || 0) - elapsed);
        }
        dbUpsert(
          "timer_state",
          { exam_code: code, time_left: curLeft, paused: true, started_at: 0 },
          "exam_code",
        );
        showToast("Timer paused", "");
      }
      // Re-open modal to reflect new state
      setTimeout(function () {
        openLiveEditModal(code);
      }, 300);
    });
  } else if (action === "adjust") {
    var dir = overlay.getAttribute("data-dir") || "add";
    var mins = parseInt(overlay.getAttribute("data-mins") || "10");
    var secs = mins * 60;

    dbGet("timer_state", { exam_code: code }).then(function (rows) {
      if (!rows.length) {
        showToast("Timer not started yet", "error");
        return;
      }
      var ts = rows[0];
      var elapsed = ts.started_at
        ? Math.floor((Date.now() - ts.started_at) / 1000)
        : 0;
      var curLeft = ts.paused
        ? ts.time_left || 0
        : Math.max(0, (ts.time_left || 0) - elapsed);
      var newLeft =
        dir === "add" ? curLeft + secs : Math.max(10, curLeft - secs);

      dbUpsert(
        "timer_state",
        {
          exam_code: code,
          time_left: newLeft,
          paused: ts.paused,
          started_at: ts.paused ? 0 : Date.now(),
        },
        "exam_code",
      );

      // Sync schedule duration so getExamStatus() shows "ended" only when the adjusted timer actually finishes
      var _sc = S.schedule && S.schedule[code];
      if (_sc && _sc.date && _sc.time) {
        var _schedStartMs = new Date(
          _sc.date + "T" + _sc.time + ":00",
        ).getTime();
        var _newDurMins = Math.ceil(
          (Date.now() - _schedStartMs + newLeft * 1000) / 60000,
        );
        if (_newDurMins > 0) {
          S.schedule[code].duration = _newDurMins;
          dbUpsert(
            "schedule",
            {
              exam_code: code,
              date: _sc.date,
              time: _sc.time,
              duration: _newDurMins,
            },
            "exam_code",
          );
        }
      }

      if (dir === "add")
        showToast("+" + mins + " min added to timer", "success");
      else showToast("\u2212" + mins + " min removed from timer", "");

      openLiveEditModal(code);
    });
  }
}

function showGuaranteedSuccessModal(isTimesUp) {
  var id = "guaranteed-success-modal";
  if (document.getElementById(id)) document.getElementById(id).remove();
  var modal = document.createElement("div");
  modal.id = id;
  modal.style.cssText =
    "position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,0.8);z-index:9999999;display:flex;align-items:center;justify-content:center;padding:20px;";

  var title = isTimesUp ? "Time's Up!" : "Exam Submitted!";
  var msg = isTimesUp
    ? "Your time has run out and your answers have been recorded."
    : "Your exam has been successfully submitted. Your results will be available once grading is complete.";

  var iconHtml = "";
  if (!isTimesUp) {
    iconHtml =
      '<div style="width:64px;height:64px;background:#e8f5e9;border-radius:50%;display:flex;align-items:center;justify-content:center;margin:0 auto 16px;font-size:30px;color:#2e7d32;">&#10003;</div>';
  }

  modal.innerHTML =
    '<div style="background:#fff;border-radius:20px;width:100%;max-width:380px;padding:36px;text-align:center;box-shadow:0 10px 50px rgba(0,0,0,0.5);">' +
    iconHtml +
    '<div style="font-family:Georgia,serif;font-size:22px;font-weight:bold;margin-bottom:10px;color:#333;">' +
    title +
    '</div><div style="font-size:14px;color:#666;line-height:1.6;margin-bottom:24px;">' +
    msg +
    '</div><button id="gsm-btn" disabled onclick="document.getElementById(\'' +
    id +
    "').remove(); showPage('page-student'); renderStudentExams(); renderStudentResults();\" style=\"width:100%;padding:14px;background:#e8006e;color:#fff;border:none;border-radius:12px;font-family:Georgia,serif;font-size:15px;font-weight:bold;cursor:not-allowed;opacity:0.6;\">Saving...</button></div>";
  document.body.appendChild(modal);
}

function openScheduledEditModal(code) {
  var ex = S.exams && S.exams[code];
  var sched = S.schedule && S.schedule[code];
  if (!ex) return;
  var titleEl = document.getElementById("eb-title");
  var codeEl = document.getElementById("eb-code");
  var infoEl = document.getElementById("eb-info");
  var dateEl = document.getElementById("eb-date");
  var timeEl = document.getElementById("eb-time-start");
  if (titleEl) titleEl.value = ex.title || "";
  if (codeEl) {
    codeEl.value = code;
    codeEl.setAttribute("data-editing", code);
  }
  if (infoEl) infoEl.value = ex.info || "";
  if (sched) {
    if (dateEl) dateEl.value = sched.date || "";
    if (timeEl) timeEl.value = sched.time || "10:00";
  }
  var dur = ex.duration || 150;
  var dh = Math.floor(dur / 60),
    dm = dur % 60;
  scrollDrumTo("drum-hours", dh);
  scrollDrumTo("drum-minutes", Math.floor(dm / 5) * 5);
  onDrumScroll("hours");
  onDrumScroll("minutes");
  var titleBar = document.querySelector("#exam-builder-view .asec-title");
  if (titleBar)
    titleBar.innerHTML =
      '<button class="back-btn" onclick="hideExamBuilder()">&larr; Back</button> Edit Exam';
  showExamBuilder();
}

function scrollDrumTo(drumId, val) {
  // Accept 'drum-minutes' as alias for 'drum-mins'
  if (drumId === "drum-minutes") drumId = "drum-mins";
  var drum = document.getElementById(drumId);
  if (!drum) return;
  var ITEM_H = 44;
  var items = drum.querySelectorAll(".drum-item");
  for (var i = 0; i < items.length; i++) {
    if (parseInt(items[i].getAttribute("data-val")) === val) {
      drum.scrollTop = i * ITEM_H;
      break;
    }
  }
  // Trigger display update
  var col = drumId === "drum-hours" ? "hours" : "mins";
  onDrumScroll(col);
}
function showTimesUp() {
  try {
    clearInterval(S.timerInt);
    S.timerPaused = false;
    var el = document.getElementById("exam-timer");
    if (el) {
      el.textContent = "0:00:00";
      el.className = "exam-timer danger";
    }

    try {
      if (typeof disableExamProtection === "function") disableExamProtection();
    } catch (e) {}
    try {
      document.body.classList.remove("exam-active");
    } catch (e) {}

    doSubmitFromTimesUp();
  } catch (err) {
    console.error("showTimesUp error:", err);
  }
}

async function doSubmitFromTimesUp() {
  showGuaranteedSuccessModal(true);
  clearInterval(window._pausePollInt);
  try {
    var _anyEssay2 = document.querySelector("textarea.essay-ta");
    if (_anyEssay2) {
      var _tid2 = _anyEssay2.id.replace("essay-ta-", "");
      if (_tid2) S.answers["t" + _tid2] = _anyEssay2.value;
    }
    if (S.student && S.activeExamCode) {
      var elapsed = S.examStartTime
        ? Math.floor((Date.now() - S.examStartTime) / 1000)
        : 0;
      var eh = Math.floor(elapsed / 3600),
        em = Math.floor((elapsed % 3600) / 60);
      var elapsedStr = (eh > 0 ? eh + "h " : "") + em + "m";
      var compKey = S.student.email + "|" + S.activeExamCode;
      var answers = JSON.parse(JSON.stringify(S.answers));
      S.completed[compKey] = {
        elapsedStr: elapsedStr,
        submittedAt: new Date().toISOString(),
        answers: answers,
      };
      await dbUpsert(
        "submissions",
        {
          student_email: S.student.email,
          exam_code: S.activeExamCode,
          answers: answers,
          elapsed_str: elapsedStr,
          exam_start_time: S.examStartTime || 0,
          submitted_at: new Date().toISOString(),
        },
        "student_email,exam_code",
      );
    }
  } catch (e) {
    console.error("doSubmitFromTimesUp error:", e);
  } finally {
    var btn = document.getElementById("gsm-btn");
    if (btn) {
      btn.disabled = false;
      btn.style.opacity = "1";
      btn.style.cursor = "pointer";
      btn.textContent = "Okay, return to dashboard";
    }
  }
}

function showSubmitConfirm() {
  try {
    if (typeof disableExamProtection === "function") disableExamProtection();
  } catch (e) {}
  try {
    document.body.classList.remove("exam-active");
  } catch (e) {}
  doSubmit();
}
async function doSubmit() {
  showGuaranteedSuccessModal(false);
  clearInterval(S.timerInt);
  clearInterval(window._pausePollInt);
  S.timerPaused = false;
  try {
    // Capture essay textarea if currently on essay task
    var _essayTa = document.getElementById("essay-ta-" + S.curTask);
    if (_essayTa) S.answers["t" + S.curTask] = _essayTa.value;
    // Also capture any essay textarea on page regardless of task number
    var _anyEssay = document.querySelector("textarea.essay-ta");
    if (_anyEssay) {
      var _tid = _anyEssay.id.replace("essay-ta-", "");
      if (_tid) S.answers["t" + _tid] = _anyEssay.value;
    }
    if (S.student && S.activeExamCode) {
      var elapsed = S.examStartTime
        ? Math.floor((Date.now() - S.examStartTime) / 1000)
        : 0;
      var eh = Math.floor(elapsed / 3600),
        em = Math.floor((elapsed % 3600) / 60);
      var elapsedStr = (eh > 0 ? eh + "h " : "") + em + "m";
      var compKey = S.student.email + "|" + S.activeExamCode;
      var answers = JSON.parse(JSON.stringify(S.answers));
      S.completed[compKey] = {
        elapsedStr: elapsedStr,
        submittedAt: new Date().toISOString(),
        answers: answers,
      };
      await dbUpsert(
        "submissions",
        {
          student_email: S.student.email,
          exam_code: S.activeExamCode,
          answers: answers,
          elapsed_str: elapsedStr,
          exam_start_time: S.examStartTime || 0,
          submitted_at: new Date().toISOString(),
        },
        "student_email,exam_code",
      );
    }
  } catch (e) {
    console.error("doSubmit error:", e);
  } finally {
    var btn = document.getElementById("gsm-btn");
    if (btn) {
      btn.disabled = false;
      btn.style.opacity = "1";
      btn.style.cursor = "pointer";
      btn.textContent = "Okay, return to dashboard";
    }
  }
}
function closeSubmitSuccess() {
  document.getElementById("submit-success-overlay").style.display = "none";
  showPage("page-student");
  renderStudentResults();
}

function closeStatModal(event) {
  var overlay = document.getElementById("stat-modal-overlay");
  if (!overlay) return;
  if (!event || event.target === overlay) overlay.classList.remove("open");
}

function openStatModal(key) {
  var titleEl = document.getElementById("stat-modal-title");
  var bodyEl = document.getElementById("stat-modal-body");
  var overlay = document.getElementById("stat-modal-overlay");
  if (!titleEl || !bodyEl || !overlay) return;

  var title = "";
  var body = "";

  if (key === "students") {
    title = "Registered Students";
    var codes = Object.keys(S.regCodes || {});
    if (!codes.length) {
      body =
        '<p style="color:var(--gray-400);font-size:13px;">No students registered yet.</p>';
    } else {
      body =
        '<table style="width:100%;border-collapse:collapse;font-size:13px;">' +
        "<thead><tr>" +
        '<th style="text-align:left;padding:6px 8px;color:var(--gray-400);font-weight:600;font-size:11px;border-bottom:1px solid var(--gray-200);">Name</th>' +
        '<th style="text-align:left;padding:6px 8px;color:var(--gray-400);font-weight:600;font-size:11px;border-bottom:1px solid var(--gray-200);">Email</th>' +
        '<th style="text-align:left;padding:6px 8px;color:var(--gray-400);font-weight:600;font-size:11px;border-bottom:1px solid var(--gray-200);">Status</th>' +
        "</tr></thead><tbody>";
      for (var i = 0; i < codes.length; i++) {
        var u = S.regCodes[codes[i]];
        var status = u.activated
          ? '<span style="color:#16a34a;font-weight:600;">Active</span>'
          : '<span style="color:var(--gray-400);">Pending</span>';
        body +=
          "<tr>" +
          '<td style="padding:8px 8px;border-bottom:1px solid var(--gray-100);">' +
          (u.name || "-") +
          "</td>" +
          '<td style="padding:8px 8px;border-bottom:1px solid var(--gray-100);color:var(--gray-500);">' +
          (u.email || "-") +
          "</td>" +
          '<td style="padding:8px 8px;border-bottom:1px solid var(--gray-100);">' +
          status +
          "</td>" +
          "</tr>";
      }
      body += "</tbody></table>";
    }
    body +=
      '<button class="abtn abtn-pink" onclick="closeStatModal();showSec(&quot;students&quot;);" style="margin-top:12px;width:100%;justify-content:center;">Go to Students</button>';
  } else if (key === "scheduled") {
    title = "Scheduled Exams";
    var schedKeys = Object.keys(S.schedule || {});
    if (!schedKeys.length) {
      body =
        '<p style="color:var(--gray-400);font-size:13px;">No exams scheduled yet.</p>';
    } else {
      body =
        '<table style="width:100%;border-collapse:collapse;font-size:13px;">' +
        "<thead><tr>" +
        '<th style="text-align:left;padding:6px 8px;color:var(--gray-400);font-weight:600;font-size:11px;border-bottom:1px solid var(--gray-200);">Exam</th>' +
        '<th style="text-align:left;padding:6px 8px;color:var(--gray-400);font-weight:600;font-size:11px;border-bottom:1px solid var(--gray-200);">Date</th>' +
        '<th style="text-align:left;padding:6px 8px;color:var(--gray-400);font-weight:600;font-size:11px;border-bottom:1px solid var(--gray-200);">Status</th>' +
        "</tr></thead><tbody>";
      for (var j = 0; j < schedKeys.length; j++) {
        var sc = S.schedule[schedKeys[j]];
        var st = getExamStatus(sc.date, sc.time, sc.duration);
        var stLabel =
          st === "live"
            ? '<span style="color:var(--pink);font-weight:600;">Live</span>'
            : st === "upcoming"
              ? '<span style="color:#d97706;font-weight:600;">Upcoming</span>'
              : '<span style="color:var(--gray-400);">Ended</span>';
        var dur = sc.duration || 150;
        var dh = Math.floor(dur / 60),
          dm = dur % 60;
        var durStr = (dh > 0 ? dh + "h " : "") + dm + "min";
        body +=
          "<tr>" +
          '<td style="padding:8px 8px;border-bottom:1px solid var(--gray-100);font-weight:600;">' +
          (sc.title || schedKeys[j]) +
          "</td>" +
          '<td style="padding:8px 8px;border-bottom:1px solid var(--gray-100);color:var(--gray-500);">' +
          (sc.date || "-") +
          " " +
          (sc.time || "") +
          " &middot; " +
          durStr +
          "</td>" +
          '<td style="padding:8px 8px;border-bottom:1px solid var(--gray-100);">' +
          stLabel +
          "</td>" +
          "</tr>";
      }
      body += "</tbody></table>";
    }
  } else if (key === "essays") {
    title = "Essays to Grade";
    body =
      '<p style="color:var(--gray-400);font-size:13px;padding:10px 0;">Essay grading is available in the Grade Essays section.</p>';
    body +=
      '<button class="abtn abtn-pink" onclick="closeStatModal();showSec(&quot;grading&quot;);" style="margin-top:8px;width:100%;justify-content:center;">Go to Grade Essays</button>';
  } else if (key === "avgscore") {
    title = "Average Score";
    body =
      '<p style="color:var(--gray-400);font-size:13px;padding:10px 0;">Score data will appear here once students complete exams. Full results are available in the Results section.</p>';
    body +=
      '<button class="abtn abtn-pink" onclick="closeStatModal();showSec(&quot;results&quot;);" style="margin-top:8px;width:100%;justify-content:center;">Go to Results</button>';
  } else {
    body =
      '<p style="color:var(--gray-400);font-size:13px;">No data available.</p>';
  }

  titleEl.textContent = title;
  bodyEl.innerHTML = body;
  overlay.classList.add("open");
}

function bsGotoAction() {
  var overlay = document.getElementById("stat-modal-overlay");
  if (overlay) overlay.classList.remove("open");
}

function regeneratePdf() {
  if (!S.pdfFile) {
    showToast("No PDF uploaded", "");
    return;
  }
  document.getElementById("pdf-parsed-result").style.display = "none";
  document.getElementById("pdf-parsing").style.display = "block";
  document.getElementById("pdf-status").textContent = "Re-parsing...";
  var prog = document.getElementById("parse-progress");
  if (prog) prog.style.width = "0%";
  var pct = 0;
  var progInt = setInterval(function () {
    pct += Math.random() * 8 + 2;
    if (pct > 90) pct = 90;
    if (prog) prog.style.width = pct + "%";
  }, 300);
  var reader = new FileReader();
  reader.onload = function (ev) {
    parsePdfWithAI(
      ev.target.result.split(",")[1],
      S.pdfFile.name,
      progInt,
      prog,
    );
  };
  reader.readAsDataURL(S.pdfFile);
}
function openExamPreview() {
  var overlay = document.getElementById("preview-overlay");
  if (!overlay) return;
  buildPreviewContent();
  overlay.style.display = "flex";
}

function closeExamPreview() {
  var el = document.getElementById("preview-overlay");
  if (el) el.style.display = "none";
}

function buildPreviewContent() {
  var body = document.getElementById("preview-body");
  if (!body) return;
  if (!S.answerKey) S.answerKey = {};
  var ls = ["A", "B", "C", "D", "E", "F", "G", "H"];
  var taskKeys = Object.keys(TASKS)
    .map(Number)
    .filter(function (k) {
      return k >= 1 && TASKS[k] && TASKS[k].type;
    })
    .sort(function (a, b) {
      return a - b;
    });
  for (var dbi2 = 0; dbi2 < taskKeys.length; dbi2++) {
    var dbt2 = TASKS[taskKeys[dbi2]];
  }
  if (
    !taskKeys.length &&
    S.parsedExam &&
    S.parsedExam.tasks &&
    S.parsedExam.tasks.length
  ) {
    applyParsedExam(S.parsedExam, S.parsedExam.introText || "exam");
    taskKeys = Object.keys(TASKS)
      .map(Number)
      .filter(function (k) {
        return k >= 1 && TASKS[k] && TASKS[k].type;
      })
      .sort(function (a, b) {
        return a - b;
      });
  }

  var h = '<div style="padding:20px;">';
  h +=
    '<div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:16px;flex-wrap:wrap;gap:10px;">';
  h +=
    '<div style="font-family:Georgia,serif;font-size:22px;font-weight:700;">Exam Preview &amp; Edit</div>';
  h +=
    '<button onclick="closeExamPreview()" style="background:none;border:1.5px solid var(--gray-300);border-radius:8px;padding:6px 14px;cursor:pointer;font-size:13px;">Close</button>';
  h += "</div>";
  h +=
    '<div style="background:#fff8e1;border:1.5px solid #f59e0b;border-radius:10px;padding:10px 14px;font-size:12px;color:#92400e;margin-bottom:16px;">Select the correct answer for each question to build the answer key. This is only visible to you and used for auto-grading.</div>';

  // Intro section
  if (S.parsedExam && S.parsedExam.hasIntro && S.parsedExam.introText) {
    h +=
      '<div id="ptask-intro" style="background:#fff;border:1.5px solid var(--pink);border-radius:14px;padding:20px;margin-bottom:16px;">';
    h +=
      '<div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:12px;">';
    h +=
      '<div style="font-weight:700;font-size:15px;">Intro Page <span style="color:var(--gray-400);font-size:12px;font-weight:400;">(shown to students before exam)</span></div>';
    h += "</div>";
    h +=
      '<label style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:var(--gray-400);display:block;margin-bottom:4px;">Intro Text</label>';
    h +=
      '<textarea id="pedit-intro" rows="6" style="width:100%;padding:10px;border:1.5px solid var(--gray-200);border-radius:8px;font-size:13px;font-family:Georgia,serif;resize:vertical;">' +
      escapeAttr(S.parsedExam.introText || "") +
      "</textarea>";
    h += "</div>";
  }

  h += '<div style="display:flex;gap:8px;flex-wrap:wrap;margin-bottom:20px;">';
  if (S.parsedExam && S.parsedExam.hasIntro) {
    h +=
      '<button onclick="previewScrollTo(\'ptask-intro\')" style="padding:6px 14px;border:1.5px solid var(--pink);border-radius:20px;background:#fff;cursor:pointer;font-size:12px;font-weight:600;color:var(--pink);">Intro</button>';
  }
  for (var pi2 = 0; pi2 < taskKeys.length; pi2++) {
    var tnp = taskKeys[pi2];
    if (!TASKS[tnp]) continue;
    h +=
      "<button onclick=\"previewScrollTo('ptask-" +
      tnp +
      '\')" style="padding:6px 14px;border:1.5px solid var(--gray-200);border-radius:20px;background:#fff;cursor:pointer;font-size:12px;font-weight:600;">Task ' +
      tnp +
      "</button>";
  }
  h += "</div>";

  for (var ti = 0; ti < taskKeys.length; ti++) {
    var tn2 = taskKeys[ti];
    var tk2 = TASKS[tn2];
    if (!tk2) continue;
    var isEssay2 = tk2.type === "essay" || tk2.type === "writing";

    h +=
      '<div id="ptask-' +
      tn2 +
      '" style="background:#fff;border:1.5px solid var(--gray-200);border-radius:14px;padding:20px;margin-bottom:16px;">';
    h +=
      '<div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:14px;flex-wrap:wrap;gap:8px;">';
    h +=
      '<div style="font-weight:700;font-size:15px;">Task ' +
      tn2 +
      ' <span style="color:var(--gray-400);font-size:12px;font-weight:400;">(' +
      tk2.type +
      ")</span></div>";
    h += '<div style="display:flex;align-items:center;gap:6px;">';
    h +=
      '<input type="number" value="' +
      (tk2.pts || 0) +
      '" min="0" onchange="previewEditPts(' +
      tn2 +
      ',this.value)" style="width:60px;padding:4px 8px;border:1.5px solid var(--gray-200);border-radius:6px;font-size:13px;font-weight:700;color:var(--pink);text-align:center;outline:none;">';
    h +=
      '<span style="font-size:11px;color:var(--pink);font-weight:600;">pts</span>';
    h += "</div></div>";
    // Editable instructions
    h += '<div style="margin-bottom:14px;">';
    h +=
      '<label style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:var(--gray-400);display:block;margin-bottom:4px;">Task Instructions</label>';
    h +=
      '<input type="text" value="' +
      escapeAttr(tk2.instructions || "") +
      '" onchange="previewEditInstr(' +
      tn2 +
      ',this.value)" placeholder="e.g. Listen and mark the correct answer A, B, C or D." style="width:100%;padding:8px 10px;border:1.5px solid var(--gray-200);border-radius:6px;font-size:13px;outline:none;">';
    h += "</div>";

    // Editable article/passage text
    if (tk2.text) {
      var txtLabel2 =
        tk2.type === "reading"
          ? "Reading Passage"
          : tk2.type === "gapfill4" || tk2.type === "gapfill5"
            ? "Article Text (with gap markers ......(1) etc)"
            : tk2.type === "match"
              ? "Intro / Context Text"
              : "Article / Passage Text";
      h += '<div style="margin-bottom:14px;">';
      h +=
        '<label style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:var(--gray-400);display:block;margin-bottom:4px;">' +
        txtLabel2 +
        "</label>";
      h +=
        '<textarea id="pedit-text-' +
        tn2 +
        '" rows="6" style="width:100%;padding:10px;border:1.5px solid var(--gray-200);border-radius:8px;font-size:13px;font-family:Georgia,serif;resize:vertical;">' +
        escapeAttr(tk2.text || "") +
        "</textarea>";
      h += "</div>";
    }
    // MCQ tasks (listening, reading, match)
    if (tk2.qs && tk2.qs.length && !isEssay2) {
      h +=
        '<div style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:var(--gray-400);margin-bottom:10px;">Questions &amp; Answer Key</div>';
      for (var qi = 0; qi < tk2.qs.length; qi++) {
        var q = tk2.qs[qi];
        var akKey = "t" + tn2 + "_q" + (qi + 1);
        var akVal = S.answerKey[akKey] || "";
        h +=
          '<div style="margin-bottom:12px;padding:12px;background:var(--gray-50);border-radius:8px;">';
        h +=
          '<div style="display:flex;align-items:flex-start;gap:8px;margin-bottom:8px;">';
        h +=
          '<span style="font-size:12px;font-weight:700;color:var(--pink);min-width:24px;">' +
          (qi + 1) +
          ".</span>";
        h +=
          '<input type="text" value="' +
          escapeAttr(q.q || "") +
          '" onchange="previewEditQ(' +
          tn2 +
          "," +
          qi +
          ',\'q\',this.value)" style="flex:1;padding:6px 10px;border:1.5px solid var(--gray-200);border-radius:6px;font-size:13px;">';
        h += "</div>";
        if (q.os && q.os.length) {
          for (var oi = 0; oi < q.os.length; oi++) {
            var optL = ls[oi];
            var isCorr = akVal === optL;
            h +=
              '<div style="display:flex;align-items:center;gap:8px;margin-bottom:5px;padding:6px 10px;border-radius:6px;' +
              (isCorr
                ? "background:#e8f5e9;border:1.5px solid #4caf50;"
                : "background:#fff;border:1.5px solid var(--gray-200);") +
              '">';
            h +=
              '<input type="checkbox" ' +
              (isCorr ? "checked" : "") +
              " onchange=\"setAK('" +
              akKey +
              "','" +
              optL +
              "'," +
              tn2 +
              ',this.checked)" style="width:16px;height:16px;cursor:pointer;accent-color:#4caf50;flex-shrink:0;">';
            h +=
              '<span style="font-size:11px;font-weight:700;color:' +
              (isCorr ? "#2e7d32" : "var(--gray-500)") +
              ';width:18px;">' +
              optL +
              ".</span>";
            h +=
              '<input type="text" value="' +
              escapeAttr(q.os[oi] || "") +
              '" onchange="previewEditQ(' +
              tn2 +
              "," +
              qi +
              ",'o" +
              oi +
              '\',this.value)" style="flex:1;padding:4px 8px;border:none;background:transparent;font-size:12px;outline:none;">';
            h += "</div>";
          }
        } else {
          // Match — dropdown
          h +=
            '<div style="display:flex;align-items:center;gap:8px;margin-top:6px;">';
          h +=
            '<span style="font-size:11px;color:var(--gray-500);">Correct answer:</span>';
          h +=
            "<select onchange=\"setAKDirect('" +
            akKey +
            '\',this.value)" style="padding:4px 8px;border:1.5px solid var(--gray-200);border-radius:6px;font-size:12px;">';
          h += '<option value="">-- select --</option>';
          var mOpts = tk2.opts || ["A", "B", "C", "D", "E", "F"];
          for (var moi = 0; moi < mOpts.length; moi++)
            h +=
              '<option value="' +
              mOpts[moi] +
              '"' +
              (akVal === mOpts[moi] ? " selected" : "") +
              ">" +
              mOpts[moi] +
              "</option>";
          h += "</select></div>";
        }
        h += "</div>";
      }
    }
    // Match passages editing
    if (tk2.type === "match") {
    }
    if (
      tk2.type === "match" &&
      tk2.passages &&
      Object.keys(tk2.passages).length
    ) {
      h +=
        '<div style="margin-bottom:14px;"><label style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:var(--gray-400);display:block;margin-bottom:8px;">Paragraphs</label>';
      var pkk = Object.keys(tk2.passages);
      for (var pki = 0; pki < pkk.length; pki++) {
        h +=
          '<div style="display:flex;gap:6px;margin-bottom:8px;align-items:flex-start;">';
        h +=
          '<span style="font-size:12px;font-weight:700;color:var(--pink);min-width:20px;margin-top:8px;">' +
          pkk[pki] +
          ".</span>";
        h +=
          '<textarea rows="3" onchange="previewEditPassage(' +
          tn2 +
          ",'" +
          pkk[pki] +
          '\',this.value)" style="flex:1;padding:6px 10px;border:1.5px solid var(--gray-200);border-radius:6px;font-size:12px;font-family:Georgia,serif;resize:vertical;">' +
          escapeAttr(tk2.passages[pkk[pki]] || "") +
          "</textarea>";
        h += "</div>";
      }
      h += "</div>";
    }

    // Gapfill4: word bank + answer key dropdowns
    if (tk2.type === "gapfill4" && tk2.wordBank) {
      var wbKeys = Object.keys(tk2.wordBank);
      h +=
        '<div style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:var(--gray-400);margin-bottom:8px;">Word Bank</div>';
      h +=
        '<div style="display:flex;flex-wrap:wrap;gap:8px;margin-bottom:14px;">';
      for (var wi = 0; wi < wbKeys.length; wi++) {
        var wk = wbKeys[wi];
        h +=
          '<div style="display:flex;align-items:center;gap:4px;"><span style="font-size:11px;font-weight:700;color:var(--gray-500);">' +
          wk +
          ".</span>";
        h +=
          '<input type="text" value="' +
          escapeAttr(tk2.wordBank[wk] || "") +
          '" onchange="previewEditWordBank(' +
          tn2 +
          ",'" +
          wk +
          '\',this.value)" style="width:90px;padding:4px 8px;border:1.5px solid var(--gray-200);border-radius:6px;font-size:12px;"></div>';
      }
      h += "</div>";
      h +=
        '<div style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:var(--gray-400);margin-bottom:8px;">Answer Key (per gap)</div>';
      var gapCnt4 = tk2.qcount || wbKeys.length;
      h += '<div style="display:flex;flex-wrap:wrap;gap:8px;">';
      for (var gi4 = 1; gi4 <= gapCnt4; gi4++) {
        var akk4 = "t" + tn2 + "_q" + gi4;
        var akv4 = S.answerKey[akk4] || "";
        h +=
          '<div style="display:flex;align-items:center;gap:6px;background:var(--gray-50);border-radius:8px;padding:6px 10px;">';
        h +=
          '<span style="font-size:12px;font-weight:700;color:var(--pink);">' +
          gi4 +
          ".</span>";
        h +=
          "<select onchange=\"setAKDirect('" +
          akk4 +
          '\',this.value)" style="padding:4px 8px;border:1.5px solid var(--gray-200);border-radius:6px;font-size:12px;">';
        h += '<option value="">--</option>';
        for (var wki = 0; wki < wbKeys.length; wki++) {
          var wkl = wbKeys[wki];
          var wkv = tk2.wordBank[wkl];
          h +=
            '<option value="' +
            wkl +
            '"' +
            (akv4 === wkl ? " selected" : "") +
            ">" +
            wkl +
            ". " +
            wkv +
            "</option>";
        }
        h += "</select></div>";
      }
      h += "</div>";
    }

    // Gapfill4: show word bank entries clearly
    if (
      tk2.type === "gapfill4" &&
      tk2.wordBank &&
      !Object.keys(tk2.wordBank).length
    ) {
      h +=
        '<div style="background:#fff3e0;border-radius:8px;padding:10px;font-size:12px;color:#92400e;margin-bottom:10px;">Word bank appears empty — the article text above contains the gaps.</div>';
    }
    // Gapfill5: choices + answer key
    if (tk2.type === "gapfill5" && tk2.wordChoices && tk2.wordChoices.length) {
      h +=
        '<div style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:var(--gray-400);margin-bottom:8px;">Gap Choices &amp; Answer Key</div>';
      for (var ci5 = 0; ci5 < tk2.wordChoices.length; ci5++) {
        var ch5 = tk2.wordChoices[ci5];
        var akk5 = "t" + tn2 + "_q" + (ci5 + 1);
        var akv5 = S.answerKey[akk5] || "";
        h +=
          '<div style="display:flex;align-items:center;gap:8px;margin-bottom:8px;flex-wrap:wrap;padding:8px;background:var(--gray-50);border-radius:8px;">';
        h +=
          '<span style="font-size:12px;font-weight:700;color:var(--pink);min-width:20px;">' +
          (ci5 + 1) +
          ".</span>";
        for (var li5 = 0; li5 < ls.length && li5 < ch5.length; li5++) {
          var w5 = (ch5[li5] || "").replace(/^[A-D]\.\s*/, "");
          var isC5 = akv5 === ls[li5];
          h +=
            '<div style="display:flex;align-items:center;gap:4px;padding:4px 8px;border-radius:6px;' +
            (isC5
              ? "background:#e8f5e9;border:1.5px solid #4caf50;"
              : "background:#fff;border:1.5px solid var(--gray-200);") +
            '">';
          h +=
            '<input type="checkbox" ' +
            (isC5 ? "checked" : "") +
            " onchange=\"setAK('" +
            akk5 +
            "','" +
            ls[li5] +
            "'," +
            tn2 +
            ',this.checked)" style="width:14px;height:14px;cursor:pointer;accent-color:#4caf50;">';
          h +=
            '<span style="font-size:11px;font-weight:700;color:' +
            (isC5 ? "#2e7d32" : "var(--gray-500)") +
            ';">' +
            ls[li5] +
            ". " +
            escapeAttr(w5) +
            "</span>";
          h += "</div>";
        }
        h += "</div>";
      }
    }

    // Dialogue: show dialogue lines + sentences + answer key
    if (tk2.type === "dialogue") {
      // Show dialogue lines
      if (tk2.dialogue && tk2.dialogue.length) {
        h +=
          '<div style="margin-bottom:14px;"><label style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:var(--gray-400);display:block;margin-bottom:8px;">Dialogue</label>';
        h +=
          '<div style="background:var(--gray-50);border-radius:8px;padding:12px;">';
        for (var dli = 0; dli < tk2.dialogue.length; dli++) {
          h +=
            '<div style="font-size:13px;margin:4px 0;padding:4px 0;border-bottom:1px solid var(--gray-100);">' +
            escapeAttr(tk2.dialogue[dli] || "") +
            "</div>";
        }
        h += "</div></div>";
      }
    }
    if (tk2.type === "dialogue" && tk2.dialogueSents) {
      var dsk2 = Object.keys(tk2.dialogueSents);
      h +=
        '<div style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:var(--gray-400);margin-bottom:8px;">Sentences</div>';
      for (var dsi2 = 0; dsi2 < dsk2.length; dsi2++) {
        var dsk2k = dsk2[dsi2];
        h +=
          '<div style="display:flex;align-items:center;gap:6px;margin-bottom:6px;">';
        h +=
          '<span style="font-size:11px;font-weight:700;color:var(--gray-500);width:16px;">' +
          dsk2k +
          ".</span>";
        h +=
          '<input type="text" value="' +
          escapeAttr(tk2.dialogueSents[dsk2k] || "") +
          '" onchange="previewEditDialogue(' +
          tn2 +
          ",'" +
          dsk2k +
          '\',this.value)" style="flex:1;padding:6px 10px;border:1.5px solid var(--gray-200);border-radius:6px;font-size:12px;">';
        h += "</div>";
      }
      var dl2 = tk2.dialogue || [];
      var dlGaps = 0;
      for (var dgi = 0; dgi < dl2.length; dgi++) {
        var dline = dl2[dgi] || "";
        if (
          dline.indexOf("...(") !== -1 ||
          dline.indexOf("(1)") !== -1 ||
          dline.indexOf("(2)") !== -1 ||
          /\(\d+\)/.test(dline)
        )
          dlGaps++;
      }
      // Also use question_count as fallback if no gaps detected in dialogue
      if (dlGaps === 0 && tk2.qcount) dlGaps = tk2.qcount;
      if (dlGaps > 0) {
        h +=
          '<div style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:var(--gray-400);margin-top:12px;margin-bottom:8px;">Answer Key (per gap)</div>';
        h += '<div style="display:flex;flex-wrap:wrap;gap:8px;">';
        for (var dg2 = 1; dg2 <= dlGaps; dg2++) {
          var akkD = "t" + tn2 + "_q" + dg2;
          var akvD = S.answerKey[akkD] || "";
          h +=
            '<div style="display:flex;align-items:center;gap:6px;background:var(--gray-50);border-radius:8px;padding:6px 10px;">';
          h +=
            '<span style="font-size:12px;font-weight:700;color:var(--pink);">' +
            dg2 +
            ".</span>";
          h +=
            "<select onchange=\"setAKDirect('" +
            akkD +
            '\',this.value)" style="padding:4px 8px;border:1.5px solid var(--gray-200);border-radius:6px;font-size:12px;">';
          h += '<option value="">--</option>';
          for (var ds3 = 0; ds3 < dsk2.length; ds3++)
            h +=
              '<option value="' +
              dsk2[ds3] +
              '"' +
              (akvD === dsk2[ds3] ? " selected" : "") +
              ">" +
              dsk2[ds3] +
              "</option>";
          h += "</select></div>";
        }
        h += "</div>";
      }
    }

    // Essay
    if (isEssay2) {
      h +=
        '<div style="background:var(--gray-50);border-radius:8px;padding:12px;font-size:13px;color:var(--gray-500);margin-bottom:10px;">Writing task — graded manually. No answer key needed.</div>';
      if (tk2.prompt) {
        h +=
          '<label style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:var(--gray-400);display:block;margin-bottom:4px;">Essay Prompt</label>';
        h +=
          '<textarea id="pedit-prompt-' +
          tn2 +
          '" rows="3" style="width:100%;padding:10px;border:1.5px solid var(--gray-200);border-radius:8px;font-size:13px;font-family:Georgia,serif;resize:vertical;">' +
          escapeAttr(tk2.prompt || "") +
          "</textarea>";
      }
    }

    h += "</div>";
  } // end for loop

  h +=
    '<button onclick="previewSaveEdits()" class="abtn abtn-pink" style="width:100%;justify-content:center;margin-top:8px;padding:14px;">Save All Changes &amp; Close</button>';
  h += "</div>";
  body.innerHTML = h;
}

function setAK(akKey, letter, tn2, checked) {
  if (!S.answerKey) S.answerKey = {};
  if (checked) S.answerKey[akKey] = letter;
  else delete S.answerKey[akKey];
  buildPreviewContent();
  setTimeout(function () {
    previewScrollTo("ptask-" + tn2);
  }, 60);
}

function setAKDirect(akKey, val) {
  if (!S.answerKey) S.answerKey = {};
  if (val) S.answerKey[akKey] = val;
  else delete S.answerKey[akKey];
}

function escapeAttr(str) {
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function previewScrollTo(id) {
  var el = document.getElementById(id);
  if (el) el.scrollIntoView({ behavior: "smooth", block: "start" });
}

function previewEditPassage(tn, key, val) {
  if (TASKS[tn] && TASKS[tn].passages) TASKS[tn].passages[key] = val;
}

function previewEditPts(tn, val) {
  if (TASKS[tn]) TASKS[tn].pts = parseInt(val) || 0;
}

function previewEditInstr(tn, val) {
  if (TASKS[tn]) TASKS[tn].instructions = val;
}

function previewEditQ(tn, qi, field, val) {
  if (!TASKS[tn] || !TASKS[tn].qs || !TASKS[tn].qs[qi]) return;
  if (field === "q") {
    TASKS[tn].qs[qi].q = val;
  } else if (field.indexOf("o") === 0) {
    var oi = parseInt(field.slice(1));
    if (!TASKS[tn].qs[qi].os) TASKS[tn].qs[qi].os = [];
    TASKS[tn].qs[qi].os[oi] = val;
  }
}

function previewEditWordBank(tn, key, val) {
  if (TASKS[tn] && TASKS[tn].wordBank) TASKS[tn].wordBank[key] = val;
}

function previewEditChoice(tn, gapIdx, optIdx, val) {
  if (TASKS[tn] && TASKS[tn].wordChoices && TASKS[tn].wordChoices[gapIdx]) {
    TASKS[tn].wordChoices[gapIdx][optIdx] = val;
  }
}

function previewEditDialogue(tn, key, val) {
  if (TASKS[tn] && TASKS[tn].dialogueSents) TASKS[tn].dialogueSents[key] = val;
}

function previewSaveEdits() {
  var introTa = document.getElementById("pedit-intro");
  if (introTa && S.parsedExam) S.parsedExam.introText = introTa.value;
  var taskKeys = Object.keys(TASKS)
    .map(Number)
    .filter(function (k) {
      return k >= 1 && TASKS[k] && TASKS[k].type;
    });
  for (var ti = 0; ti < taskKeys.length; ti++) {
    var tn = taskKeys[ti];
    var ta = document.getElementById("pedit-text-" + tn);
    if (ta && TASKS[tn]) TASKS[tn].text = ta.value;
    var tp = document.getElementById("pedit-prompt-" + tn);
    if (tp && TASKS[tn]) TASKS[tn].prompt = tp.value;
  }
  // Sync ALL changes from TASKS back to S.parsedExam and S.parsedTasks
  if (S.parsedExam && S.parsedExam.tasks) {
    for (var pi3 = 0; pi3 < S.parsedExam.tasks.length; pi3++) {
      var tn3 = S.parsedExam.tasks[pi3].task;
      var tk3 = TASKS[tn3];
      if (!tk3) continue;
      // Sync every field
      S.parsedExam.tasks[pi3].points = tk3.pts;
      S.parsedExam.tasks[pi3].instructions = tk3.instructions || "";
      if (tk3.text !== undefined) S.parsedExam.tasks[pi3].text = tk3.text;
      if (tk3.prompt !== undefined) S.parsedExam.tasks[pi3].prompt = tk3.prompt;
      if (tk3.qs)
        S.parsedExam.tasks[pi3].questions = tk3.qs.map(function (q) {
          return q.q;
        });
      if (tk3.qs)
        S.parsedExam.tasks[pi3].options = tk3.qs.map(function (q) {
          return q.os || [];
        });
      if (tk3.passages) S.parsedExam.tasks[pi3].passages = tk3.passages;
      if (tk3.wordBank) S.parsedExam.tasks[pi3].word_bank = tk3.wordBank;
      if (tk3.wordChoices) S.parsedExam.tasks[pi3].choices = tk3.wordChoices;
      if (tk3.dialogue) S.parsedExam.tasks[pi3].dialogue = tk3.dialogue;
      if (tk3.dialogueSents)
        S.parsedExam.tasks[pi3].options = tk3.dialogueSents;
    }
    S.parsedTasks = S.parsedExam.tasks;
  }
  // Save everything to Supabase immediately
  var examCode =
    S.activeExamCode || (S.examCodes ? Object.keys(S.examCodes)[0] : null);
  if (examCode && S.exams[examCode]) {
    dbUpsert(
      "exams",
      {
        code: examCode,
        parsed_exam: S.parsedExam,
        tasks: S.parsedTasks || [],
        answer_key: S.answerKey || {},
      },
      "code",
    );
  }
  showToast("Changes saved!", "success");
  closeExamPreview();
}

function closeExamPreview() {
  var el = document.getElementById("preview-overlay");
  if (el) el.style.display = "none";
}

function copyExamIDs() {
  var ids = (window._ebGeneratedIDs || []).join("\n");
  if (!ids) {
    showToast("No IDs generated yet", "");
    return;
  }
  safeCopy(ids, "IDs copied!");
}

function copyUnlockCode() {
  var el = document.getElementById("exam-unlock-display");
  if (el) safeCopy(el.textContent, "Unlock code copied!");
}

function genExamIDs() {
  var prefixEl = document.getElementById("eb-id-prefix");
  var countEl = document.getElementById("eb-id-count");
  var prefix = prefixEl ? prefixEl.value.trim() || "GE" : "GE";
  var count = countEl ? parseInt(countEl.value) || 30 : 30;
  window._ebGeneratedIDs = [];
  for (var i = 1; i <= count; i++) {
    var num = i < 10 ? "00" + i : i < 100 ? "0" + i : "" + i;
    window._ebGeneratedIDs.push(prefix + "-" + num);
  }
  // Show count label
  var countLabel = document.getElementById("exam-ids-count-label");
  if (countLabel)
    countLabel.textContent =
      count + " ID" + (count !== 1 ? "s" : "") + " generated";
  // Populate grid
  var grid = document.getElementById("exam-ids-grid");
  if (grid) {
    var html2 = "";
    for (var j = 0; j < window._ebGeneratedIDs.length; j++) {
      html2 +=
        '<span style="background:var(--gray-100);border-radius:6px;padding:2px 8px;font-family:monospace;font-size:11px;">' +
        window._ebGeneratedIDs[j] +
        "</span>";
    }
    grid.innerHTML = html2;
  }
  // Show the preview panel
  var preview = document.getElementById("exam-ids-preview");
  if (preview) preview.style.display = "block";
}

function removeAudio() {
  if (window.S_audioURL) {
    URL.revokeObjectURL(window.S_audioURL);
    window.S_audioURL = null;
  }
  window.S_audioFile = null;
  var player = document.getElementById("audio-player");
  if (player) {
    player.src = "";
    player.style.display = "none";
  }
  var preview = document.getElementById("audio-preview");
  if (preview) preview.style.display = "none";
  showToast("Audio removed", "");
}

async function scheduleExam(code) {
  var selEl = document.getElementById("sched-exam-sel");
  var dateEl = document.getElementById("sched-date");
  var timeEl = document.getElementById("sched-time");
  if (!selEl || !dateEl) return;
  var examCode = code || selEl.value;
  var date = dateEl.value;
  var time = timeEl ? timeEl.value : "10:00";
  if (!examCode || !date) {
    showToast("Please select an exam and date", "");
    return;
  }
  var ex = S.exams[examCode];
  if (!ex) return;
  S.schedule[examCode] = {
    date: date,
    time: time,
    title: ex.title,
    duration: ex.duration || 150,
  };
  await dbUpsert(
    "schedule",
    {
      exam_code: examCode,
      date: date,
      time: time,
      duration: ex.duration || 150,
    },
    "exam_code",
  );
  renderSchedTable();
  showToast("Exam scheduled!", "success");
}

async function handleAudioUpload(e) {
  var file = e.target ? e.target.files[0] : e.files ? e.files[0] : null;
  if (!file) return;
  // Revoke old blob URL if any
  if (window.S_audioURL && window.S_audioURL.startsWith("blob:"))
    URL.revokeObjectURL(window.S_audioURL);
  window.S_audioFile = file;
  // Show local preview immediately
  window.S_audioURL = URL.createObjectURL(file);
  S.audioURL = window.S_audioURL;
  var player = document.getElementById("audio-player");
  if (player) {
    player.src = window.S_audioURL;
    player.style.display = "block";
  }
  var nameEl = document.getElementById("audio-name");
  if (nameEl) nameEl.textContent = file.name + " (uploading...)";
  var preview = document.getElementById("audio-preview");
  if (preview) preview.style.display = "block";

  // Upload to Supabase Storage
  try {
    var fileName =
      "audio_" + Date.now() + "_" + file.name.replace(/[^a-zA-Z0-9._-]/g, "_");
    var { data, error } = await sb.storage
      .from("exam-audio")
      .upload(fileName, file, { upsert: true });
    if (error) throw error;
    // Get public URL
    var { data: urlData } = sb.storage
      .from("exam-audio")
      .getPublicUrl(fileName);
    var publicURL = urlData.publicUrl;
    window.S_audioURL = publicURL;
    S.audioURL = publicURL;
    if (player) player.src = publicURL;
    if (nameEl) nameEl.textContent = file.name + " ✓";
    showToast("Audio uploaded: " + file.name, "success");
  } catch (err) {
    // Fall back to blob URL if upload fails
    console.error("Audio upload error:", err);
    if (nameEl) nameEl.textContent = file.name + " (local only)";
    showToast("Audio saved locally (storage upload failed)", "");
  }
}

function onDrumScroll(col) {
  var drumId = col === "hours" ? "drum-hours" : "drum-mins";
  var drum = document.getElementById(drumId);
  if (!drum) return;
  var ITEM_H = 44;
  var idx = Math.round(drum.scrollTop / ITEM_H);
  var items = drum.querySelectorAll(".drum-item");
  if (idx < 0) idx = 0;
  if (idx >= items.length) idx = items.length - 1;
  var val = parseInt(items[idx].getAttribute("data-val")) || 0;

  // Update selected highlight
  for (var i = 0; i < items.length; i++) {
    items[i].className = i === idx ? "drum-item selected" : "drum-item";
  }

  // Read both drums to compute total
  var hDrum = document.getElementById("drum-hours");
  var mDrum = document.getElementById("drum-mins");
  var hIdx = hDrum ? Math.round(hDrum.scrollTop / ITEM_H) : 0;
  var mIdx = mDrum ? Math.round(mDrum.scrollTop / ITEM_H) : 6;
  var hItems = hDrum ? hDrum.querySelectorAll(".drum-item") : [];
  var mItems = mDrum ? mDrum.querySelectorAll(".drum-item") : [];
  if (hIdx >= hItems.length) hIdx = hItems.length - 1;
  if (mIdx >= mItems.length) mIdx = mItems.length - 1;
  var hVal = hItems.length
    ? parseInt(hItems[hIdx].getAttribute("data-val")) || 0
    : 2;
  var mVal = mItems.length
    ? parseInt(mItems[mIdx].getAttribute("data-val")) || 0
    : 30;
  var totalMins = hVal * 60 + mVal;

  // Save to hidden input
  var hidden = document.getElementById("eb-time-min");
  if (hidden) hidden.value = totalMins;

  // Update display text
  var disp = document.getElementById("drum-display");
  if (disp) {
    var hTxt = hVal > 0 ? hVal + " hour" + (hVal !== 1 ? "s" : "") : "";
    var mTxt = mVal > 0 ? mVal + " minute" + (mVal !== 1 ? "s" : "") : "";
    disp.textContent =
      hTxt && mTxt ? hTxt + " " + mTxt : hTxt || mTxt || "0 minutes";
  }
}
function onExamCodeInput() {
  var el = document.getElementById("eb-code");
  if (!el) return;
  el.value = el.value.toUpperCase().replace(/[^A-Z0-9-]/g, "");
  var code = el.value;
  // Show/update the unlock code banner
  var banner = document.getElementById("exam-unlock-preview");
  var display = document.getElementById("exam-unlock-display");
  if (banner && display) {
    if (code) {
      display.textContent = code;
      banner.style.display = "block";
    } else {
      banner.style.display = "none";
    }
  }
}

function renderTaskList() {
  var container = document.getElementById("task-list");
  if (!container) return;

  var examFilter = document.getElementById("task-exam-filter");
  var studentFilter = document.getElementById("task-student-filter");
  var filterExam = examFilter ? examFilter.value : "";
  var filterStudent = studentFilter ? studentFilter.value : "";

  // Populate exam filter
  if (examFilter && examFilter.options.length <= 1) {
    var examKeys2 = Object.keys(S.exams || {});
    for (var ei = 0; ei < examKeys2.length; ei++) {
      var opt = document.createElement("option");
      opt.value = examKeys2[ei];
      opt.textContent = S.exams[examKeys2[ei]].title || examKeys2[ei];
      examFilter.appendChild(opt);
    }
  }
  // Populate student filter
  if (studentFilter && studentFilter.options.length <= 1) {
    var regKeys = Object.keys(S.regCodes || {});
    for (var si = 0; si < regKeys.length; si++) {
      var u = S.regCodes[regKeys[si]];
      if (!u.activated) continue;
      var opt2 = document.createElement("option");
      opt2.value = u.email;
      opt2.textContent = u.name;
      studentFilter.appendChild(opt2);
    }
  }

  // Find all completed exams
  var rows = "";
  var compKeys = Object.keys(S.completed || {});

  if (!compKeys.length) {
    container.innerHTML =
      '<div style="color:var(--gray-400);font-size:13px;padding:16px 0;">No submitted tasks yet.</div>';
    return;
  }

  for (var ci = 0; ci < compKeys.length; ci++) {
    var compKey = compKeys[ci];
    var parts = compKey.split("|");
    var stEmail = parts[0];
    var exCode = parts[1];
    if (filterExam && exCode !== filterExam) continue;
    if (filterStudent && stEmail !== filterStudent) continue;

    var comp = S.completed[compKey];
    var ex = S.exams[exCode];
    if (!ex) continue;
    var stName = "";
    var rks = Object.keys(S.regCodes);
    for (var ri = 0; ri < rks.length; ri++) {
      if (S.regCodes[rks[ri]].email === stEmail) {
        stName = S.regCodes[rks[ri]].name;
        break;
      }
    }
    var grades = S.grades[compKey] || {};
    var taskCount = ex.tasks && ex.tasks.length ? ex.tasks.length : 7;

    rows +=
      '<div class="acard" style="margin-bottom:12px;">' +
      '<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:12px;flex-wrap:wrap;gap:6px;">' +
      "<div>" +
      '<div style="font-weight:700;font-size:14px;">' +
      (stName || stEmail) +
      "</div>" +
      '<div style="font-size:12px;color:var(--gray-500);">' +
      ex.title +
      " &middot; " +
      (comp.elapsedStr || "") +
      "</div>" +
      "</div>" +
      "</div>" +
      '<div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(110px,1fr));gap:8px;">';

    for (var ti = 1; ti <= taskCount; ti++) {
      var tKey2 = "t" + ti;
      var isEssay2 =
        ex.tasks && ex.tasks[ti - 1] && ex.tasks[ti - 1].type === "essay";
      var tg = isEssay2 ? grades.essay || null : grades[tKey2] || null;
      var maxPts =
        ex.tasks && ex.tasks[ti - 1] && ex.tasks[ti - 1].pts
          ? ex.tasks[ti - 1].pts
          : isEssay2
            ? 16
            : 10;
      var graded = tg ? true : false;
      var gradeBtn = isEssay2
        ? 'onclick="showSec(\'grading\')" title="Go to Grade Essays"'
        : 'onclick="openTaskGradeModal(&quot;' +
          stEmail +
          "&quot;,&quot;" +
          exCode +
          "&quot;," +
          ti +
          "," +
          maxPts +
          ')"';

      var tType2 =
        (ex.tasks && ex.tasks[ti - 1] && ex.tasks[ti - 1].type) || "";
      var tLbl2 =
        tType2 === "essay" || tType2 === "writing"
          ? "Writing"
          : tType2 === "listening"
            ? "Listening"
            : tType2 === "reading"
              ? "Reading"
              : tType2 === "match"
                ? "Matching"
                : tType2 === "gapfill4" || tType2 === "gapfill5"
                  ? "Gap Fill"
                  : tType2 === "dialogue"
                    ? "Dialogue"
                    : "";
      rows +=
        '<div class="rt" style="cursor:pointer;" ' +
        gradeBtn +
        ">" +
        '<div class="rt-label">Task ' +
        ti +
        (tLbl2
          ? "<br><span style='font-size:9px;font-weight:400;opacity:.65;'>" +
            tLbl2 +
            "</span>"
          : "") +
        "</div>" +
        (graded
          ? '<div class="rt-score">' +
            tg.score +
            '<span class="rt-max">/' +
            tg.max +
            "</span></div>"
          : '<div class="rt-score" style="font-size:12px;color:var(--pink);">Grade</div>') +
        "</div>";
    }
    rows += "</div></div>";
  }

  container.innerHTML =
    rows ||
    '<div style="color:var(--gray-400);font-size:13px;padding:16px 0;">No matching submissions.</div>';
}
function renderStudentResults() {
  var container = document.getElementById("student-results-section");
  if (!container || !S.student) return;
  // Always reload grades fresh from Supabase before rendering
  dbGet("grades").then(function (gradeRows) {
    gradeRows.forEach(function (r) {
      var key = r.student_email + "|" + r.exam_code;
      if (!S.grades[key]) S.grades[key] = {};
      var obj = {
        score: r.score,
        max: r.max_score,
        feedback: r.feedback || "",
      };
      if (r.fluency !== null && r.fluency !== undefined) {
        obj.fluency = r.fluency;
        obj.accuracy = r.accuracy;
      }
      S.grades[key][r.task_key] = obj;
    });
    _doRenderStudentResults();
  });
}
function _doRenderStudentResults() {
  var container = document.getElementById("student-results-section");
  if (!container || !S.student) return;
  var email = S.student.email;

  // Collect completed exams for this student
  var completedExams = [];
  var examKeys = Object.keys(S.exams || {});
  for (var i = 0; i < examKeys.length; i++) {
    var code = examKeys[i];
    var ex = S.exams[code];
    var sched = S.schedule && S.schedule[code];
    if (!sched) continue;
    var status = getExamStatus(sched.date, sched.time, sched.duration);
    var compKey2 = email + "|" + code;
    var wasCompleted = S.completed && S.completed[compKey2];
    if (status === "ended" || wasCompleted) {
      completedExams.push({
        code: code,
        exam: ex,
        sched: sched,
        compData: S.completed[compKey2] || null,
      });
    }
  }

  if (!completedExams.length) {
    container.style.display = "none";
    return;
  }
  container.style.display = "block";
  var html2 = '<h3 class="sec-title" style="margin-top:28px;">My Results</h3>';

  for (var j = 0; j < completedExams.length; j++) {
    var item = completedExams[j];
    var code = item.code;
    var ex = item.exam;
    var gradeKey = email + "|" + code;
    var grades = S.grades[gradeKey] || {};

    // Also merge essay grade stored under email|ESSAY key (legacy)
    var essayGrades = S.grades[email + "|ESSAY"] || {};
    if (essayGrades.essay && !grades.essay) grades.essay = essayGrades.essay;

    var taskCount = ex.tasks && ex.tasks.length ? ex.tasks.length : 7;
    var totalScore = 0,
      totalMax = ex.pts || 70,
      allGraded = true;
    // Recalculate totalMax from actual grade max values if all graded
    var gradeMaxSum = 0;
    var taskCards = "";

    for (var t = 1; t <= taskCount; t++) {
      var tKey = "t" + t;
      var isEssay =
        ex.tasks &&
        ex.tasks[t - 1] &&
        (ex.tasks[t - 1].type === "essay" ||
          ex.tasks[t - 1].type === "writing");
      var tType3 = (ex.tasks && ex.tasks[t - 1] && ex.tasks[t - 1].type) || "";
      var tLbl3 =
        tType3 === "essay" || tType3 === "writing"
          ? "Writing"
          : tType3 === "listening"
            ? "Listening"
            : tType3 === "reading"
              ? "Reading"
              : tType3 === "match"
                ? "Matching"
                : tType3 === "gapfill4" || tType3 === "gapfill5"
                  ? "Gap Fill"
                  : tType3 === "dialogue"
                    ? "Dialogue"
                    : "";
      var tGrade = isEssay
        ? grades["t" + t] ||
          grades["wt" + t] ||
          grades["essay"] ||
          grades["t7"] ||
          grades["wt7"]
        : grades[tKey];
      // Extra fallback: if no grade found by tKey, check all keys for a match
      if (!tGrade) {
        var allKeys = Object.keys(grades);
        for (var ki = 0; ki < allKeys.length; ki++) {
          if (allKeys[ki] === "t" + t || allKeys[ki] === "wt" + t) {
            tGrade = grades[allKeys[ki]];
            break;
          }
        }
      }

      if (tGrade) {
        totalScore += tGrade.score;
        gradeMaxSum += tGrade.max;
        var dotMark = tGrade.feedback ? " \u2022" : "";
        taskCards +=
          '<div class="rt" onclick="showTaskFeedback(\'' +
          code +
          "','" +
          tKey +
          '\')" style="cursor:pointer;">' +
          '<div class="rt-label">Task ' +
          t +
          dotMark +
          (tLbl3
            ? "<br><span style='font-size:9px;font-weight:400;opacity:.65;'>" +
              tLbl3 +
              "</span>"
            : "") +
          "</div>" +
          '<div class="rt-score">' +
          tGrade.score +
          '<span class="rt-max">/' +
          tGrade.max +
          "</span></div>" +
          "</div>";
      } else {
        allGraded = false;
        taskCards +=
          '<div class="rt pending">' +
          '<div class="rt-label">Task ' +
          t +
          (tLbl3
            ? "<br><span style='font-size:9px;font-weight:400;opacity:.65;'>" +
              tLbl3 +
              "</span>"
            : "") +
          "</div>" +
          '<div class="rt-score" style="font-size:13px;">Pending</div>' +
          '<div class="rt-max">' +
          (isEssay ? "Essay grading" : "Pending") +
          "</div>" +
          "</div>";
      }
    }

    // Total + pass/fail
    if (gradeMaxSum > 0) totalMax = gradeMaxSum;
    var passMark = Math.ceil((ex.pts || totalMax) * 0.25);
    var passed = totalScore >= passMark;
    var statusBadge;
    if (!allGraded) {
      statusBadge =
        '<span class="ec-status s-upcoming" style="font-size:10px;">Grading pending</span>';
    } else if (passed) {
      statusBadge =
        '<span class="ec-status" style="background:#e8f5e9;color:#2e7d32;font-size:10px;">PASS</span>';
    } else {
      statusBadge =
        '<span class="ec-status" style="background:#ffebee;color:#c62828;font-size:10px;">FAIL</span>';
    }

    var totalCard = allGraded
      ? '<div class="rt highlight"><div class="rt-label">Total</div><div class="rt-score">' +
        totalScore +
        '<span class="rt-max">/' +
        totalMax +
        "</span></div></div>"
      : '<div class="rt highlight"><div class="rt-label">Total</div><div class="rt-score" style="font-size:14px;">Pending</div></div>';

    var compInfo = item.compData;
    var metaTime = compInfo ? "Completed in " + compInfo.elapsedStr : "";

    html2 +=
      '<div class="exam-card">' +
      '<div class="ec-head">' +
      "<div>" +
      '<div class="ec-title">' +
      ex.title +
      "</div>" +
      '<div class="ec-meta"><span>' +
      (item.sched.date || "") +
      "</span>" +
      (metaTime ? "<span>" + metaTime + "</span>" : "") +
      "</div>" +
      "</div>" +
      '<div style="display:flex;flex-direction:column;align-items:flex-end;gap:6px;">' +
      '<span class="ec-status s-done">Completed</span>' +
      statusBadge +
      "</div>" +
      "</div>" +
      '<div class="results-grid">' +
      taskCards +
      totalCard +
      "</div>" +
      "</div>";
  }

  container.innerHTML = html2;
} // end _doRenderStudentResults

function showTaskFeedback(examCode, taskKey) {
  var email = S.student ? S.student.email : "";
  var gradeKey = email + "|" + examCode;
  var grade = S.grades[gradeKey] && S.grades[gradeKey][taskKey];
  // also check writing task key
  if (!grade) {
    var tNum = taskKey.replace("t", "");
    grade = S.grades[gradeKey] && S.grades[gradeKey]["wt" + tNum];
  }
  if (!grade) return;
  var taskNum = taskKey.replace("t", "");
  var isEssay = grade.fluency !== undefined;

  var modal = document.getElementById("info-modal");
  var body = document.getElementById("info-modal-body");
  var titleEl = modal ? modal.querySelector("h3") : null;
  if (!modal || !body) return;

  if (titleEl) titleEl.textContent = "Task " + taskNum + " Result";

  var h =
    '<div style="background:var(--gray-50);border-radius:10px;padding:14px 16px;font-size:13px;line-height:2;">';
  h +=
    '<div><strong>Score:</strong> <span style="font-size:18px;font-weight:700;color:var(--pink);">' +
    grade.score +
    "</span> / " +
    grade.max +
    "</div>";
  if (isEssay) {
    h +=
      "<div><strong>Fluency:</strong> " +
      grade.fluency +
      " / " +
      Math.floor(grade.max / 2) +
      "</div>";
    h +=
      "<div><strong>Accuracy:</strong> " +
      grade.accuracy +
      " / " +
      Math.floor(grade.max / 2) +
      "</div>";
  }
  h += "</div>";
  if (grade.feedback) {
    h +=
      '<div style="margin-top:10px;background:#fff;border:1.5px solid var(--gray-200);border-radius:10px;padding:14px 16px;">';
    h +=
      '<div style="font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:var(--pink);margin-bottom:6px;">Feedback</div>';
    h +=
      '<div style="font-size:13px;line-height:1.7;color:var(--charcoal);">' +
      grade.feedback +
      "</div>";
    h += "</div>";
  } else {
    h +=
      '<div style="margin-top:10px;font-size:12px;color:var(--gray-400);text-align:center;">No feedback provided for this task.</div>';
  }
  body.innerHTML = h;
  modal.classList.add("open");
}

function openTaskGradeModal(studentEmail, examCode, taskNum, maxScore) {
  var modal = document.getElementById("info-modal");
  var body = document.getElementById("info-modal-body");
  var titleEl = modal ? modal.querySelector("h3") : null;
  if (!modal || !body) return;
  if (titleEl)
    titleEl.textContent =
      "Task " + taskNum + " — " + studentEmail.split("@")[0];

  var gradeKey = studentEmail + "|" + examCode;
  var existing = S.grades[gradeKey] && S.grades[gradeKey]["t" + taskNum];
  var curFb = existing ? existing.feedback : "";

  // Get student answers
  var comp = S.completed[gradeKey];
  var taskAns = comp && comp.answers && comp.answers["t" + taskNum];
  var ex = S.exams[examCode];
  var taskDef = ex && ex.tasks && ex.tasks[taskNum - 1];
  var isWriting =
    taskDef && (taskDef.type === "essay" || taskDef.type === "writing");

  // Auto-grade from answer key
  var autoScore = 0;
  var qCount = 0;
  var ansHtml = "";

  if (isWriting) {
    var wText = (comp && comp.answers && comp.answers["t" + taskNum]) || "";
    ansHtml =
      '<div style="margin-bottom:14px;">' +
      '<div style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:var(--gray-400);margin-bottom:6px;">Student Response</div>' +
      '<div style="background:var(--gray-50);border:1px solid var(--gray-200);border-radius:8px;padding:10px 12px;font-size:13px;line-height:1.7;white-space:pre-wrap;max-height:200px;overflow-y:auto;">' +
      (wText || '<span style="color:var(--gray-400);">No response</span>') +
      "</div></div>";
    // Writing - manual grade
    var curScore = existing ? existing.score : "";
    var curMax = existing ? existing.max : maxScore || 16;
    ansHtml +=
      '<div style="display:flex;gap:10px;margin-bottom:10px;">' +
      '<div style="flex:1;"><label style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:var(--gray-400);display:block;margin-bottom:4px;">Score</label>' +
      '<input type="number" id="tg-score" value="' +
      curScore +
      '" min="0" max="' +
      (maxScore || 16) +
      '" style="width:100%;padding:8px 10px;border:1.5px solid var(--gray-200);border-radius:8px;font-size:14px;outline:none;"></div>' +
      '<div style="flex:1;"><label style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:var(--gray-400);display:block;margin-bottom:4px;">Out of</label>' +
      '<input type="number" id="tg-max" value="' +
      (maxScore || curMax) +
      '" min="1" style="width:100%;padding:8px 10px;border:1.5px solid var(--gray-200);border-radius:8px;font-size:14px;outline:none;"></div>' +
      "</div>";
  } else if (taskAns && typeof taskAns === "object") {
    // Determine question count
    var qcnt = 0;
    if (taskDef && taskDef.qs) qcnt = taskDef.qs.length;
    else if (taskDef && taskDef.qcount) qcnt = taskDef.qcount;
    else {
      var ks = Object.keys(taskAns)
        .map(Number)
        .filter(function (x) {
          return !isNaN(x);
        });
      if (ks.length) qcnt = Math.max.apply(null, ks);
    }

    // Check against answer key
    var hasAK = false;
    for (var qi2 = 1; qi2 <= qcnt; qi2++) {
      var akKey2 = "t" + taskNum + "_q" + qi2;
      if (S.answerKey && S.answerKey[akKey2]) {
        hasAK = true;
        break;
      }
    }

    ansHtml = '<div style="margin-bottom:14px;">';
    ansHtml +=
      '<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:8px;">';
    ansHtml +=
      '<div style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:var(--gray-400);">Student Answers</div>';
    if (!hasAK)
      ansHtml +=
        '<div style="font-size:11px;color:#f59e0b;font-weight:600;">No answer key set — score manually below</div>';
    ansHtml += "</div>";
    ansHtml += '<div style="display:flex;flex-wrap:wrap;gap:6px;">';

    for (var qi3 = 1; qi3 <= qcnt; qi3++) {
      var stuAns = taskAns[qi3];
      var isBlank = stuAns === undefined || stuAns === null || stuAns === "";
      var correctAns = S.answerKey && S.answerKey["t" + taskNum + "_q" + qi3];
      var isCorrect = !isBlank && correctAns && stuAns === correctAns;
      var isWrong = !isBlank && correctAns && stuAns !== correctAns;
      qCount++;
      if (isCorrect) autoScore++;

      var bg = isBlank
        ? "#fff0f0"
        : isCorrect
          ? "#e8f5e9"
          : isWrong
            ? "#fff0f0"
            : "var(--gray-100)";
      var col = isBlank
        ? "#c00"
        : isCorrect
          ? "#2e7d32"
          : isWrong
            ? "#c62828"
            : "var(--charcoal)";
      var icon = isCorrect ? " ✓" : isWrong ? " ✗" : "";

      ansHtml +=
        '<div style="background:' +
        bg +
        ";border-radius:8px;padding:5px 10px;font-size:12px;font-family:monospace;color:" +
        col +
        ";border:1px solid " +
        (isCorrect ? "#a5d6a7" : isWrong ? "#ef9a9a" : "var(--gray-200)") +
        ';">';
      ansHtml +=
        '<span style="color:var(--gray-500);">' +
        qi3 +
        ".</span> <strong>" +
        (isBlank ? "—" : stuAns) +
        "</strong>" +
        icon;
      if (isWrong && correctAns)
        ansHtml +=
          ' <span style="font-size:10px;color:#2e7d32;">(' +
          correctAns +
          ")</span>";
      ansHtml += "</div>";
    }
    ansHtml += "</div>";
    if (hasAK && qCount > 0) {
      ansHtml +=
        '<div style="margin-top:10px;padding:10px;background:' +
        (autoScore === qCount ? "#e8f5e9" : "var(--gray-50)") +
        ";border-radius:8px;font-size:13px;font-weight:600;color:" +
        (autoScore === qCount ? "#2e7d32" : "var(--charcoal)") +
        ';">Auto score: ' +
        autoScore +
        " / " +
        qCount +
        "</div>";
    }
    ansHtml += "</div>";

    // Score inputs pre-filled with auto score
    var finalScore = existing ? existing.score : hasAK ? autoScore : "";
    var finalMax = existing ? existing.max : qCount || maxScore || 0;
    ansHtml +=
      '<div style="display:flex;gap:10px;margin-bottom:10px;">' +
      '<div style="flex:1;"><label style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:var(--gray-400);display:block;margin-bottom:4px;">Score</label>' +
      '<input type="number" id="tg-score" value="' +
      finalScore +
      '" min="0" max="' +
      (finalMax || qCount || maxScore || 100) +
      '" style="width:100%;padding:8px 10px;border:1.5px solid var(--gray-200);border-radius:8px;font-size:14px;outline:none;"></div>' +
      '<div style="flex:1;"><label style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:var(--gray-400);display:block;margin-bottom:4px;">Out of</label>' +
      '<input type="number" id="tg-max" value="' +
      (finalMax || qCount || maxScore || 0) +
      '" min="1" style="width:100%;padding:8px 10px;border:1.5px solid var(--gray-200);border-radius:8px;font-size:14px;outline:none;"></div>' +
      "</div>";
  } else {
    ansHtml =
      '<div style="color:var(--gray-400);font-size:13px;margin-bottom:14px;">No answers submitted for this task.</div>' +
      '<div style="display:flex;gap:10px;margin-bottom:10px;">' +
      '<div style="flex:1;"><label style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:var(--gray-400);display:block;margin-bottom:4px;">Score</label>' +
      '<input type="number" id="tg-score" value="" min="0" max="' +
      (maxScore || 100) +
      '" style="width:100%;padding:8px 10px;border:1.5px solid var(--gray-200);border-radius:8px;font-size:14px;outline:none;"></div>' +
      '<div style="flex:1;"><label style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:var(--gray-400);display:block;margin-bottom:4px;">Out of</label>' +
      '<input type="number" id="tg-max" value="' +
      (maxScore || 0) +
      '" min="1" style="width:100%;padding:8px 10px;border:1.5px solid var(--gray-200);border-radius:8px;font-size:14px;outline:none;"></div>' +
      "</div>";
  }

  body.innerHTML =
    '<div style="font-size:12px;color:var(--gray-500);margin-bottom:14px;">Exam: <strong>' +
    examCode +
    "</strong></div>" +
    ansHtml +
    '<label style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:var(--gray-400);display:block;margin-bottom:4px;">Feedback (optional)</label>' +
    '<textarea id="tg-feedback" rows="2" style="width:100%;padding:8px 10px;border:1.5px solid var(--gray-200);border-radius:8px;font-size:13px;outline:none;resize:vertical;">' +
    curFb +
    "</textarea>" +
    '<button class="abtn abtn-pink" style="width:100%;justify-content:center;margin-top:12px;" onclick="doSaveTaskGrade(&quot;' +
    studentEmail +
    "&quot;,&quot;" +
    examCode +
    "&quot;," +
    taskNum +
    "," +
    (maxScore || 0) +
    ')">Save Grade</button>';

  modal.classList.add("open");
}
async function doSaveTaskGrade(studentEmail, examCode, taskNum, defaultMax) {
  var scoreEl = document.getElementById("tg-score");
  var maxEl = document.getElementById("tg-max");
  var fbEl = document.getElementById("tg-feedback");
  if (!scoreEl || !maxEl) return;
  var score = parseInt(scoreEl.value);
  var max = parseInt(maxEl.value) || defaultMax;
  var fb = fbEl ? fbEl.value.trim() : "";
  if (isNaN(score) || score < 0) {
    showToast("Enter a valid score", "");
    return;
  }
  if (score > max) {
    showToast("Score cannot exceed " + max, "");
    return;
  }
  await saveTaskGrade(studentEmail, examCode, taskNum, score, max, fb);
  closeOverlay("info-modal");
}

/* ── Drag & Drop / Tap functions for gapfill4 ── */

function wbChipTap(tn, key, el) {
  // If already selected, deselect
  var prev = document.querySelector(".wb-chip.wb-selected");
  if (prev && prev === el) {
    prev.classList.remove("wb-selected");
    return;
  }
  if (prev) prev.classList.remove("wb-selected");
  el.classList.add("wb-selected");
}

function gapTap(tn, gn, el) {
  var selected = document.querySelector(".wb-chip.wb-selected");
  if (!selected) {
    // If gap has a word, return it to bank
    var curKey = S.answers["t" + tn] && S.answers["t" + tn][gn];
    if (curKey) {
      delete S.answers["t" + tn][gn];
      // Restore chip
      var chip = document.querySelector('.wb-chip[data-key="' + curKey + '"]');
      if (chip) chip.classList.remove("wb-placed");
      el.innerHTML = '<span class="gap-arrow">&#9660;</span>';
      el.classList.remove("gap-filled");
    }
    return;
  }
  var key = selected.getAttribute("data-key");
  var val = selected.getAttribute("data-val");
  // If this gap already has a word, return old word to bank
  var oldKey = S.answers["t" + tn] && S.answers["t" + tn][gn];
  if (oldKey && oldKey !== key) {
    var oldChip = document.querySelector('.wb-chip[data-key="' + oldKey + '"]');
    if (oldChip) oldChip.classList.remove("wb-placed");
  }
  // Place word in gap
  if (!S.answers["t" + tn]) S.answers["t" + tn] = {};
  // Remove key from any other gap it was in
  var allGaps = document.querySelectorAll(".gap-drop");
  for (var i = 0; i < allGaps.length; i++) {
    var g = allGaps[i];
    var gNum = parseInt(g.getAttribute("data-gn"));
    var gTn = parseInt(g.getAttribute("data-tn"));
    if (gTn === tn && S.answers["t" + tn][gNum] === key && gNum !== gn) {
      delete S.answers["t" + tn][gNum];
      g.innerHTML = '<span class="gap-arrow">&#9660;</span>';
      g.classList.remove("gap-filled");
    }
  }
  S.answers["t" + tn][gn] = key;
  el.textContent = val;
  el.classList.add("gap-filled");
  selected.classList.remove("wb-selected");
  selected.classList.add("wb-placed");
}

function dropOnGap(e, tn, gn) {
  e.preventDefault();
  var key = e.dataTransfer.getData("text/plain");
  if (!key) return;
  var chip = document.querySelector('.wb-chip[data-key="' + key + '"]');
  var val = chip ? chip.getAttribute("data-val") : key;
  var el = document.getElementById("gap-t" + tn + "-" + gn);
  if (!el) return;
  // Return old word if gap was filled
  var oldKey = S.answers["t" + tn] && S.answers["t" + tn][gn];
  if (oldKey && oldKey !== key) {
    var oldChip = document.querySelector('.wb-chip[data-key="' + oldKey + '"]');
    if (oldChip) oldChip.classList.remove("wb-placed");
  }
  // Remove key from any other gap
  var allGaps2 = document.querySelectorAll(".gap-drop");
  for (var i = 0; i < allGaps2.length; i++) {
    var g2 = allGaps2[i];
    var gNum2 = parseInt(g2.getAttribute("data-gn"));
    var gTn2 = parseInt(g2.getAttribute("data-tn"));
    if (
      gTn2 === tn &&
      S.answers["t" + tn] &&
      S.answers["t" + tn][gNum2] === key &&
      gNum2 !== gn
    ) {
      delete S.answers["t" + tn][gNum2];
      g2.innerHTML = '<span class="gap-arrow">&#9660;</span>';
      g2.classList.remove("gap-filled");
    }
  }
  if (!S.answers["t" + tn]) S.answers["t" + tn] = {};
  S.answers["t" + tn][gn] = key;
  el.textContent = val;
  el.classList.add("gap-filled");
  if (chip) {
    chip.classList.remove("wb-dragging");
    chip.classList.add("wb-placed");
  }
}

function initDragDrop() {
  var chips = document.querySelectorAll(".wb-chip");
  for (var i = 0; i < chips.length; i++) {
    chips[i].ondragstart = function (e) {
      e.dataTransfer.setData("text/plain", this.getAttribute("data-key"));
      this.classList.add("wb-dragging");
    };
    chips[i].ondragend = function (e) {
      this.classList.remove("wb-dragging");
    };
  }
}

/* ── Inline dropdown for gapfill5 ── */
function selDD(tn, q, el) {
  if (!S.answers["t" + tn]) S.answers["t" + tn] = {};
  S.answers["t" + tn][q] = el.value;
}

// ── Load all data from Supabase ───────────────────────────────────────────────
async function loadAllData() {
  try {
    var students = await dbGet("students");
    S.regCodes = {};
    students.forEach(function (s) {
      S.regCodes[s.reg_code] = {
        name: s.name,
        email: s.email,
        phone: s.phone || "",
        pw: s.password,
        activated: s.activated,
        id: s.student_id || "",
      };
    });
    var exams = await dbGet("exams");
    S.exams = {};
    S.examCodes = {};
    exams.forEach(function (e) {
      S.exams[e.code] = {
        title: e.title,
        code: e.code,
        duration: e.duration || 150,
        hasAudio: e.has_audio,
        tasks: e.tasks || [],
        info: e.info || "",
        parsedExam: e.parsed_exam || null,
        pts: e.pts || 70,
        audioURL: e.audio_url || null,
      };
      S.examCodes[e.code] = true;
      // Load answer key
      if (e.answer_key && Object.keys(e.answer_key).length) {
        if (!S.answerKey) S.answerKey = {};
        Object.assign(S.answerKey, e.answer_key);
      }
    });
    var examIds = await dbGet("exam_ids");
    S.examIDs = {};
    S.usedIDs = {};
    examIds.forEach(function (r) {
      if (!S.examIDs[r.exam_code]) S.examIDs[r.exam_code] = [];
      S.examIDs[r.exam_code].push(r.student_id);
      if (r.used) S.usedIDs[r.exam_code + ":" + r.student_id] = true;
    });
    var schedRows = await dbGet("schedule");
    S.schedule = {};
    schedRows.forEach(function (r) {
      S.schedule[r.exam_code] = {
        date: r.date,
        time: r.time,
        title:
          (S.exams[r.exam_code] && S.exams[r.exam_code].title) || r.exam_code,
        duration: r.duration || 150,
      };
    });
    var subs = await dbGet("submissions");
    S.completed = {};
    subs.forEach(function (r) {
      if (r.elapsed_str !== "in_progress") {
        S.completed[r.student_email + "|" + r.exam_code] = {
          elapsedStr: r.elapsed_str || "",
          answers: r.answers || {},
          submittedAt: r.submitted_at,
          examStartTime: r.exam_start_time || 0,
        };
      }
    });
    var gradeRows = await dbGet("grades");
    S.grades = {};
    gradeRows.forEach(function (r) {
      var key = r.student_email + "|" + r.exam_code;
      if (!S.grades[key]) S.grades[key] = {};
      var obj = {
        score: r.score,
        max: r.max_score,
        feedback: r.feedback || "",
      };
      if (r.fluency !== null && r.fluency !== undefined) {
        obj.fluency = r.fluency;
        obj.accuracy = r.accuracy;
      }
      S.grades[key][r.task_key] = obj;
    });
    var anns = await dbGet("announcements");
    if (anns.length) {
      var latest = anns[anns.length - 1];
      var tEl = document.getElementById("ann-title"),
        bEl = document.getElementById("ann-body");
      if (tEl) tEl.textContent = latest.title;
      if (bEl) bEl.textContent = latest.body;
      if (latest.pinned) {
        var pt = document.getElementById("pin-title"),
          pb = document.getElementById("pin-body");
        if (pt) pt.textContent = latest.title;
        if (pb) pb.textContent = latest.body;
        var pa = document.getElementById("pinned-ann");
        if (pa) pa.classList.add("show");
      }
    }
  } catch (e) {
    console.error("loadAllData error", e);
  }
  // Refresh student exam list if student is logged in
  if (S.student) {
    renderStudentExams();
    renderStudentResults();
  }
  // Refresh all admin panels if admin is logged in
  refreshStudentsTable();
  refreshCodesTable();
  renderExamList();
  renderSchedTable();
  updateAdminDashboard();
}

function updateAdminDashboard() {
  // Students
  var codes = Object.values(S.regCodes || {});
  var total = codes.filter(function (u) {
    return u.activated;
  }).length;
  var pending = codes.filter(function (u) {
    return !u.activated;
  }).length;
  var sv = document.getElementById("stat-students");
  var ss = document.getElementById("stat-students-sub");
  if (sv) sv.textContent = total;
  if (ss) ss.textContent = pending + " pending";

  // Scheduled / live
  var schedKeys = Object.keys(S.schedule || {});
  var liveCount = 0,
    scheduledCount = 0,
    liveCode = null;
  schedKeys.forEach(function (code) {
    var sc = S.schedule[code];
    var st = getExamStatus(sc.date, sc.time, sc.duration);
    if (st === "live") {
      liveCount++;
      liveCode = code;
    }
    if (st === "upcoming") {
      scheduledCount++;
    }
  });
  var sv2 = document.getElementById("stat-scheduled");
  var ss2 = document.getElementById("stat-scheduled-sub");
  if (sv2) sv2.textContent = scheduledCount + liveCount;
  if (ss2) ss2.textContent = liveCount + " live";

  // Essays to grade
  var essayCount = 0;
  Object.keys(S.completed || {}).forEach(function (ck) {
    var parts = ck.split("|");
    var exCode = parts[1];
    var ex = S.exams[exCode];
    if (!ex || !ex.tasks) return;
    var grades = S.grades[ck] || {};
    ex.tasks.forEach(function (t, ti) {
      if (t.type === "essay" || t.type === "writing") {
        if (!grades["t" + (ti + 1)]) essayCount++;
      }
    });
  });
  var sv3 = document.getElementById("stat-essays");
  if (sv3) sv3.textContent = essayCount;

  // Submissions
  var subCount = Object.keys(S.completed || {}).length;
  var sv4 = document.getElementById("stat-submissions");
  var ss4 = document.getElementById("stat-submissions-sub");
  if (sv4) sv4.textContent = subCount;
  if (ss4) ss4.textContent = "total";

  // Live exam card
  var liveCard = document.getElementById("admin-live-card");
  var liveTitle = document.getElementById("admin-live-title");
  var liveInfo = document.getElementById("admin-live-info");
  if (liveCard) {
    if (liveCode) {
      liveCard.style.display = "block";
      var ex = S.exams[liveCode];
      if (liveTitle)
        liveTitle.textContent = " Live — " + (ex ? ex.title : liveCode);
      var sc = S.schedule[liveCode];
      var startMs = new Date(sc.date + "T" + sc.time + ":00").getTime();
      var totalSecs = (sc.duration || 150) * 60;
      var elapsedSecs = Math.floor((Date.now() - startMs) / 1000);
      var remSecs = Math.max(0, totalSecs - elapsedSecs);
      var rh = Math.floor(remSecs / 3600),
        rm = Math.floor((remSecs % 3600) / 60),
        rs = remSecs % 60;
      var joined = 0,
        submitted = 0;
      Object.keys(S.usedIDs || {}).forEach(function (k) {
        if (k.startsWith(liveCode + ":")) joined++;
      });
      Object.keys(S.completed || {}).forEach(function (ck) {
        if (ck.endsWith("|" + liveCode)) submitted++;
      });
      var totalIds =
        S.examIDs && S.examIDs[liveCode] ? S.examIDs[liveCode].length : 0;
      if (liveInfo)
        liveInfo.innerHTML =
          "<span><strong>" +
          joined +
          "</strong> joined</span>" +
          "<span><strong>" +
          (totalIds - joined) +
          "</strong> not joined</span>" +
          "<span><strong>" +
          submitted +
          "</strong> submitted</span>" +
          "<span>Remaining: <strong>" +
          rh +
          ":" +
          pad(rm) +
          ":" +
          pad(rs) +
          "</strong></span>";
    } else {
      liveCard.style.display = "none";
    }
  }
}

function startRealtimeSubscriptions() {
  sb.channel("subs")
    .on(
      "postgres_changes",
      { event: "*", schema: "public", table: "submissions" },
      function (payload) {
        var r = payload.new;
        if (!r) return;
        if (r.elapsed_str !== "in_progress") {
          S.completed[r.student_email + "|" + r.exam_code] = {
            elapsedStr: r.elapsed_str || "",
            answers: r.answers || {},
            submittedAt: r.submitted_at,
          };
        } else {
          delete S.completed[r.student_email + "|" + r.exam_code];
        }
        var el = document.getElementById("sec-tasks");
        if (el && el.classList.contains("active")) renderTaskList();
        var el2 = document.getElementById("sec-grading");
        if (el2 && el2.classList.contains("active")) buildEssays();
      },
    )
    .subscribe();
  sb.channel("grd")
    .on(
      "postgres_changes",
      { event: "*", schema: "public", table: "grades" },
      function (payload) {
        var r = payload.new;
        if (!r) return;
        var key = r.student_email + "|" + r.exam_code;
        if (!S.grades[key]) S.grades[key] = {};
        var obj = {
          score: r.score,
          max: r.max_score,
          feedback: r.feedback || "",
        };
        if (r.fluency !== null && r.fluency !== undefined) {
          obj.fluency = r.fluency;
          obj.accuracy = r.accuracy;
        }
        S.grades[key][r.task_key] = obj;
        if (S.student && S.student.email === r.student_email)
          renderStudentResults();
      },
    )
    .subscribe();
}

// ── Helper functions used by admin panel ─────────────────────────────────────
function refreshStudentsTable() {
  var tbody = document.getElementById("stud-tbody");
  if (!tbody) return;
  var rows = "";
  Object.keys(S.regCodes).forEach(function (rc) {
    var u = S.regCodes[rc];
    if (!u.name && !u.activated) return;
    var bdg = u.activated
      ? '<span class="bdg bdg-green">Active</span>'
      : '<span class="bdg bdg-orange">Pending</span>';
    rows +=
      "<tr><td><strong>" +
      (u.name || "-") +
      '</strong></td><td><span class="code-tag">' +
      rc +
      "</span></td><td>" +
      (u.email || "-") +
      "</td><td>" +
      (u.phone || "-") +
      '</td><td><span class="code-tag">' +
      (u.id || "-") +
      "</span></td><td>" +
      bdg +
      "</td></tr>";
  });
  tbody.innerHTML =
    rows ||
    '<tr><td colspan="6" style="text-align:center;color:var(--gray-400);padding:16px;">No students yet.</td></tr>';
}

function refreshCodesTable() {
  var tbody = document.getElementById("codes-tbody");
  if (!tbody) return;
  var rows = "";
  Object.keys(S.regCodes).forEach(function (rc) {
    var u = S.regCodes[rc];
    var bdg = u.activated
      ? '<span class="bdg bdg-green">Used</span>'
      : '<span class="bdg bdg-orange">Unused</span>';
    rows +=
      '<tr><td><span class="code-tag">' +
      rc +
      "</span></td><td>" +
      bdg +
      "</td><td>" +
      (u.name || "&mdash;") +
      '</td><td><button class="abtn abtn-ghost" style="padding:3px 8px;font-size:11px;" onclick="revokeCode(this,\'' +
      rc +
      "')\">Revoke</button></td></tr>";
  });
  tbody.innerHTML =
    rows ||
    '<tr id="codes-placeholder"><td colspan="4" style="text-align:center;color:var(--gray-400);padding:16px;">No codes yet.</td></tr>';
}

function renderResultsTable() {
  var tbody = document.querySelector("#sec-results tbody");
  if (!tbody) return;
  var rows = "";
  Object.keys(S.completed || {}).forEach(function (ck) {
    var pts = ck.split("|");
    var stEmail = pts[0],
      exCode = pts[1];
    var grades = S.grades[ck] || {};
    var ex = S.exams[exCode];
    var stObj = Object.values(S.regCodes).find(function (u) {
      return u.email === stEmail;
    });
    var name = stObj ? stObj.name : stEmail;
    var taskCount = ex && ex.tasks ? ex.tasks.length : 7;
    var total = 0,
      allGraded = true,
      cells = "";
    for (var ti = 1; ti <= taskCount; ti++) {
      var tg = grades["t" + ti];
      if (tg) {
        total += tg.score;
        cells += "<td>" + tg.score + "/" + tg.max + "</td>";
      } else {
        allGraded = false;
        cells += '<td style="color:var(--gray-400);">—</td>';
      }
    }
    rows +=
      "<tr><td><strong>" +
      name +
      "</strong></td>" +
      cells +
      "<td>" +
      (allGraded
        ? "<strong>" + total + "</strong>"
        : '<span style="color:var(--gray-400);">Pending</span>') +
      "</td></tr>";
  });
  tbody.innerHTML =
    rows ||
    '<tr><td colspan="10" style="text-align:center;color:var(--gray-400);padding:20px;">No results yet.</td></tr>';
}

function renderSchedTable() {
  var tbody = document.getElementById("sched-tbody");
  if (!tbody) return;
  var keys = Object.keys(S.schedule || {});
  if (!keys.length) {
    tbody.innerHTML =
      '<tr><td colspan="7" style="text-align:center;color:var(--gray-400);padding:20px;font-size:13px;">No exams scheduled yet.</td></tr>';
    return;
  }
  var rows = "";
  keys.forEach(function (code) {
    var sc = S.schedule[code];
    var ex = S.exams[code];
    var status = getExamStatus(sc.date, sc.time, sc.duration);
    var dur = sc.duration || 150;
    var dh = Math.floor(dur / 60),
      dm = dur % 60;
    var durStr = (dh > 0 ? dh + "h " : "") + dm + "min";
    var bdg =
      status === "live"
        ? '<span class="bdg bdg-green">Live</span>'
        : status === "upcoming"
          ? '<span class="bdg bdg-orange">Upcoming</span>'
          : '<span class="bdg" style="background:var(--gray-100);color:var(--gray-600);">Ended</span>';
    rows +=
      "<tr><td><strong>" +
      (ex ? ex.title : code) +
      '</strong></td><td><span class="code-tag">' +
      code +
      "</span></td><td>" +
      (sc.date || "-") +
      "</td><td>" +
      (sc.time || "-") +
      "</td><td>" +
      durStr +
      "</td><td>" +
      bdg +
      '</td><td><button class="abtn abtn-ghost" style="padding:3px 10px;font-size:11px;" onclick="openExamEdit(\'+code+\')">' +
      (status === "live" ? "Manage" : status === "ended" ? "Renew" : "Edit") +
      "</button></td></tr>";
  });
  tbody.innerHTML = rows;
}
