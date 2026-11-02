"""Easy Exam: the teacher uploads questions and runs the exam, students type a name and roll number.

Needs only Python 3.9+. Everything is stored in exam.db next to this file.
    python exam.py                    start the exam server
    python exam.py --reset-password   forget the teacher password
"""
import csv
import hashlib
import html
import io
import json
import os
import re
import secrets
import socket
import sqlite3
import sys
import threading
import time
import traceback
import webbrowser
import xml.etree.ElementTree as ET
import zipfile
from email.parser import BytesParser
from http.cookies import SimpleCookie
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, quote, urlsplit

HERE = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.environ.get("EASYEXAM_DATA", HERE)
DB_PATH = os.path.join(DATA_DIR, "exam.db")
PORT = int(os.environ.get("EASYEXAM_PORT", "1001"))
MAX_BODY = 10 * 1024 * 1024
TEACHER_SESSION_SECONDS = 12 * 3600
LETTERS = "abcd"

e = lambda value: html.escape(str(value), quote=True)
fmt = lambda number: f"{number:g}"


# ---------------------------------------------------------------- database

def connect():
    conn = sqlite3.connect(DB_PATH, timeout=20)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA synchronous=NORMAL")
    return conn


def init_db():
    os.makedirs(DATA_DIR, exist_ok=True)
    conn = connect()
    conn.execute("PRAGMA journal_mode=WAL")
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, value TEXT);
        CREATE TABLE IF NOT EXISTS sessions (token TEXT PRIMARY KEY, created REAL);
        CREATE TABLE IF NOT EXISTS questions (
            id INTEGER PRIMARY KEY, text TEXT NOT NULL, a TEXT, b TEXT, c TEXT, d TEXT,
            correct TEXT NOT NULL, marks REAL NOT NULL DEFAULT 1);
        CREATE TABLE IF NOT EXISTS students (
            id INTEGER PRIMARY KEY, name TEXT NOT NULL, roll TEXT NOT NULL,
            roll_key TEXT NOT NULL UNIQUE, token TEXT, submitted INTEGER NOT NULL DEFAULT 0);
        CREATE TABLE IF NOT EXISTS answers (
            student_id INTEGER, question_id INTEGER, choice TEXT,
            PRIMARY KEY (student_id, question_id));
    """)
    conn.commit()
    conn.close()


def get_setting(db, key, default=""):
    row = db.execute("SELECT value FROM settings WHERE key=?", (key,)).fetchone()
    return row["value"] if row else default


def set_setting(db, key, value):
    db.execute("INSERT OR REPLACE INTO settings (key, value) VALUES (?, ?)", (key, str(value)))


def exam_state(db):
    """Returns setup, running or ended. A running exam ends by itself when the time is up."""
    state = get_setting(db, "state", "setup")
    if state == "running" and time.time() >= float(get_setting(db, "ends_at", "0")):
        set_setting(db, "state", "ended")
        db.commit()
        return "ended"
    return state


def seconds_left(db):
    return max(0, int(float(get_setting(db, "ends_at", "0")) - time.time()))


def hash_password(password, salt=None):
    salt = salt or secrets.token_hex(16)
    digest = hashlib.pbkdf2_hmac("sha256", password.encode(), salt.encode(), 200_000).hex()
    return f"{salt}${digest}"


def check_password(password, stored):
    return secrets.compare_digest(hash_password(password, stored.split("$")[0]), stored)


def student_rows(db):
    """All students with their answered count and score, in roll number order."""
    rows = db.execute("""
        SELECT s.id, s.name, s.roll, s.token, s.submitted, COUNT(a.question_id) AS answered,
               COALESCE(SUM(CASE WHEN a.choice = q.correct THEN q.marks END), 0) AS score
        FROM students s
        LEFT JOIN answers a ON a.student_id = s.id
        LEFT JOIN questions q ON q.id = a.question_id
        GROUP BY s.id""").fetchall()
    natural = lambda r: [int(p) if p.isdigit() else p.lower() for p in re.split(r"(\d+)", r["roll"])]
    return sorted(rows, key=lambda r: [(isinstance(p, str), p) for p in natural(r)])


def question_totals(db):
    row = db.execute("SELECT COUNT(*) AS n, COALESCE(SUM(marks), 0) AS total FROM questions").fetchone()
    return row["n"], row["total"]


def student_status(row, state):
    if row["submitted"]:
        return "Submitted"
    return {"setup": "Waiting", "running": "Writing", "ended": "Finished"}[state]


# ---------------------------------------------------------------- question files

def read_xlsx(data):
    with zipfile.ZipFile(io.BytesIO(data)) as book:
        ns = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
        shared = []
        if "xl/sharedStrings.xml" in book.namelist():
            root = ET.fromstring(book.read("xl/sharedStrings.xml"))
            shared = ["".join(t.text or "" for t in si.iter(f"{{{ns['m']}}}t")) for si in root.findall("m:si", ns)]
        sheets = [n for n in book.namelist() if re.fullmatch(r"xl/worksheets/sheet\d+\.xml", n)]
        first = min(sheets, key=lambda n: int(re.search(r"\d+", n.rsplit("/", 1)[1]).group()))
        rows = []
        for row in ET.fromstring(book.read(first)).iter(f"{{{ns['m']}}}row"):
            cells = {}
            for cell in row.findall("m:c", ns):
                letters = re.match(r"[A-Z]+", cell.get("r", "A")).group()
                column = 0
                for char in letters:
                    column = column * 26 + ord(char) - 64
                value = cell.find("m:v", ns)
                if cell.get("t") == "s" and value is not None:
                    cells[column] = shared[int(value.text)]
                elif cell.get("t") == "inlineStr":
                    cells[column] = "".join(t.text or "" for t in cell.iter(f"{{{ns['m']}}}t"))
                elif value is not None:
                    cells[column] = value.text or ""
            rows.append([cells.get(i, "") for i in range(1, max(cells, default=0) + 1)])
        return rows


def read_question_file(data):
    if data[:4] == b"\xd0\xcf\x11\xe0":
        raise ValueError("This is an old .xls file. In Excel, use Save As and choose .xlsx.")
    if data[:2] == b"PK":
        try:
            return read_xlsx(data)
        except (zipfile.BadZipFile, ET.ParseError, KeyError, ValueError, IndexError):
            raise ValueError("This Excel file could not be read. Please save it again as .xlsx.")
    try:
        return list(csv.reader(io.StringIO(data.decode("utf-8-sig"), newline="")))
    except UnicodeDecodeError:
        raise ValueError('This CSV file is not UTF-8. In Excel, use Save As and choose "CSV UTF-8", or upload .xlsx.')


def parse_question(values):
    """Takes question, A, B, C, D, correct letter and marks. Returns (question, list of problems)."""
    values = [str(v).strip() for v in values] + [""] * 7
    text, a, b, c, d, correct, marks = values[:7]
    correct = correct.lower()
    problems = []
    options = {"a": a, "b": b, "c": c, "d": d}
    if not text:
        problems.append("question is empty")
    if sum(1 for v in options.values() if v) < 2:
        problems.append("needs at least 2 options")
    if correct not in LETTERS or not correct:
        problems.append("correct answer must be A, B, C or D")
    elif not options[correct]:
        problems.append(f"correct answer is {correct.upper()} but option {correct.upper()} is empty")
    try:
        marks_value = float(marks) if marks else 1.0
        if not 0 < marks_value <= 1000:
            problems.append("marks must be more than 0")
    except ValueError:
        marks_value = 1.0
        problems.append("marks must be a number")
    return (text, a, b, c, d, correct, marks_value), problems


def import_questions(db, data):
    """Saves every question or none of them. Returns (added, skipped duplicates, problems)."""
    try:
        rows = read_question_file(data)
    except ValueError as error:
        return 0, 0, [str(error)]
    questions, problems = [], []
    for number, row in enumerate(rows, 1):
        if not any(str(v).strip() for v in row):
            continue
        if number == 1 and str(row[0]).strip().lower() in ("question", "questions"):
            continue
        question, row_problems = parse_question(row)
        if row_problems:
            problems.append(f"Row {number}: " + ", ".join(row_problems))
        questions.append(question)
    if problems:
        return 0, 0, problems
    if not questions:
        return 0, 0, ["No questions were found in this file."]
    seen = {r["text"] for r in db.execute("SELECT text FROM questions")}
    added = skipped = 0
    for question in questions:
        if question[0] in seen:
            skipped += 1
            continue
        seen.add(question[0])
        db.execute("INSERT INTO questions (text, a, b, c, d, correct, marks) VALUES (?,?,?,?,?,?,?)", question)
        added += 1
    db.commit()
    return added, skipped, []


# ---------------------------------------------------------------- pages

CSS = """
:root{--bg:#f4f5f7;--card:#fff;--ink:#1c2430;--soft:#5d6878;--line:#d9dde3;--main:#1d5fbf;--bad:#b3261e;--good:#17713c;--tint:#e9f0fb}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--ink);font:16px/1.5 system-ui,-apple-system,"Segoe UI","Nirmala UI","Noto Sans Devanagari",sans-serif}
main{max-width:860px;margin:0 auto;padding:20px 16px 60px}
h1{font-size:24px;margin:0 0 4px}h2{font-size:18px;margin:0 0 12px}
.card{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:18px;margin:0 0 16px}
.narrow{max-width:420px;margin:8vh auto 0}
.soft{color:var(--soft)}.small{font-size:14px}
label{display:block;margin:0 0 12px;font-weight:600}
input[type=text],input[type=password],input[type=number],textarea,select{display:block;width:100%;margin-top:4px;padding:10px;font:inherit;border:1px solid var(--line);border-radius:6px;background:#fff;color:inherit}
button,.btn{display:inline-block;padding:10px 16px;font:inherit;font-weight:600;border:1px solid var(--main);border-radius:6px;background:var(--main);color:#fff;cursor:pointer;text-decoration:none}
button.plain,.btn.plain{background:#fff;color:var(--main)}
button.bad{background:var(--bad);border-color:var(--bad)}
button.mini,.btn.mini{padding:4px 10px;font-size:14px}
button.plain.bad{background:#fff;color:var(--bad)}
form.inline{display:inline}
.row{display:flex;flex-wrap:wrap;gap:10px;align-items:center}
.grid{display:grid;grid-template-columns:1fr 1fr;gap:0 12px}
.msg{padding:10px 14px;border-radius:6px;margin:0 0 16px;background:var(--tint)}
.err{padding:10px 14px;border-radius:6px;margin:0 0 16px;background:#fbeaea;color:var(--bad)}
.err ul{margin:4px 0 0;padding-left:20px}
.link{font-size:22px;font-weight:700;overflow-wrap:anywhere}
.scroll{overflow-x:auto}
table{width:100%;border-collapse:collapse;font-size:15px}
th,td{padding:8px;border-bottom:1px solid var(--line);text-align:left}
th{color:var(--soft);font-size:14px}
ol.qs{margin:0;padding:0;list-style:none}
ol.qs li{padding:12px 0;border-bottom:1px solid var(--line)}
ol.qs li:last-child{border-bottom:0}
.right{color:var(--good);font-weight:700}
.q{white-space:pre-wrap;font-weight:600;margin-bottom:8px}
.opt{display:flex;gap:10px;align-items:flex-start;font-weight:400;padding:10px;margin:6px 0;border:1px solid var(--line);border-radius:6px;cursor:pointer}
.opt:has(input:checked){border-color:var(--main);background:var(--tint)}
.opt input{margin-top:5px;flex:none}
.bar{position:sticky;top:0;z-index:2;display:flex;justify-content:space-between;gap:12px;align-items:center;background:var(--ink);color:#fff;padding:10px 16px;font-weight:600}
#timer{font-size:20px;font-variant-numeric:tabular-nums}
#net{display:none;background:var(--bad);color:#fff;padding:8px 16px;text-align:center}
.big{font-size:40px;font-weight:700}
@media (max-width:560px){.grid{grid-template-columns:1fr}}
"""

STUDENT_JS = """
(function () {
  var view = document.body.dataset.view, pending = {}, left = Number(document.body.dataset.left || 0);
  var net = document.getElementById('net'), timer = document.getElementById('timer');
  var paper = document.getElementById('paper'), count = document.getElementById('count');

  function showNet(bad) { if (net) net.style.display = bad ? 'block' : 'none'; }
  function send(q, c) {
    return fetch('/answer', {method: 'POST', headers: {'Content-Type': 'application/x-www-form-urlencoded'},
      body: 'q=' + encodeURIComponent(q) + '&c=' + encodeURIComponent(c)}).then(function (r) {
        if (r.status === 409) { location.reload(); return; }
        if (!r.ok) throw new Error('save');
        if (pending[q] === c) delete pending[q];
      });
  }
  function flush() {
    return Promise.all(Object.keys(pending).map(function (q) { return send(q, pending[q]); }))
      .then(function () { showNet(false); return true; })
      .catch(function () { showNet(true); return false; });
  }
  function poll() {
    flush();
    fetch('/status').then(function (r) { return r.json(); }).then(function (s) {
      if (s.view !== view) { location.reload(); return; }
      left = s.left;
      if (!Object.keys(pending).length) showNet(false);
    }).catch(function () { showNet(true); });
  }
  if (paper) {
    var total = paper.querySelectorAll('.question').length;
    var answered = function () { return paper.querySelectorAll('input:checked').length; };
    paper.addEventListener('change', function (ev) {
      pending[ev.target.dataset.q] = ev.target.value;
      count.textContent = answered() + ' / ' + total;
      flush();
    });
    paper.addEventListener('submit', function (ev) {
      ev.preventDefault();
      flush().then(function (saved) {
        if (!saved) { alert('Answers are not saved yet. Check the Wi-Fi and try again.\\nउत्तर सहेजे नहीं गए। वाई-फाई जाँचें और फिर प्रयास करें।'); return; }
        var blank = total - answered();
        var ask = (blank ? blank + ' question(s) not answered. / ' + blank + ' प्रश्न बाकी हैं।\\n' : '') + 'Submit your exam? You cannot change it afterwards.\\nक्या आप परीक्षा जमा करना चाहते हैं?';
        if (confirm(ask)) paper.submit();
      });
    });
  }
  if (timer) setInterval(function () {
    left = Math.max(0, left - 1);
    var m = Math.floor(left / 60), s = left % 60;
    timer.textContent = m + ':' + (s < 10 ? '0' : '') + s;
    if (left === 0) poll();
  }, 1000);
  if (view !== 'join') setInterval(poll, 5000);
})();
"""

TEACHER_JS = """
(function () {
  var box = document.getElementById('students');
  document.addEventListener('submit', function (ev) {
    var ask = ev.target.dataset.confirm;
    if (ask && !confirm(ask)) ev.preventDefault();
  });
  if (box) setInterval(function () {
    fetch('/teacher/students').then(function (r) {
      if (r.redirected) { location.reload(); return null; }
      return r.text();
    }).then(function (text) {
      if (text === null) return;
      box.innerHTML = text;
      var now = box.querySelector('[data-state]');
      if (now && now.dataset.state !== box.dataset.state) location.reload();
    }).catch(function () {});
  }, 5000);
})();
"""


def page(title, body, attrs="", script=""):
    js = f"<script>{script}</script>" if script else ""
    return (f'<!doctype html><html lang="en"><head><meta charset="utf-8">'
            f'<meta name="viewport" content="width=device-width,initial-scale=1"><title>{e(title)}</title>'
            f"<style>{CSS}</style></head><body{attrs}>{body}{js}</body></html>")


def student_link():
    """The address students type. Prefers the school network over VPN or virtual adapters."""
    try:
        addresses = {info[4][0] for info in socket.getaddrinfo(socket.gethostname(), None, socket.AF_INET)}
    except OSError:
        addresses = set()
    for prefix in ("192.168.", "10.", "172."):
        for address in sorted(addresses):
            if address.startswith(prefix):
                return f"http://{address}:{PORT}"
    return f"http://{sorted(addresses)[0] if addresses else '127.0.0.1'}:{PORT}"


def join_page(db, error=""):
    title = get_setting(db, "title", "Exam")
    if exam_state(db) == "ended":
        body = f'<main><div class="card narrow"><h1>{e(title)}</h1><p>This exam has ended.<br>यह परीक्षा समाप्त हो गई है।</p></div></main>'
        return page(title, body)
    problem = f'<div class="err">{e(error)}</div>' if error else ""
    body = f"""<main><form class="card narrow" method="post" action="/join">
        <h1>{e(title)}</h1><p class="soft">Enter your details to begin. / शुरू करने के लिए अपना विवरण भरें।</p>{problem}
        <label>Name / नाम<input type="text" name="name" maxlength="100" required autocomplete="off"></label>
        <label>Roll number / रोल नंबर<input type="text" name="roll" maxlength="30" required autocomplete="off"></label>
        <button>Continue / आगे बढ़ें</button></form></main>"""
    return page(title, body)


def waiting_page(db, student):
    title = get_setting(db, "title", "Exam")
    body = f"""<main><div class="card narrow"><h1>{e(title)}</h1>
        <p><b>{e(student['name'])}</b> &middot; Roll {e(student['roll'])}</p>
        <p>Please wait. The exam starts when your teacher begins it.<br>कृपया प्रतीक्षा करें। शिक्षक के शुरू करने पर परीक्षा शुरू होगी।</p>
        <p class="soft small">This page opens the exam by itself. / यह पृष्ठ अपने आप खुल जाएगा।</p></div></main>"""
    return page(title, body, ' data-view="waiting"', STUDENT_JS)


def exam_page(db, student):
    title = get_setting(db, "title", "Exam")
    chosen = {r["question_id"]: r["choice"] for r in
              db.execute("SELECT question_id, choice FROM answers WHERE student_id=?", (student["id"],))}
    questions = db.execute("SELECT * FROM questions ORDER BY id").fetchall()
    parts = []
    for number, q in enumerate(questions, 1):
        options = "".join(
            f'<label class="opt"><input type="radio" name="q{q["id"]}" value="{letter}" data-q="{q["id"]}"'
            f'{" checked" if chosen.get(q["id"]) == letter else ""}><span>{e(q[letter])}</span></label>'
            for letter in LETTERS if q[letter])
        parts.append(f'<div class="card question"><div class="q">{number}. {e(q["text"])}</div>'
                     f'<div class="soft small">{fmt(q["marks"])} mark(s) / अंक</div>{options}</div>')
    left = seconds_left(db)
    body = f"""<div class="bar"><span>{e(student['name'])} &middot; {e(student['roll'])}</span>
        <span id="count">{len(chosen)} / {len(questions)}</span><span id="timer">{left // 60}:{left % 60:02d}</span></div>
        <div id="net">Connection lost. Keep working, answers save when it returns. / कनेक्शन टूट गया है।</div>
        <main><h1 style="margin-bottom:14px">{e(title)}</h1>
        <form id="paper" method="post" action="/submit">{''.join(parts)}
        <button>Submit exam / परीक्षा जमा करें</button></form></main>"""
    return page(title, body, f' data-view="exam" data-left="{left}"', STUDENT_JS)


def done_page(db, student, view):
    title = get_setting(db, "title", "Exam")
    score = ""
    if view == "results":
        row = next(r for r in student_rows(db) if r["id"] == student["id"])
        score = f'<p class="big">{fmt(row["score"])} / {fmt(question_totals(db)[1])}</p>'
    body = f"""<main><div class="card narrow"><h1>{e(title)}</h1>
        <p><b>{e(student['name'])}</b> &middot; Roll {e(student['roll'])}</p>{score}
        <p>Your exam is finished. You can close this page.<br>आपकी परीक्षा पूरी हो गई है। आप यह पृष्ठ बंद कर सकते हैं।</p></div></main>"""
    return page(title, body, f' data-view="{view}"', STUDENT_JS)


def students_fragment(db):
    state = exam_state(db)
    rows = student_rows(db)
    count, total = question_totals(db)
    parts = []
    for r in rows:
        actions = ""
        if state != "ended" and (r["token"] or r["submitted"]):
            actions += (f'<form class="inline" method="post" action="/teacher/student/reentry" '
                        f'data-confirm="Let {e(r["name"])} enter again? Their saved answers are kept.">'
                        f'<input type="hidden" name="id" value="{r["id"]}"><button class="plain mini">Let back in</button></form> ')
        actions += (f'<form class="inline" method="post" action="/teacher/student/remove" '
                    f'data-confirm="Remove {e(r["name"])} and their answers?">'
                    f'<input type="hidden" name="id" value="{r["id"]}"><button class="plain bad mini">Remove</button></form>')
        parts.append(f'<tr><td>{e(r["roll"])}</td><td>{e(r["name"])}</td><td>{r["answered"]} / {count}</td>'
                     f'<td>{fmt(r["score"])} / {fmt(total)}</td><td>{student_status(r, state)}</td><td>{actions}</td></tr>')
    table = ('<div class="scroll"><table><thead><tr><th>Roll</th><th>Name</th><th>Answered</th><th>Score</th>'
             f'<th>Status</th><th></th></tr></thead><tbody>{"".join(parts)}</tbody></table></div>'
             if rows else '<p class="soft">No students have joined yet.</p>')
    return f'<div data-state="{state}"><p class="soft small">{len(rows)} student(s). Updates by itself.</p>{table}</div>'


def winners_card(db, state):
    """The top 5 ranks once the exam ends. Equal scores share a rank, and a score of 0 does not win."""
    if state != "ended":
        return ""
    rows = sorted((r for r in student_rows(db) if r["score"] > 0), key=lambda r: -r["score"])
    total = question_totals(db)[1]
    ordinals = {1: "1st", 2: "2nd", 3: "3rd", 4: "4th", 5: "5th"}
    parts, rank, previous = [], 0, None
    for place, r in enumerate(rows, 1):
        if r["score"] != previous:
            rank, previous = place, r["score"]
        if rank > 5:
            break
        parts.append(f'<tr><td><b>{ordinals[rank]}</b></td><td>{e(r["name"])}</td><td>{e(r["roll"])}</td>'
                     f'<td>{fmt(r["score"])} / {fmt(total)}</td></tr>')
    if not parts:
        return '<section class="card"><h2>Winners</h2><p class="soft">No student scored any marks.</p></section>'
    return ('<section class="card"><h2>Winners</h2><div class="scroll"><table><thead><tr><th>Rank</th><th>Name</th>'
            f'<th>Roll</th><th>Score</th></tr></thead><tbody>{"".join(parts)}</tbody></table></div>'
            '<p class="soft small" style="margin:10px 0 0">Students with the same score share a rank.</p></section>')


def control_card(db, state):
    count, total = question_totals(db)
    title = get_setting(db, "title", "")
    if state == "setup":
        minutes = get_setting(db, "minutes", "30")
        return f"""<section class="card"><h2>Exam</h2><form method="post" action="/teacher/start"
            data-confirm="Start the exam now? The timer begins for every student.">
            <div class="grid"><label>Exam name<input type="text" name="title" maxlength="120" value="{e(title)}" required></label>
            <label>Time in minutes<input type="number" name="minutes" min="1" max="600" value="{e(minutes)}" required></label></div>
            <p class="soft small">{count} question(s), {fmt(total)} mark(s) in total.</p>
            <button{'' if count else ' disabled'}>Start exam</button></form></section>"""
    if state == "running":
        return f"""<section class="card"><h2>{e(title)}: running</h2>
            <p>Time left when this page opened: <b>{seconds_left(db) // 60} min</b></p><div class="row">
            <form method="post" action="/teacher/addtime"><button class="plain">Add 5 minutes</button></form>
            <form method="post" action="/teacher/stop" data-confirm="Stop the exam now for every student?"><button class="bad">Stop exam</button></form>
            </div></section>"""
    show = get_setting(db, "show_scores", "0") == "1"
    return f"""<section class="card"><h2>{e(title)}: finished</h2><div class="row">
        <a class="btn" href="/teacher/results.csv">Download results</a>
        <form method="post" action="/teacher/scores"><button class="plain">{'Hide scores from students' if show else 'Show scores to students'}</button></form>
        <form method="post" action="/teacher/new" data-confirm="Start a new exam? All students and their results will be deleted. Download the results first. Questions are kept.">
        <button class="plain bad">New exam</button></form></div></section>"""


def questions_card(db, state, edit_id=None):
    questions = db.execute("SELECT * FROM questions ORDER BY id").fetchall()
    locked = state != "setup"
    forms = ""
    if locked:
        forms = '<p class="soft">Questions can be changed after you choose New exam.</p>'
    else:
        q = next((x for x in questions if x["id"] == edit_id), None)
        value = lambda key, default="": e(q[key]) if q else default
        marks = fmt(q["marks"]) if q else "1"
        choices = "".join(f'<option value="{l}"{" selected" if q and q["correct"] == l else ""}>{l.upper()}</option>' for l in LETTERS)
        forms = f"""<form method="post" action="/teacher/upload" enctype="multipart/form-data" style="margin-bottom:16px">
            <label>Upload an Excel (.xlsx) or CSV file<input type="file" name="file" accept=".xlsx,.csv" required style="display:block;margin-top:6px;font-weight:400"></label>
            <div class="row"><button>Upload questions</button><a class="btn plain" href="/teacher/template.csv">Download sample file</a></div></form>
            <details{' open' if q else ''}><summary style="cursor:pointer;font-weight:600;margin-bottom:10px">{'Edit question' if q else 'Or type one question'}</summary>
            <form method="post" action="/teacher/question"><input type="hidden" name="id" value="{q['id'] if q else ''}">
            <label>Question<textarea name="text" rows="2" required>{value('text')}</textarea></label>
            <div class="grid"><label>Option A<input type="text" name="a" value="{value('a')}" required></label>
            <label>Option B<input type="text" name="b" value="{value('b')}" required></label>
            <label>Option C<input type="text" name="c" value="{value('c')}"></label>
            <label>Option D<input type="text" name="d" value="{value('d')}"></label>
            <label>Correct answer<select name="correct">{choices}</select></label>
            <label>Marks<input type="number" name="marks" min="0.25" step="0.25" value="{marks}" required></label></div>
            <button>{'Save changes' if q else 'Add question'}</button></form></details>"""
    items = []
    for number, q in enumerate(questions, 1):
        options = " &nbsp; ".join(
            f'<span class="{"right" if q["correct"] == l else ""}">{l.upper()}. {e(q[l])}</span>' for l in LETTERS if q[l])
        actions = "" if locked else (
            f'<div class="row" style="margin-top:6px"><a class="btn plain mini" href="/teacher?edit={q["id"]}">Edit</a>'
            f'<form method="post" action="/teacher/question/delete"><input type="hidden" name="id" value="{q["id"]}">'
            '<button class="plain bad mini">Delete</button></form></div>')
        items.append(f'<li><div class="q" style="margin:0">{number}. {e(q["text"])}</div>'
                     f'<div class="small">{options} <span class="soft">&middot; {fmt(q["marks"])} mark(s)</span></div>{actions}</li>')
    clear = "" if locked or not questions else (
        '<form method="post" action="/teacher/questions/clear" data-confirm="Delete all questions?" style="margin-top:12px">'
        '<button class="plain bad mini">Delete all questions</button></form>')
    listing = f'<ol class="qs" style="margin-top:16px">{"".join(items)}</ol>{clear}' if items else ""
    return f'<section class="card"><h2>Questions ({len(questions)})</h2>{forms}{listing}</section>'


def teacher_page(db, msg="", errors=None, edit_id=None):
    state = exam_state(db)
    notes = f'<div class="msg">{e(msg)}</div>' if msg else ""
    if errors:
        items = "".join(f"<li>{e(x)}</li>" for x in errors[:30])
        more = f"<li>and {len(errors) - 30} more</li>" if len(errors) > 30 else ""
        notes += f'<div class="err"><b>Nothing was saved. Fix these and upload again:</b><ul>{items}{more}</ul></div>'
    body = f"""<main><div class="row" style="justify-content:space-between;margin-bottom:16px"><h1>Teacher</h1>
        <form method="post" action="/teacher/logout"><button class="plain mini">Log out</button></form></div>{notes}
        <section class="card"><span class="soft small">Students open this address on the same Wi-Fi:</span>
        <div class="link">{e(student_link())}</div></section>
        {control_card(db, state)}
        {winners_card(db, state)}
        <section class="card"><h2>Students</h2><div id="students" data-state="{state}">{students_fragment(db)}</div></section>
        {questions_card(db, state, edit_id)}</main>"""
    return page("Teacher", body, script=TEACHER_JS)


def login_page(error=""):
    problem = f'<div class="err">{e(error)}</div>' if error else ""
    body = f"""<main><form class="card narrow" method="post" action="/teacher/login"><h1>Teacher login</h1>{problem}
        <label>Password<input type="password" name="password" required autofocus></label><button>Log in</button>
        <p class="soft small">Forgot it? On the exam computer, run RESET_PASSWORD.bat.</p></form></main>"""
    return page("Teacher login", body)


def setup_page(local, error=""):
    if not local:
        body = '<main><div class="card narrow"><h1>Not set up yet</h1><p>Open this page on the exam computer to set the teacher password.</p></div></main>'
        return page("Teacher", body)
    problem = f'<div class="err">{e(error)}</div>' if error else ""
    body = f"""<main><form class="card narrow" method="post" action="/teacher/setup"><h1>Set teacher password</h1>
        <p class="soft">Only you will use this password. Students never need it.</p>{problem}
        <label>New password<input type="password" name="password" minlength="6" required autofocus></label>
        <button>Save password</button></form></main>"""
    return page("Teacher", body)


# ---------------------------------------------------------------- server

class LoginLimiter:
    """Blocks an address for a minute after 5 wrong passwords."""
    def __init__(self):
        self.lock = threading.Lock()
        self.fails = {}

    def blocked(self, ip):
        with self.lock:
            count, until = self.fails.get(ip, (0, 0))
            return count >= 5 and time.time() < until

    def record(self, ip, ok):
        with self.lock:
            if ok:
                self.fails.pop(ip, None)
            else:
                count, until = self.fails.get(ip, (0, 0))
                self.fails[ip] = (1 if time.time() >= until else count + 1, time.time() + 60)


limiter = LoginLimiter()


class Handler(BaseHTTPRequestHandler):
    protocol_version = "HTTP/1.0"
    timeout = 30

    def log_message(self, *args):
        pass

    # ------------ helpers
    def send(self, body, status=200, content_type="text/html; charset=utf-8", headers=()):
        data = body.encode("utf-8") if isinstance(body, str) else body
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", "no-store")
        for key, value in headers:
            self.send_header(key, value)
        self.end_headers()
        self.wfile.write(data)

    def redirect(self, where, headers=()):
        self.send("", 303, headers=[("Location", where), *headers])

    def to_teacher(self, msg=""):
        self.redirect("/teacher" + (f"?msg={quote(msg)}" if msg else ""))

    def cookie(self, name):
        morsel = SimpleCookie(self.headers.get("Cookie", "")).get(name)
        return morsel.value if morsel else None

    def is_local(self):
        return self.client_address[0] in ("127.0.0.1", "::1")

    def is_teacher(self):
        token = self.cookie("teacher")
        if not token:
            return False
        row = self.db.execute("SELECT created FROM sessions WHERE token=?", (token,)).fetchone()
        return bool(row) and time.time() - row["created"] < TEACHER_SESSION_SECONDS

    def current_student(self):
        token = self.cookie("student")
        if not token:
            return None
        return self.db.execute("SELECT * FROM students WHERE token=?", (token,)).fetchone()

    def student_view(self, student):
        state = exam_state(self.db)
        if student is None:
            return "join"
        if state == "ended" or student["submitted"]:
            return "results" if state == "ended" and get_setting(self.db, "show_scores", "0") == "1" else "done"
        return "waiting" if state == "setup" else "exam"

    def read_form(self):
        length = int(self.headers.get("Content-Length") or 0)
        if length > MAX_BODY:
            self.close_connection = True
            raise OverflowError
        raw = self.rfile.read(length)
        if self.headers.get_content_type() == "multipart/form-data":
            head = f"Content-Type: {self.headers['Content-Type']}\r\n\r\n".encode()
            parts = {}
            for part in BytesParser().parsebytes(head + raw).walk():
                name = part.get_param("name", header="content-disposition")
                if name:
                    parts[name] = part.get_payload(decode=True)
            return parts
        return {k: v[0] for k, v in parse_qs(raw.decode("utf-8", "replace"), keep_blank_values=True).items()}

    # ------------ routing
    def do_GET(self):
        self.dispatch(self.handle_get)

    def do_POST(self):
        self.dispatch(self.handle_post)

    def dispatch(self, handler):
        self.db = connect()
        parts = urlsplit(self.path)
        self.route, self.query = parts.path, parse_qs(parts.query)
        try:
            handler()
            self.db.commit()
        except OverflowError:
            self.send("That file is too big. The limit is 10 MB.", 413, "text/plain; charset=utf-8")
        except (BrokenPipeError, ConnectionResetError, ConnectionAbortedError):
            pass
        except Exception:
            traceback.print_exc()
            try:
                self.send("Something went wrong. Please try again.", 500, "text/plain; charset=utf-8")
            except OSError:
                pass
        finally:
            self.db.close()

    def handle_get(self):
        if self.route == "/":
            student = self.current_student()
            view = self.student_view(student)
            if view == "join":
                self.send(join_page(self.db))
            elif view == "waiting":
                self.send(waiting_page(self.db, student))
            elif view == "exam":
                self.send(exam_page(self.db, student))
            else:
                self.send(done_page(self.db, student, view))
        elif self.route == "/status":
            view = self.student_view(self.current_student())
            self.send(json.dumps({"view": view, "left": seconds_left(self.db)}), content_type="application/json")
        elif self.route == "/teacher":
            if not get_setting(self.db, "password"):
                self.send(setup_page(self.is_local()))
            elif not self.is_teacher():
                self.send(login_page())
            else:
                edit = self.query.get("edit", [""])[0]
                self.send(teacher_page(self.db, self.query.get("msg", [""])[0], edit_id=int(edit) if edit.isdigit() else None))
        elif not self.route.startswith("/teacher/"):
            self.send("Not found", 404, "text/plain; charset=utf-8")
        elif not self.is_teacher():
            self.redirect("/teacher")
        elif self.route == "/teacher/students":
            self.send(students_fragment(self.db))
        elif self.route == "/teacher/template.csv":
            self.send_csv("questions_sample.csv", [
                ["Question", "A", "B", "C", "D", "Correct", "Marks"],
                ["What is 2 + 2?", "3", "4", "5", "6", "B", "1"],
                ["भारत की राजधानी क्या है?", "मुंबई", "नई दिल्ली", "कोलकाता", "चेन्नई", "B", "2"],
                ["The sun rises in the east.", "True", "False", "", "", "A", "1"]])
        elif self.route == "/teacher/results.csv":
            self.send_results()
        else:
            self.send("Not found", 404, "text/plain; charset=utf-8")

    def send_csv(self, filename, rows):
        out = io.StringIO()
        csv.writer(out).writerows(rows)
        self.send("﻿" + out.getvalue(), content_type="text/csv; charset=utf-8",
                  headers=[("Content-Disposition", f'attachment; filename="{filename}"')])

    def send_results(self):
        state = exam_state(self.db)
        total = question_totals(self.db)[1]
        # A leading ' stops Excel running a typed name such as =SUM(...) as a formula.
        safe = lambda text: "'" + text if text[:1] in ("=", "+", "-", "@", "\t", "\r") else text
        rows = [["Roll", "Name", "Score", "Out of", "Answered", "Status"]]
        for r in student_rows(self.db):
            rows.append([safe(r["roll"]), safe(r["name"]), fmt(r["score"]), fmt(total), r["answered"], student_status(r, state)])
        self.send_csv("results.csv", rows)

    def handle_post(self):
        form = self.read_form()
        if self.route.startswith("/teacher"):
            return self.teacher_post(form)
        if self.route == "/join":
            return self.join(form)
        student = self.current_student()
        if self.route == "/answer":
            if self.student_view(student) != "exam":
                return self.send("closed", 409, "text/plain")
            question, choice = form.get("q", ""), form.get("c", "")
            row = self.db.execute("SELECT * FROM questions WHERE id=?", (question,)).fetchone()
            if not row or choice not in LETTERS or not choice or not row[choice]:
                return self.send("bad answer", 400, "text/plain")
            self.db.execute("INSERT OR REPLACE INTO answers (student_id, question_id, choice) VALUES (?,?,?)",
                            (student["id"], row["id"], choice))
            self.db.commit()
            return self.send("ok", content_type="text/plain")
        if self.route == "/submit":
            if student:
                self.db.execute("UPDATE students SET submitted=1 WHERE id=?", (student["id"],))
                self.db.commit()
            return self.redirect("/")
        self.send("Not found", 404, "text/plain; charset=utf-8")

    def join(self, form):
        name = " ".join(form.get("name", "").split())[:100]
        roll = " ".join(form.get("roll", "").split())[:30]
        if not name or not roll:
            return self.send(join_page(self.db, "Please enter your name and roll number. / कृपया नाम और रोल नंबर भरें।"))
        if exam_state(self.db) == "ended":
            return self.send(join_page(self.db))
        token = secrets.token_urlsafe(24)
        key = roll.casefold()
        taken = "This roll number has already entered. Ask your teacher to let you back in. / यह रोल नंबर पहले से दर्ज है, शिक्षक से कहें।"
        existing = self.db.execute("SELECT * FROM students WHERE roll_key=?", (key,)).fetchone()
        if existing and existing["token"]:
            return self.send(join_page(self.db, taken))
        try:
            if existing:
                self.db.execute("UPDATE students SET token=? WHERE id=?", (token, existing["id"]))
            else:
                self.db.execute("INSERT INTO students (name, roll, roll_key, token) VALUES (?,?,?,?)", (name, roll, key, token))
            self.db.commit()
        except sqlite3.IntegrityError:
            return self.send(join_page(self.db, taken))
        self.redirect("/", [("Set-Cookie", f"student={token}; Path=/; Max-Age=86400; HttpOnly; SameSite=Lax")])

    def start_session(self):
        token = secrets.token_urlsafe(32)
        self.db.execute("DELETE FROM sessions WHERE created < ?", (time.time() - TEACHER_SESSION_SECONDS,))
        self.db.execute("INSERT INTO sessions (token, created) VALUES (?, ?)", (token, time.time()))
        self.db.commit()
        self.redirect("/teacher", [("Set-Cookie", f"teacher={token}; Path=/; Max-Age={TEACHER_SESSION_SECONDS}; HttpOnly; SameSite=Lax")])

    def teacher_post(self, form):
        db = self.db
        ip = self.client_address[0]
        if self.route == "/teacher/setup":
            if get_setting(db, "password") or not self.is_local():
                return self.redirect("/teacher")
            password = form.get("password", "")
            if len(password) < 6:
                return self.send(setup_page(True, "Use at least 6 characters."))
            set_setting(db, "password", hash_password(password))
            return self.start_session()
        if self.route == "/teacher/login":
            stored = get_setting(db, "password")
            if not stored:
                return self.redirect("/teacher")
            if limiter.blocked(ip):
                return self.send(login_page("Too many wrong tries. Wait one minute."))
            ok = check_password(form.get("password", ""), stored)
            limiter.record(ip, ok)
            return self.start_session() if ok else self.send(login_page("Wrong password."))
        if not self.is_teacher():
            return self.redirect("/teacher")

        state = exam_state(db)
        if self.route == "/teacher/logout":
            db.execute("DELETE FROM sessions WHERE token=?", (self.cookie("teacher"),))
            db.commit()
            return self.redirect("/teacher", [("Set-Cookie", "teacher=; Path=/; Max-Age=0")])

        if self.route in ("/teacher/upload", "/teacher/question", "/teacher/question/delete", "/teacher/questions/clear") and state != "setup":
            return self.to_teacher("Questions cannot be changed during or after an exam. Choose New exam first.")
        if self.route == "/teacher/upload":
            added, skipped, problems = import_questions(db, form.get("file") or b"")
            if problems:
                return self.send(teacher_page(db, errors=problems))
            return self.to_teacher(f"{added} question(s) added." + (f" {skipped} already there, skipped." if skipped else ""))
        if self.route == "/teacher/question":
            question, problems = parse_question([form.get(k, "") for k in ("text", "a", "b", "c", "d", "correct", "marks")])
            if problems:
                return self.to_teacher("Not saved: " + ", ".join(problems))
            if form.get("id", "").isdigit():
                db.execute("UPDATE questions SET text=?, a=?, b=?, c=?, d=?, correct=?, marks=? WHERE id=?", (*question, form["id"]))
                return self.to_teacher("Question saved.")
            db.execute("INSERT INTO questions (text, a, b, c, d, correct, marks) VALUES (?,?,?,?,?,?,?)", question)
            return self.to_teacher("Question added.")
        if self.route == "/teacher/question/delete":
            db.execute("DELETE FROM questions WHERE id=?", (form.get("id", ""),))
            return self.to_teacher("Question deleted.")
        if self.route == "/teacher/questions/clear":
            db.execute("DELETE FROM questions")
            return self.to_teacher("All questions deleted.")

        if self.route == "/teacher/start":
            minutes = form.get("minutes", "")
            if state != "setup" or not question_totals(db)[0]:
                return self.to_teacher("The exam cannot start yet.")
            if not minutes.isdigit() or not 1 <= int(minutes) <= 600:
                return self.to_teacher("Enter the time in minutes, from 1 to 600.")
            set_setting(db, "title", form.get("title", "").strip()[:120] or "Exam")
            set_setting(db, "minutes", minutes)
            set_setting(db, "ends_at", time.time() + int(minutes) * 60)
            set_setting(db, "state", "running")
            return self.to_teacher("The exam has started.")
        if self.route == "/teacher/stop":
            if state == "running":
                set_setting(db, "state", "ended")
            return self.to_teacher("The exam has stopped. All saved answers are marked.")
        if self.route == "/teacher/addtime":
            if state == "running":
                set_setting(db, "ends_at", float(get_setting(db, "ends_at")) + 300)
            return self.to_teacher("5 minutes added for every student.")
        if self.route == "/teacher/scores":
            set_setting(db, "show_scores", "0" if get_setting(db, "show_scores", "0") == "1" else "1")
            return self.to_teacher()
        if self.route == "/teacher/new":
            if state == "running":
                return self.to_teacher("Stop the exam first.")
            db.execute("DELETE FROM answers")
            db.execute("DELETE FROM students")
            set_setting(db, "state", "setup")
            set_setting(db, "show_scores", "0")
            return self.to_teacher("Ready for a new exam.")
        if self.route == "/teacher/student/reentry":
            # The next person to enter this roll number takes over the saved answers.
            db.execute("UPDATE students SET token=NULL, submitted=0 WHERE id=?", (form.get("id", ""),))
            return self.to_teacher("They can now enter again with the same roll number.")
        if self.route == "/teacher/student/remove":
            db.execute("DELETE FROM answers WHERE student_id=?", (form.get("id", ""),))
            db.execute("DELETE FROM students WHERE id=?", (form.get("id", ""),))
            return self.to_teacher("Student removed.")
        self.send("Not found", 404, "text/plain; charset=utf-8")


class Server(ThreadingHTTPServer):
    daemon_threads = True
    # On Windows, address reuse would let two copies listen on the same port without any error.
    allow_reuse_address = False
    request_queue_size = 128


def main():
    init_db()
    if "--reset-password" in sys.argv:
        db = connect()
        db.execute("DELETE FROM settings WHERE key='password'")
        db.execute("DELETE FROM sessions")
        db.commit()
        db.close()
        print("Teacher password cleared. Start the exam and open the Teacher page on this computer to set a new one.")
        return
    try:
        server = Server(("0.0.0.0", PORT), Handler)
    except OSError:
        print(f"Port {PORT} is already in use. Easy Exam may already be running: check for another open window.")
        input("Press Enter to close.")
        return
    print("\n  Easy Exam is running. Keep this window open during the exam.\n")
    print(f"  Teacher (this computer):  http://127.0.0.1:{PORT}/teacher")
    print(f"  Students (same Wi-Fi):    {student_link()}\n")
    if "--no-browser" not in sys.argv:
        webbrowser.open(f"http://127.0.0.1:{PORT}/teacher")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()
