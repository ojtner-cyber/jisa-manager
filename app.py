"""
지사 관리 시스템 v1.0
Flask + SQLite | 본사 전용 지사 관리 플랫폼
"""
from flask import Flask, render_template, request, jsonify, session, redirect, url_for, send_file
import sqlite3, json, os, csv, io
from datetime import date, datetime
from functools import wraps

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "jisa-manager-secret-2025")

# Render 환경: /opt/render/project/src 하위에 data 폴더 사용 (영구 디스크 마운트 시)
# 로컬: 스크립트와 같은 폴더
_base = os.environ.get("DATA_DIR", os.path.dirname(os.path.abspath(__file__)))
DB_FILE   = os.path.join(_base, "jisa.db")
GOAL_FILE = os.path.join(_base, "sales_goals.json")

REGIONS = ["서울","경기","인천","강원","충북","충남","대전","세종","경북","경남","대구","부산","울산","전북","전남","광주","제주"]

# ── DB 초기화 ─────────────────────────────────
def get_db():
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_db()
    conn.execute("""CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        email TEXT UNIQUE NOT NULL,
        password TEXT NOT NULL,
        name TEXT NOT NULL,
        role TEXT DEFAULT 'user',
        created_at TEXT DEFAULT CURRENT_TIMESTAMP)""")
    conn.execute("""CREATE TABLE IF NOT EXISTS branches (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL,
        region TEXT,
        manager TEXT,
        phone TEXT,
        email TEXT,
        address TEXT,
        status TEXT DEFAULT '운영중',
        contract_date TEXT,
        fee_rate REAL DEFAULT 0,
        note TEXT,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP)""")
    conn.execute("""CREATE TABLE IF NOT EXISTS sales (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        branch_id INTEGER NOT NULL,
        year INTEGER NOT NULL,
        month INTEGER NOT NULL,
        target INTEGER DEFAULT 0,
        actual INTEGER DEFAULT 0,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP,
        UNIQUE(branch_id, year, month),
        FOREIGN KEY(branch_id) REFERENCES branches(id))""")
    # 기본 계정 생성
    conn.execute("INSERT OR IGNORE INTO users(email,password,name,role) VALUES(?,?,?,?)",
        ("test@visang.com","visang123!","관리자","admin"))
    conn.execute("INSERT OR IGNORE INTO users(email,password,name,role) VALUES(?,?,?,?)",
        ("user@visang.com","visang123!","일반사용자","user"))
    # 샘플 지사
    sample = [
        ("서울 강남지사","서울","김철수","010-1234-5678","kangnam@visang.com","서울시 강남구","운영중","2023-01-01",5.0),
        ("경기 수원지사","경기","이영희","010-2345-6789","suwon@visang.com","경기도 수원시","운영중","2023-03-15",4.5),
        ("부산 해운대지사","부산","박민준","010-3456-7890","haeundae@visang.com","부산시 해운대구","운영중","2022-07-01",5.5),
        ("대구 중구지사","대구","최수진","010-4567-8901","daegu@visang.com","대구시 중구","운영중","2023-06-01",4.0),
        ("인천 부평지사","인천","정지훈","010-5678-9012","bupyeong@visang.com","인천시 부평구","일시중단","2022-12-01",3.5),
        ("광주 서구지사","광주","한소희","010-6789-0123","gwangju@visang.com","광주시 서구","운영중","2024-01-15",5.0),
    ]
    for s in sample:
        conn.execute("INSERT OR IGNORE INTO branches(name,region,manager,phone,email,address,status,contract_date,fee_rate) VALUES(?,?,?,?,?,?,?,?,?)", s)
    # 샘플 판매 데이터
    import random; random.seed(42)
    branches = conn.execute("SELECT id FROM branches").fetchall()
    y = datetime.now().year
    for b in branches:
        for m in range(1, 13):
            t = random.randint(800, 2000) * 10000
            a = int(t * random.uniform(0.6, 1.2))
            conn.execute("INSERT OR IGNORE INTO sales(branch_id,year,month,target,actual) VALUES(?,?,?,?,?)",
                (b[0], y, m, t, a))
    conn.commit(); conn.close()

# ── 인증 ──────────────────────────────────────
def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if "user" not in session:
            if request.path.startswith("/api/"):
                return jsonify({"error":"unauthorized"}), 401
            return redirect("/login")
        return f(*args, **kwargs)
    return decorated

def admin_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if session.get("user",{}).get("role") != "admin":
            return jsonify({"error":"forbidden"}), 403
        return f(*args, **kwargs)
    return login_required(decorated)

# ── 페이지 라우트 ──────────────────────────────
@app.route("/")
@login_required
def index():
    return redirect("/dashboard")

@app.route("/login", methods=["GET","POST"])
def login():
    if request.method == "POST":
        d = request.json
        conn = get_db()
        user = conn.execute("SELECT * FROM users WHERE email=? AND password=?",
            (d.get("email",""), d.get("password",""))).fetchone()
        conn.close()
        if user:
            session["user"] = dict(user)
            return jsonify({"ok":True, "role": user["role"]})
        return jsonify({"ok":False, "msg":"이메일 또는 비밀번호가 올바르지 않습니다."}), 401
    return render_template("index.html", regions=REGIONS)

@app.route("/logout")
def logout():
    session.clear()
    return redirect("/login")

# ── 대시보드 API ───────────────────────────────
@app.route("/api/dashboard")
@login_required
def api_dashboard():
    y = request.args.get("year", str(datetime.now().year))
    conn = get_db()
    total_branches = conn.execute("SELECT COUNT(*) FROM branches").fetchone()[0]
    active = conn.execute("SELECT COUNT(*) FROM branches WHERE status='운영중'").fetchone()[0]
    stats = conn.execute("""
        SELECT SUM(target) t, SUM(actual) a
        FROM sales WHERE year=?""", (y,)).fetchone()
    monthly = [dict(r) for r in conn.execute("""
        SELECT month, SUM(target) target, SUM(actual) actual
        FROM sales WHERE year=?
        GROUP BY month ORDER BY month""", (y,)).fetchall()]
    top5 = [dict(r) for r in conn.execute("""
        SELECT b.name, SUM(s.actual) total
        FROM sales s JOIN branches b ON s.branch_id=b.id
        WHERE s.year=?
        GROUP BY s.branch_id ORDER BY total DESC LIMIT 5""", (y,)).fetchall()]
    region_stats = [dict(r) for r in conn.execute("""
        SELECT b.region, SUM(s.actual) total
        FROM sales s JOIN branches b ON s.branch_id=b.id
        WHERE s.year=?
        GROUP BY b.region ORDER BY total DESC""", (y,)).fetchall()]
    conn.close()
    return jsonify({
        "total_branches": total_branches,
        "active_branches": active,
        "total_target": stats["t"] or 0,
        "total_actual": stats["a"] or 0,
        "achievement": round((stats["a"] or 0)/(stats["t"] or 1)*100,1),
        "monthly": monthly,
        "top5": top5,
        "region_stats": region_stats,
    })

# ── 지사 API ───────────────────────────────────
@app.route("/api/branches")
@login_required
def api_branches():
    region = request.args.get("region","")
    status = request.args.get("status","")
    q_str  = request.args.get("q","").strip()
    conn = get_db()
    q = "SELECT * FROM branches WHERE 1=1"
    params = []
    if region: q += " AND region=?"; params.append(region)
    if status: q += " AND status=?"; params.append(status)
    if q_str:  q += " AND (name LIKE ? OR manager LIKE ? OR address LIKE ?)"; params+=[f"%{q_str}%"]*3
    q += " ORDER BY name"
    rows = [dict(r) for r in conn.execute(q, params).fetchall()]
    # 올해 실적 붙이기
    y = datetime.now().year
    for row in rows:
        sales = conn.execute("SELECT SUM(target) t, SUM(actual) a FROM sales WHERE branch_id=? AND year=?",
            (row["id"], y)).fetchone()
        row["year_target"] = sales["t"] or 0
        row["year_actual"] = sales["a"] or 0
        row["achievement"] = round((sales["a"] or 0)/(sales["t"] or 1)*100,1)
    conn.close()
    return jsonify(rows)

@app.route("/api/branches", methods=["POST"])
@login_required
def api_branches_add():
    d = request.json
    conn = get_db()
    conn.execute("""INSERT INTO branches(name,region,manager,phone,email,address,status,contract_date,fee_rate,note)
        VALUES(?,?,?,?,?,?,?,?,?,?)""",
        (d["name"],d.get("region",""),d.get("manager",""),d.get("phone",""),
         d.get("email",""),d.get("address",""),d.get("status","운영중"),
         d.get("contract_date",""),float(d.get("fee_rate",0)),d.get("note","")))
    conn.commit(); conn.close()
    return jsonify({"ok":True})

@app.route("/api/branches/<int:bid>", methods=["GET"])
@login_required
def api_branch_get(bid):
    conn = get_db()
    row = conn.execute("SELECT * FROM branches WHERE id=?", (bid,)).fetchone()
    conn.close()
    return jsonify(dict(row) if row else {})

@app.route("/api/branches/<int:bid>", methods=["PUT"])
@login_required
def api_branches_update(bid):
    d = request.json
    conn = get_db()
    conn.execute("""UPDATE branches SET name=?,region=?,manager=?,phone=?,email=?,
        address=?,status=?,contract_date=?,fee_rate=?,note=? WHERE id=?""",
        (d["name"],d.get("region",""),d.get("manager",""),d.get("phone",""),
         d.get("email",""),d.get("address",""),d.get("status","운영중"),
         d.get("contract_date",""),float(d.get("fee_rate",0)),d.get("note",""),bid))
    conn.commit(); conn.close()
    return jsonify({"ok":True})

@app.route("/api/branches/<int:bid>", methods=["DELETE"])
@login_required
def api_branches_delete(bid):
    conn = get_db()
    conn.execute("DELETE FROM sales WHERE branch_id=?", (bid,))
    conn.execute("DELETE FROM branches WHERE id=?", (bid,))
    conn.commit(); conn.close()
    return jsonify({"ok":True})

# ── 판매부수 API ───────────────────────────────
@app.route("/api/sales")
@login_required
def api_sales():
    bid  = request.args.get("branch_id")
    year = request.args.get("year", str(datetime.now().year))
    conn = get_db()
    if bid:
        rows = [dict(r) for r in conn.execute("""
            SELECT s.*, b.name branch_name FROM sales s
            JOIN branches b ON s.branch_id=b.id
            WHERE s.branch_id=? AND s.year=? ORDER BY s.month""", (bid, year)).fetchall()]
    else:
        rows = [dict(r) for r in conn.execute("""
            SELECT s.*, b.name branch_name, b.region FROM sales s
            JOIN branches b ON s.branch_id=b.id
            WHERE s.year=? ORDER BY b.name, s.month""", (year,)).fetchall()]
    conn.close()
    return jsonify(rows)

@app.route("/api/sales", methods=["POST"])
@login_required
def api_sales_save():
    d = request.json  # [{branch_id, year, month, target, actual}, ...]
    conn = get_db()
    for row in d:
        conn.execute("""INSERT INTO sales(branch_id,year,month,target,actual)
            VALUES(?,?,?,?,?)
            ON CONFLICT(branch_id,year,month) DO UPDATE SET target=excluded.target, actual=excluded.actual""",
            (row["branch_id"], row["year"], row["month"], row.get("target",0), row.get("actual",0)))
    conn.commit(); conn.close()
    return jsonify({"ok":True})

# ── 권한 관리 API ──────────────────────────────
@app.route("/api/users")
@login_required
def api_users():
    conn = get_db()
    rows = [dict(r) for r in conn.execute("SELECT id,email,name,role,created_at FROM users ORDER BY name").fetchall()]
    conn.close()
    return jsonify(rows)

@app.route("/api/users", methods=["POST"])
@login_required
def api_users_add():
    d = request.json
    conn = get_db()
    try:
        conn.execute("INSERT INTO users(email,password,name,role) VALUES(?,?,?,?)",
            (d["email"],d["password"],d["name"],d.get("role","user")))
        conn.commit()
    except sqlite3.IntegrityError:
        conn.close()
        return jsonify({"ok":False,"msg":"이미 존재하는 이메일입니다."}), 400
    conn.close()
    return jsonify({"ok":True})

@app.route("/api/users/<int:uid>", methods=["PUT"])
@login_required
def api_users_update(uid):
    d = request.json
    conn = get_db()
    if d.get("password"):
        conn.execute("UPDATE users SET name=?,role=?,password=? WHERE id=?",
            (d["name"],d["role"],d["password"],uid))
    else:
        conn.execute("UPDATE users SET name=?,role=? WHERE id=?", (d["name"],d["role"],uid))
    conn.commit(); conn.close()
    return jsonify({"ok":True})

@app.route("/api/users/<int:uid>", methods=["DELETE"])
@login_required
def api_users_delete(uid):
    if uid == session["user"]["id"]:
        return jsonify({"ok":False,"msg":"본인 계정은 삭제할 수 없습니다."}), 400
    conn = get_db()
    conn.execute("DELETE FROM users WHERE id=?", (uid,))
    conn.commit(); conn.close()
    return jsonify({"ok":True})

# ── 엑셀 내보내기 ──────────────────────────────
@app.route("/api/export/branches")
@login_required
def export_branches():
    conn = get_db()
    rows = conn.execute("SELECT * FROM branches ORDER BY name").fetchall()
    conn.close()
    buf = io.StringIO()
    w = csv.writer(buf)
    w.writerow(["ID","지사명","지역","담당자","전화","이메일","주소","상태","계약일","수수료율","메모"])
    for r in rows:
        w.writerow([r["id"],r["name"],r["region"],r["manager"],r["phone"],
                    r["email"],r["address"],r["status"],r["contract_date"],r["fee_rate"],r["note"]])
    buf.seek(0)
    return send_file(io.BytesIO(buf.getvalue().encode("utf-8-sig")), mimetype="text/csv",
                     as_attachment=True, download_name=f"지사목록_{date.today()}.csv")

@app.route("/api/export/sales")
@login_required
def export_sales():
    year = request.args.get("year", str(datetime.now().year))
    conn = get_db()
    rows = conn.execute("""
        SELECT b.name,b.region,s.month,s.target,s.actual,
               ROUND(CAST(s.actual AS REAL)/NULLIF(s.target,0)*100,1) pct
        FROM sales s JOIN branches b ON s.branch_id=b.id
        WHERE s.year=? ORDER BY b.name,s.month""", (year,)).fetchall()
    conn.close()
    buf = io.StringIO()
    w = csv.writer(buf)
    w.writerow(["지사명","지역","월","목표부수","실적부수","달성률(%)"])
    for r in rows:
        w.writerow([r["name"],r["region"],f"{r['month']}월",r["target"],r["actual"],r["pct"]])
    buf.seek(0)
    return send_file(io.BytesIO(buf.getvalue().encode("utf-8-sig")), mimetype="text/csv",
                     as_attachment=True, download_name=f"판매부수_{year}.csv")

@app.route("/api/me")
@login_required
def api_me():
    return jsonify(session.get("user",{}))


# Render/gunicorn 실행 시 자동 초기화
init_db()

if __name__ == "__main__":
    import webbrowser, threading
    threading.Timer(1.2, lambda: webbrowser.open("http://localhost:5001")).start()
    app.run(debug=False, port=5001)
