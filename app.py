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
    conn.execute("""CREATE TABLE IF NOT EXISTS sales_data (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        sale_date TEXT,
        seller_name TEXT,
        item_code TEXT,
        item_name TEXT,
        item_group TEXT,
        quantity INTEGER DEFAULT 1,
        unit_price INTEGER DEFAULT 0,
        supply_price INTEGER DEFAULT 0,
        vat INTEGER DEFAULT 0,
        total INTEGER DEFAULT 0,
        buyer TEXT,
        buyer_phone TEXT,
        real_seller TEXT,
        upload_batch TEXT,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP)""")
    conn.execute("""CREATE TABLE IF NOT EXISTS sellers (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT UNIQUE NOT NULL,
        real_name TEXT,
        first_seen TEXT,
        total_sales INTEGER DEFAULT 0,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP)""")
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

# ── 엑셀(.xlsx) 판매현황 업로드 ───────────────
def parse_xlsx_sales(file_bytes):
    """ZIP 기반으로 xlsx 직접 파싱 (스타일 오류 무시)"""
    import zipfile, xml.etree.ElementTree as ET, re
    from datetime import datetime as dt

    results = []
    with zipfile.ZipFile(io.BytesIO(file_bytes)) as z:
        # 공유문자열
        strings = []
        if 'xl/sharedStrings.xml' in z.namelist():
            sst = z.read('xl/sharedStrings.xml').decode('utf-8')
            sst_root = ET.fromstring(sst)
            ns2 = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
            for si in sst_root.findall(f'{{{ns2}}}si'):
                t_els = si.findall(f'.//{{{ns2}}}t')
                strings.append(''.join(t.text or '' for t in t_els))

        sheet_xml = z.read('xl/worksheets/sheet1.xml').decode('utf-8')
        root = ET.fromstring(sheet_xml)
        ns2 = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

        for row in root.findall(f'.//{{{ns2}}}row'):
            rnum = int(row.get('r', 0))
            if rnum <= 2: continue  # 헤더 행 스킵

            row_vals = {}
            for cell in row.findall(f'{{{ns2}}}c'):
                ref = cell.get('r', '')
                col = ''.join(c for c in ref if c.isalpha())
                t = cell.get('t', '')
                is_el = cell.find(f'{{{ns2}}}is')
                v_el = cell.find(f'{{{ns2}}}v')
                val = ''
                if is_el is not None:
                    t_els = is_el.findall(f'.//{{{ns2}}}t')
                    val = ''.join(x.text or '' for x in t_els)
                elif t == 's' and v_el is not None:
                    idx = int(v_el.text)
                    val = strings[idx] if idx < len(strings) else ''
                elif v_el is not None:
                    val = v_el.text or ''
                if val:
                    row_vals[col] = val

            if not row_vals.get('C'):
                continue  # 판매처명 없으면 스킵

            # 일자 파싱 (2026/05/03 -5 형태)
            raw_date = row_vals.get('B', '')
            sale_date = re.sub(r'\s*-\d+$', '', raw_date).strip()
            try:
                dt.strptime(sale_date, '%Y/%m/%d')
                sale_date = sale_date.replace('/', '-')
            except:
                sale_date = ''

            results.append({
                'sale_date':    sale_date,
                'seller_name':  row_vals.get('C', '').strip(),
                'item_code':    row_vals.get('G', '').strip(),
                'item_name':    row_vals.get('H', '').strip(),
                'item_group':   row_vals.get('AA', '').strip(),
                'quantity':     int(float(row_vals.get('I', 1) or 1)),
                'unit_price':   int(float(row_vals.get('K', 0) or 0)),
                'supply_price': int(float(row_vals.get('L', 0) or 0)),
                'vat':          int(float(row_vals.get('M', 0) or 0)),
                'total':        int(float(row_vals.get('N', 0) or 0)),
                'buyer':        row_vals.get('D', '').strip(),
                'buyer_phone':  row_vals.get('E', '').strip(),
                'real_seller':  row_vals.get('AE', '').strip(),
            })
    return results

@app.route("/api/upload/xlsx/preview", methods=["POST"])
@login_required
def upload_xlsx_preview():
    f = request.files.get("file")
    if not f: return jsonify({"error": "파일이 없습니다"}), 400
    try:
        rows = parse_xlsx_sales(f.read())
    except Exception as e:
        return jsonify({"error": f"파일 파싱 오류: {str(e)}"}), 400

    # 판매처 목록 추출
    sellers = {}
    for r in rows:
        name = r['seller_name']
        if name not in sellers:
            sellers[name] = {'count': 0, 'total': 0, 'real_name': r.get('real_seller', '')}
        sellers[name]['count'] += 1
        sellers[name]['total'] += r['total']

    return jsonify({
        "count": len(rows),
        "seller_count": len(sellers),
        "sellers": [{"name": k, "count": v['count'], "total": v['total'],
                     "real_name": v['real_name']} for k, v in sorted(sellers.items())],
        "sample": rows[:5],
        "date_range": {
            "from": min((r['sale_date'] for r in rows if r['sale_date']), default=''),
            "to":   max((r['sale_date'] for r in rows if r['sale_date']), default=''),
        }
    })

@app.route("/api/upload/xlsx/commit", methods=["POST"])
@login_required
def upload_xlsx_commit():
    f = request.files.get("file")
    if not f: return jsonify({"error": "파일이 없습니다"}), 400
    try:
        rows = parse_xlsx_sales(f.read())
    except Exception as e:
        return jsonify({"error": str(e)}), 400

    batch = datetime.now().strftime("%Y%m%d%H%M%S")
    conn = get_db()

    # 기존 데이터 삭제 (전체 덮어쓰기)
    conn.execute("DELETE FROM sales_data")
    conn.execute("DELETE FROM sellers")

    # 판매처 등록
    seller_set = {}
    for r in rows:
        name = r['seller_name']
        if name not in seller_set:
            seller_set[name] = {'total': 0, 'count': 0, 'real_name': r.get('real_seller', '')}
        seller_set[name]['total'] += r['total']
        seller_set[name]['count'] += 1

    today = date.today().isoformat()
    for name, info in seller_set.items():
        conn.execute("""INSERT OR IGNORE INTO sellers(name, real_name, first_seen, total_sales)
                        VALUES(?,?,?,?)""", (name, info['real_name'], today, info['total']))

    # 판매 데이터 저장
    for r in rows:
        conn.execute("""INSERT INTO sales_data
            (sale_date,seller_name,item_code,item_name,item_group,quantity,
             unit_price,supply_price,vat,total,buyer,buyer_phone,real_seller,upload_batch)
            VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            (r['sale_date'], r['seller_name'], r['item_code'], r['item_name'],
             r['item_group'], r['quantity'], r['unit_price'], r['supply_price'],
             r['vat'], r['total'], r['buyer'], r['buyer_phone'], r['real_seller'], batch))

    conn.commit(); conn.close()
    return jsonify({"ok": True, "rows": len(rows), "sellers": len(seller_set), "batch": batch})

# ── 판매현황 조회 API ──────────────────────────
@app.route("/api/sales-data/summary")
@login_required
def sales_data_summary():
    conn = get_db()
    # 전체 요약
    total = conn.execute("SELECT COUNT(*) c, SUM(total) t, SUM(quantity) q FROM sales_data").fetchone()
    # 판매처별
    by_seller = [dict(r) for r in conn.execute("""
        SELECT seller_name, COUNT(*) cnt, SUM(quantity) qty, SUM(total) total
        FROM sales_data GROUP BY seller_name ORDER BY total DESC LIMIT 20""").fetchall()]
    # 품목그룹별
    by_group = [dict(r) for r in conn.execute("""
        SELECT item_group, COUNT(*) cnt, SUM(quantity) qty, SUM(total) total
        FROM sales_data WHERE item_group != '' GROUP BY item_group ORDER BY total DESC""").fetchall()]
    # 일별
    by_date = [dict(r) for r in conn.execute("""
        SELECT sale_date, COUNT(*) cnt, SUM(total) total
        FROM sales_data WHERE sale_date != '' GROUP BY sale_date ORDER BY sale_date""").fetchall()]
    # 품목별 TOP20
    by_item = [dict(r) for r in conn.execute("""
        SELECT item_name, SUM(quantity) qty, SUM(total) total
        FROM sales_data GROUP BY item_name ORDER BY total DESC LIMIT 20""").fetchall()]
    conn.close()
    return jsonify({
        "total_count": total["c"] or 0,
        "total_amount": total["t"] or 0,
        "total_quantity": total["q"] or 0,
        "by_seller": by_seller,
        "by_group": by_group,
        "by_date": by_date,
        "by_item": by_item,
    })

@app.route("/api/sellers")
@login_required
def api_sellers():
    conn = get_db()
    rows = [dict(r) for r in conn.execute(
        "SELECT * FROM sellers ORDER BY total_sales DESC").fetchall()]
    conn.close()
    return jsonify(rows)

# ── 엑셀 템플릿 다운로드 ───────────────────────
@app.route("/api/template/branches")
@login_required
def template_branches():
    buf = io.StringIO()
    w = csv.writer(buf)
    w.writerow(["지사명","지역","담당자","전화","이메일","주소","상태","계약일","수수료율","메모"])
    w.writerow(["서울 강남지사","서울","홍길동","010-1234-5678","example@visang.com","서울시 강남구","운영중","2024-01-01",5.0,"예시 데이터"])
    buf.seek(0)
    return send_file(io.BytesIO(buf.getvalue().encode("utf-8-sig")), mimetype="text/csv",
                     as_attachment=True, download_name="지사_업로드_양식.csv")

@app.route("/api/template/sales")
@login_required
def template_sales():
    buf = io.StringIO()
    w = csv.writer(buf)
    w.writerow(["지사명","연도","1월목표","1월실적","2월목표","2월실적","3월목표","3월실적",
                "4월목표","4월실적","5월목표","5월실적","6월목표","6월실적",
                "7월목표","7월실적","8월목표","8월실적","9월목표","9월실적",
                "10월목표","10월실적","11월목표","11월실적","12월목표","12월실적"])
    w.writerow(["서울 강남지사", 2026,
                1000,850, 1200,1100, 1100,980, 1300,1250, 1400,1300, 1200,1150,
                1100,1000, 1300,1200, 1400,1350, 1500,1420, 1600,1500, 1800,1700])
    buf.seek(0)
    return send_file(io.BytesIO(buf.getvalue().encode("utf-8-sig")), mimetype="text/csv",
                     as_attachment=True, download_name="판매부수_업로드_양식.csv")

# ── 엑셀 업로드 (미리보기) ────────────────────
@app.route("/api/upload/branches/preview", methods=["POST"])
@login_required
def upload_branches_preview():
    f = request.files.get("file")
    if not f: return jsonify({"error":"파일이 없습니다"}), 400
    content = f.read().decode("utf-8-sig")
    reader = csv.DictReader(io.StringIO(content))
    rows, errors = [], []
    REGIONS = ["서울","경기","인천","강원","충북","충남","대전","세종","경북","경남","대구","부산","울산","전북","전남","광주","제주"]
    for i, row in enumerate(reader, 1):
        name = row.get("지사명","").strip()
        region = row.get("지역","").strip()
        if not name:
            errors.append(f"{i}행: 지사명 누락")
            continue
        if region and region not in REGIONS:
            errors.append(f"{i}행 [{name}]: 알 수 없는 지역 '{region}'")
        rows.append({
            "name": name, "region": region,
            "manager": row.get("담당자","").strip(),
            "phone": row.get("전화","").strip(),
            "email": row.get("이메일","").strip(),
            "address": row.get("주소","").strip(),
            "status": row.get("상태","운영중").strip() or "운영중",
            "contract_date": row.get("계약일","").strip(),
            "fee_rate": float(row.get("수수료율",0) or 0),
            "note": row.get("메모","").strip(),
        })
    return jsonify({"rows": rows, "errors": errors, "count": len(rows)})

@app.route("/api/upload/branches/commit", methods=["POST"])
@login_required
def upload_branches_commit():
    data = request.json
    rows = data.get("rows", [])
    mode = data.get("mode", "append")  # append | overwrite
    conn = get_db()
    if mode == "overwrite":
        conn.execute("DELETE FROM branches")
        conn.execute("DELETE FROM sales")
    added = 0
    for r in rows:
        existing = conn.execute("SELECT id FROM branches WHERE name=?", (r["name"],)).fetchone()
        if existing:
            conn.execute("""UPDATE branches SET region=?,manager=?,phone=?,email=?,
                            address=?,status=?,contract_date=?,fee_rate=?,note=? WHERE id=?""",
                (r["region"],r["manager"],r["phone"],r["email"],r["address"],
                 r["status"],r["contract_date"],r["fee_rate"],r["note"],existing["id"]))
        else:
            conn.execute("""INSERT INTO branches(name,region,manager,phone,email,address,status,contract_date,fee_rate,note)
                            VALUES(?,?,?,?,?,?,?,?,?,?)""",
                (r["name"],r["region"],r["manager"],r["phone"],r["email"],
                 r["address"],r["status"],r["contract_date"],r["fee_rate"],r["note"]))
            added += 1
    conn.commit(); conn.close()
    return jsonify({"ok": True, "added": added, "total": len(rows)})

@app.route("/api/upload/sales/preview", methods=["POST"])
@login_required
def upload_sales_preview():
    f = request.files.get("file")
    if not f: return jsonify({"error":"파일이 없습니다"}), 400
    content = f.read().decode("utf-8-sig")
    reader = csv.DictReader(io.StringIO(content))
    rows, errors = [], []
    conn = get_db()
    for i, row in enumerate(reader, 1):
        name = row.get("지사명","").strip()
        year = row.get("연도","").strip()
        if not name or not year:
            errors.append(f"{i}행: 지사명 또는 연도 누락"); continue
        branch = conn.execute("SELECT id FROM branches WHERE name=?", (name,)).fetchone()
        if not branch:
            errors.append(f"{i}행: '{name}' 지사가 시스템에 없음 (지사 먼저 등록 필요)")
            continue
        months = []
        for m in range(1, 13):
            t = int(row.get(f"{m}월목표", 0) or 0)
            a = int(row.get(f"{m}월실적", 0) or 0)
            months.append({"month": m, "target": t, "actual": a})
        rows.append({"branch_id": branch["id"], "branch_name": name,
                     "year": int(year), "months": months})
    conn.close()
    return jsonify({"rows": rows, "errors": errors, "count": len(rows)})

@app.route("/api/upload/sales/commit", methods=["POST"])
@login_required
def upload_sales_commit():
    data = request.json
    rows = data.get("rows", [])
    conn = get_db()
    for r in rows:
        for m in r["months"]:
            conn.execute("""INSERT INTO sales(branch_id,year,month,target,actual) VALUES(?,?,?,?,?)
                ON CONFLICT(branch_id,year,month) DO UPDATE SET target=excluded.target,actual=excluded.actual""",
                (r["branch_id"], r["year"], m["month"], m["target"], m["actual"]))
    conn.commit(); conn.close()
    return jsonify({"ok": True, "total": len(rows)})

# Render/gunicorn 실행 시 자동 초기화
init_db()

if __name__ == "__main__":
    import webbrowser, threading
    threading.Timer(1.2, lambda: webbrowser.open("http://localhost:5001")).start()
    app.run(debug=False, port=5001)
