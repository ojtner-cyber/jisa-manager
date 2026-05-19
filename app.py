"""
매장 관리 시스템 v1.0
Flask + SQLite | 본사 전용 지사 관리 플랫폼
"""
from flask import Flask, render_template, request, jsonify, session, redirect, url_for, send_file
import sqlite3, json, os, csv, io
from datetime import date, datetime
from functools import wraps

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "enfix-manager-secret-2025")
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB 허용

# Render Disk 마운트 경로 우선 사용 → 없으면 스크립트 폴더
_data_dir = "/data" if os.path.isdir("/data") else os.path.dirname(os.path.abspath(__file__))
DB_FILE   = os.path.join(_data_dir, "jisa.db")
GOAL_FILE = os.path.join(_data_dir, "sales_goals.json")

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
        ceo TEXT,
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
        note TEXT DEFAULT '',
        created_at TEXT DEFAULT CURRENT_TIMESTAMP)""")
    conn.execute("""CREATE TABLE IF NOT EXISTS sellers (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT UNIQUE NOT NULL,
        real_name TEXT,
        first_seen TEXT,
        total_sales INTEGER DEFAULT 0,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP)""")

    # ── 기존 DB 마이그레이션 (컬럼 누락 시 자동 추가) ──
    try:
        existing_cols = [r[1] for r in conn.execute("PRAGMA table_info(sales_data)").fetchall()]
        if 'note' not in existing_cols:
            conn.execute("ALTER TABLE sales_data ADD COLUMN note TEXT DEFAULT ''")
        if 'upload_batch' not in existing_cols:
            conn.execute("ALTER TABLE sales_data ADD COLUMN upload_batch TEXT DEFAULT ''")
        # branches 테이블 마이그레이션
        branch_cols = [r[1] for r in conn.execute("PRAGMA table_info(branches)").fetchall()]
        if 'ceo' not in branch_cols:
            conn.execute("ALTER TABLE branches ADD COLUMN ceo TEXT DEFAULT ''")
    except Exception:
        pass

    # 기본 계정만 생성 (샘플 데이터 없음)
    conn.execute("INSERT OR IGNORE INTO users(email,password,name,role) VALUES(?,?,?,?)",
        ("hwkim@enfix.com","hwkim123!","관리자","admin"))
    conn.execute("INSERT OR IGNORE INTO users(email,password,name,role) VALUES(?,?,?,?)",
        ("user@visang.com","hwkim123!","일반사용자","user"))
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

    # 판매현황(sales_data) 기반 실적
    sd_total = conn.execute("""SELECT SUM(total) t, COUNT(*) c FROM sales_data
        WHERE sale_date LIKE ?""", (f"{y}%",)).fetchone()

    # 월별: sales_data 기반
    monthly_sd = [dict(r) for r in conn.execute("""
        SELECT CAST(strftime('%m', sale_date) AS INTEGER) month, SUM(total) actual, COUNT(*) cnt
        FROM sales_data WHERE sale_date LIKE ? AND sale_date != ''
        GROUP BY month ORDER BY month""", (f"{y}%",)).fetchall()]
    # 목표는 sales 테이블 유지
    monthly_target = {r["month"]: r["target"] for r in conn.execute("""
        SELECT month, SUM(target) target FROM sales WHERE year=?
        GROUP BY month""", (y,)).fetchall()}
    monthly = []
    for m in range(1, 13):
        sd_row = next((r for r in monthly_sd if r["month"]==m), None)
        monthly.append({"month": m, "target": monthly_target.get(m, 0),
                        "actual": sd_row["actual"] if sd_row else 0})

    # TOP5 판매처 (real_seller 기준)
    top5 = [dict(r) for r in conn.execute("""
        SELECT real_seller name, SUM(total) total FROM sales_data
        WHERE sale_date LIKE ? AND real_seller != ''
        GROUP BY real_seller ORDER BY total DESC LIMIT 5""",
        (f"{y}%",)).fetchall()]

    # 지역별 (branches + sales_data 조인 — real_seller 기준)
    region_stats = [dict(r) for r in conn.execute("""
        SELECT b.region, SUM(sd.total) total
        FROM sales_data sd JOIN branches b ON sd.real_seller=b.name
        WHERE sd.sale_date LIKE ?
        GROUP BY b.region ORDER BY total DESC""", (f"{y}%",)).fetchall()]

    conn.close()
    total_actual = int(sd_total["t"] or 0)
    total_count  = int(sd_total["c"] or 0)
    return jsonify({
        "total_branches": total_branches,
        "active_branches": active,
        "total_target": 0,
        "total_actual": total_actual,
        "total_count": total_count,
        "achievement": 0,
        "monthly": monthly,
        "top5": top5,
        "region_stats": region_stats,
    })

# ── 판매처(지사) API ───────────────────────────
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
    y = datetime.now().year
    for row in rows:
        # 판매현황(sales_data)에서 실적 연동 — real_seller 기준
        sd = conn.execute("""
            SELECT SUM(total) total FROM sales_data
            WHERE real_seller=? AND sale_date LIKE ?""",
            (row["name"], f"{y}%")).fetchone()
        row["year_actual"] = int(sd["total"] or 0)
    conn.close()
    return jsonify(rows)

@app.route("/api/branches", methods=["POST"])
@login_required
def api_branches_add():
    d = request.json
    conn = get_db()
    conn.execute("""INSERT INTO branches(name,ceo,region,manager,phone,email,address,status,note)
        VALUES(?,?,?,?,?,?,?,?,?)""",
        (d["name"],d.get("ceo",""),d.get("region",""),d.get("manager",""),d.get("phone",""),
         d.get("email",""),d.get("address",""),d.get("status","운영중"),d.get("note","")))
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
    conn.execute("""UPDATE branches SET name=?,ceo=?,region=?,manager=?,phone=?,email=?,
        address=?,status=?,note=? WHERE id=?""",
        (d["name"],d.get("ceo",""),d.get("region",""),d.get("manager",""),d.get("phone",""),
         d.get("email",""),d.get("address",""),d.get("status","운영중"),d.get("note",""),bid))
    conn.commit(); conn.close()
    return jsonify({"ok":True})

# ── 매장 정보 xlsx 업로드 ─────────────────────
def parse_region_from_address(addr):
    """주소에서 지역 추출"""
    addr = addr or ''
    region_map = [
        ('서울', '서울'), ('경기', '경기'), ('인천', '인천'), ('강원', '강원'),
        ('충북', '충북'), ('충남', '충남'), ('대전', '대전'), ('세종', '세종'),
        ('경북', '경북'), ('경남', '경남'), ('대구', '대구'), ('부산', '부산'),
        ('울산', '울산'), ('전북', '전북'), ('전남', '전남'), ('광주', '광주'),
        ('제주', '제주'),
    ]
    for key, region in region_map:
        if key in addr:
            return region
    return ''

@app.route("/api/upload/stores", methods=["POST"])
@login_required
def upload_stores():
    """매장 정보 xlsx 업로드 — E열:실적용거래처명, F열:전화, M열:담당자, N열:주소, B열:업체구분"""
    import zipfile, xml.etree.ElementTree as ET
    f = request.files.get("file")
    if not f: return jsonify({"error": "파일이 없습니다"}), 400

    file_bytes = f.read()
    stores = []
    try:
        with zipfile.ZipFile(io.BytesIO(file_bytes)) as z:
            strings = []
            if 'xl/sharedStrings.xml' in z.namelist():
                sst = z.read('xl/sharedStrings.xml').decode('utf-8')
                sr = ET.fromstring(sst)
                ns2 = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
                for si in sr.findall(f'{{{ns2}}}si'):
                    strings.append(''.join(t.text or '' for t in si.findall(f'.//{{{ns2}}}t')))

            sheet_xml = z.read('xl/worksheets/sheet1.xml').decode('utf-8')
            root = ET.fromstring(sheet_xml)
            ns2 = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

            def cell_val(cell):
                t = cell.get('t', '')
                is_el = cell.find(f'{{{ns2}}}is')
                v_el  = cell.find(f'{{{ns2}}}v')
                if is_el is not None:
                    return ''.join(x.text or '' for x in is_el.findall(f'.//{{{ns2}}}t'))
                if t == 's' and v_el is not None:
                    idx = int(v_el.text)
                    return strings[idx] if idx < len(strings) else ''
                return v_el.text or '' if v_el is not None else ''

            current_group = ''
            for row in root.findall(f'.//{{{ns2}}}row'):
                rnum = int(row.get('r', 0))
                if rnum < 5: continue  # 헤더 스킵

                vals = {}
                for c in row.findall(f'{{{ns2}}}c'):
                    ref = c.get('r', '')
                    col = ''.join(x for x in ref if x.isalpha())
                    v = cell_val(c)
                    if v: vals[col] = v

                # B열에 업체구분이 있으면 그룹 업데이트
                if 'B' in vals and vals['B'] not in ('업체구분', '※ 오프라인 거래처별 리스트'):
                    current_group = vals['B']

                name = vals.get('E', '').strip()
                if not name: continue

                phone   = vals.get('F', '').strip()
                manager = vals.get('M', '').strip()
                address = vals.get('N', '').strip()
                region  = parse_region_from_address(address)

                stores.append({
                    'name':    name,
                    'group':   current_group,
                    'phone':   phone,
                    'manager': manager,
                    'address': address,
                    'region':  region,
                    'note':    current_group,
                })
    except Exception as e:
        return jsonify({"error": f"파일 파싱 오류: {str(e)}"}), 400

    # preview 모드
    if request.args.get('preview') == '1':
        return jsonify({"stores": stores, "count": len(stores)})

    # 저장
    conn = get_db()
    added, updated = 0, 0
    for s in stores:
        existing = conn.execute("SELECT id FROM branches WHERE name=?", (s['name'],)).fetchone()
        if existing:
            conn.execute("""UPDATE branches SET phone=?,manager=?,address=?,region=?,note=?
                WHERE id=?""", (s['phone'], s['manager'], s['address'], s['region'], s['note'], existing['id']))
            updated += 1
        else:
            conn.execute("""INSERT INTO branches(name,region,manager,phone,address,status,note)
                VALUES(?,?,?,?,?,?,?)""",
                (s['name'], s['region'], s['manager'], s['phone'], s['address'], '운영중', s['note']))
            added += 1
    conn.commit(); conn.close()
    return jsonify({"ok": True, "added": added, "updated": updated, "total": len(stores)})

# ── 판매부수 페이지용 — 매장별 실적 ──────────────
@app.route("/api/sales-by-store")
@login_required
def api_sales_by_store():
    year   = request.args.get("year",   str(datetime.now().year))
    seller = request.args.get("seller", "").strip()
    month  = request.args.get("month",  "").strip()
    conn   = get_db()

    if month:
        date_cond = f"{year}-{month.zfill(2)}%"
    else:
        date_cond = f"{year}%"

    if seller:
        # 특정 매장 → 월별 실적 반환 (seller_name 포함)
        rows = [dict(r) for r in conn.execute("""
            SELECT ? AS seller_name,
                   CAST(strftime('%m', sale_date) AS INTEGER) AS month,
                   COUNT(*) cnt, SUM(total) total, SUM(quantity) qty
            FROM sales_data
            WHERE real_seller=? AND sale_date LIKE ? AND sale_date != ''
            GROUP BY month ORDER BY month""", (seller, seller, date_cond)).fetchall()]
        conn.close()
        return jsonify(rows)
    else:
        # 전체 매장 요약
        rows = [dict(r) for r in conn.execute("""
            SELECT real_seller AS seller_name,
                   COUNT(*) cnt, SUM(total) total, SUM(quantity) qty
            FROM sales_data
            WHERE sale_date LIKE ? AND real_seller != '' AND real_seller IS NOT NULL
            GROUP BY real_seller ORDER BY total DESC""", (date_cond,)).fetchall()]
        conn.close()
        return jsonify(rows)

# ── 판매현황 — 판매처 수 전체 반환 ──────────────
@app.route("/api/sales-data/summary")
@login_required
def sales_data_summary():
    conn = get_db()
    total = conn.execute("SELECT COUNT(*) c, SUM(total) t, SUM(quantity) q FROM sales_data").fetchone()
    # 전체 판매처 (제한 없음)
    by_seller = [dict(r) for r in conn.execute("""
        SELECT real_seller seller_name, COUNT(*) cnt, SUM(quantity) qty, SUM(total) total
        FROM sales_data WHERE real_seller != ''
        GROUP BY real_seller ORDER BY total DESC""").fetchall()]
    by_group = [dict(r) for r in conn.execute("""
        SELECT item_group, COUNT(*) cnt, SUM(quantity) qty, SUM(total) total
        FROM sales_data WHERE item_group != '' GROUP BY item_group ORDER BY total DESC""").fetchall()]
    by_date = [dict(r) for r in conn.execute("""
        SELECT sale_date, COUNT(*) cnt, SUM(total) total
        FROM sales_data WHERE sale_date != '' GROUP BY sale_date ORDER BY sale_date""").fetchall()]
    by_item = [dict(r) for r in conn.execute("""
        SELECT item_name, SUM(quantity) qty, SUM(total) total
        FROM sales_data GROUP BY item_name ORDER BY total DESC LIMIT 20""").fetchall()]
    conn.close()
    return jsonify({
        "total_count": total["c"] or 0,
        "total_amount": total["t"] or 0,
        "total_quantity": total["q"] or 0,
        "seller_count": len(by_seller),
        "by_seller": by_seller,
        "by_group": by_group,
        "by_date": by_date,
        "by_item": by_item,
    })

# ── xlsx 판매현황 — real_seller 기준으로 저장 ──
@app.route("/api/branches/from-xlsx", methods=["POST"])
@login_required
def branches_from_xlsx():
    """판매현황 xlsx에서 실적용거래처명(real_seller) 기준으로 판매처 등록"""
    conn = get_db()
    sellers = [dict(r) for r in conn.execute("""
        SELECT real_seller, COUNT(*) cnt, SUM(total) total
        FROM sales_data WHERE real_seller != ''
        GROUP BY real_seller ORDER BY real_seller""").fetchall()]

    added, updated = 0, 0
    for s in sellers:
        name = s["real_seller"]
        existing = conn.execute("SELECT id FROM branches WHERE name=?", (name,)).fetchone()
        if not existing:
            conn.execute("""INSERT INTO branches(name,region,manager,phone,address,status,note)
                VALUES(?,?,?,?,?,?,?)""", (name,"","","","","운영중",""))
            added += 1
        else:
            updated += 1
    conn.commit(); conn.close()
    return jsonify({"ok": True, "added": added, "updated": updated, "total": len(sellers)})
    return jsonify({"ok": True, "added": added, "updated": updated, "total": len(sellers)})

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
    """xlsx 파싱 — 수량 -1 제외, 특이사항 '교환' 제외, 베이비하우스 본사 → 수취인으로 매장 파악"""
    import zipfile, xml.etree.ElementTree as ET, re
    from datetime import datetime as dt

    results = []
    with zipfile.ZipFile(io.BytesIO(file_bytes)) as z:
        strings = []
        if 'xl/sharedStrings.xml' in z.namelist():
            sst = z.read('xl/sharedStrings.xml').decode('utf-8')
            sst_root = ET.fromstring(sst)
            ns2 = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
            for si in sst_root.findall(f'{{{ns2}}}si'):
                strings.append(''.join(t.text or '' for t in si.findall(f'.//{{{ns2}}}t')))

        sheet_xml = z.read('xl/worksheets/sheet1.xml').decode('utf-8')
        root = ET.fromstring(sheet_xml)
        ns2 = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

        for row in root.findall(f'.//{{{ns2}}}row'):
            rnum = int(row.get('r', 0))
            if rnum <= 2: continue

            row_vals = {}
            for cell in row.findall(f'{{{ns2}}}c'):
                ref = cell.get('r', '')
                col = ''.join(c for c in ref if c.isalpha())
                t = cell.get('t', '')
                is_el = cell.find(f'{{{ns2}}}is')
                v_el  = cell.find(f'{{{ns2}}}v')
                val = ''
                if is_el is not None:
                    val = ''.join(x.text or '' for x in is_el.findall(f'.//{{{ns2}}}t'))
                elif t == 's' and v_el is not None:
                    idx = int(v_el.text)
                    val = strings[idx] if idx < len(strings) else ''
                elif v_el is not None:
                    val = v_el.text or ''
                if val:
                    row_vals[col] = val

            if not row_vals.get('C'):
                continue

            # 수량 파싱 및 -1 제외
            try:
                qty = int(float(row_vals.get('I', 0) or 0))
            except:
                qty = 0
            if qty <= 0:
                continue  # 수량 -1 또는 0 제외

            # 특이사항(P열)에 '교환' 포함 시 제외
            note = row_vals.get('P', '').strip()
            if '교환' in note:
                continue

            # 일자 파싱
            raw_date = row_vals.get('B', '')
            sale_date = re.sub(r'\s*-\d+$', '', raw_date).strip()
            try:
                dt.strptime(sale_date, '%Y/%m/%d')
                sale_date = sale_date.replace('/', '-')
            except:
                sale_date = ''

            # 실적용거래처명(AE열) 처리
            real_seller = row_vals.get('AE', '').strip()
            buyer       = row_vals.get('D', '').strip()

            # 베이비하우스_본사 → 수취인명으로 대체
            # 단, 수취인에 "고객님"이 포함된 경우는 제외 (개인 고객 주문)
            if '본사' in real_seller:
                if buyer and '고객님' in buyer:
                    continue  # 베이비하우스_본사이고 수취인이 "고객님"인 경우만 제외
                elif buyer:
                    real_seller = buyer

            # 언더바 정규화: "베이비하우스_영통점" → "베이비하우스 영통점"
            # 실적용거래처명의 언더바를 공백으로 통일
            real_seller = real_seller.replace('_', ' ')

            results.append({
                'sale_date':    sale_date,
                'seller_name':  row_vals.get('C', '').strip(),
                'item_code':    row_vals.get('G', '').strip(),
                'item_name':    row_vals.get('H', '').strip(),
                'item_group':   row_vals.get('AA', '').strip(),
                'quantity':     qty,
                'unit_price':   int(float(row_vals.get('K', 0) or 0)),
                'supply_price': int(float(row_vals.get('L', 0) or 0)),
                'vat':          int(float(row_vals.get('M', 0) or 0)),
                'total':        int(float(row_vals.get('N', 0) or 0)),
                'buyer':        buyer,
                'buyer_phone':  row_vals.get('E', '').strip(),
                'real_seller':  real_seller,
                'note':         note,
            })
    return results

@app.errorhandler(413)
def too_large(e):
    return jsonify({"error": "파일이 너무 큽니다 (최대 50MB)"}), 413

@app.route("/api/upload/xlsx/preview", methods=["POST"])
@login_required
def upload_xlsx_preview():
    f = request.files.get("file")
    if not f: return jsonify({"error": "파일이 없습니다"}), 400
    if not f.filename.lower().endswith(('.xlsx', '.xls')):
        return jsonify({"error": "xlsx 파일만 업로드 가능합니다"}), 400
    try:
        data = f.read()
        rows = parse_xlsx_sales(data)
    except Exception as e:
        return jsonify({"error": f"파일 파싱 오류: {str(e)}"}), 400

    # 날짜 범위
    dates = [r['sale_date'] for r in rows if r['sale_date']]
    d_from = min(dates, default='')
    d_to   = max(dates, default='')

    # 기간 파악 (월 단위)
    months = sorted(set(d[:7] for d in dates if d))

    # real_seller 기준 집계
    sellers = {}
    for r in rows:
        name = r['real_seller'] or r['seller_name']
        if name not in sellers:
            sellers[name] = {'count': 0, 'total': 0, 'qty': 0}
        sellers[name]['count'] += 1
        sellers[name]['total'] += r['total']
        sellers[name]['qty']   += r['quantity']

    # 이미 저장된 해당 월 데이터 여부 확인
    conn = get_db()
    existing_months = []
    for m in months:
        cnt = conn.execute("SELECT COUNT(*) FROM sales_data WHERE sale_date LIKE ?",
                           (f"{m}%",)).fetchone()[0]
        if cnt > 0:
            existing_months.append(m)
    conn.close()

    return jsonify({
        "count": len(rows),
        "seller_count": len(sellers),
        "months": months,
        "existing_months": existing_months,
        "sellers": [{"name": k, "count": v['count'], "total": v['total'], "qty": v['qty']}
                    for k, v in sorted(sellers.items(), key=lambda x: -x[1]['total'])],
        "date_range": {"from": d_from, "to": d_to},
    })

@app.route("/api/upload/xlsx/commit", methods=["POST"])
@login_required
def upload_xlsx_commit():
    f = request.files.get("file")
    if not f: return jsonify({"error": "파일이 없습니다"}), 400
    try:
        data = f.read()
        rows = parse_xlsx_sales(data)
    except Exception as e:
        return jsonify({"error": f"파싱 오류: {str(e)}"}), 400

    if not rows:
        return jsonify({"error": "유효한 데이터가 없습니다. 수량이 0 이하이거나 교환 처리된 행만 있을 수 있습니다."}), 400

    overwrite = request.form.get("overwrite", "0") == "1"
    batch = datetime.now().strftime("%Y%m%d%H%M%S")
    conn = get_db()

    # 해당 월 데이터만 교체 (누적 방식)
    dates = [r['sale_date'] for r in rows if r['sale_date']]
    months = sorted(set(d[:7] for d in dates if d))
    for m in months:
        conn.execute("DELETE FROM sales_data WHERE sale_date LIKE ?", (f"{m}%",))

    # 판매 데이터 저장
    for r in rows:
        conn.execute("""INSERT INTO sales_data
            (sale_date,seller_name,item_code,item_name,item_group,quantity,
             unit_price,supply_price,vat,total,buyer,buyer_phone,real_seller,upload_batch,note)
            VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            (r['sale_date'], r['seller_name'], r['item_code'], r['item_name'],
             r['item_group'], r['quantity'], r['unit_price'], r['supply_price'],
             r['vat'], r['total'], r['buyer'], r['buyer_phone'],
             r['real_seller'], batch, r.get('note', '')))

    conn.commit(); conn.close()
    return jsonify({"ok": True, "rows": len(rows), "months": months, "batch": batch})

@app.route("/api/sellers")
@login_required
def api_sellers():
    conn = get_db()
    rows = [dict(r) for r in conn.execute(
        "SELECT * FROM sellers ORDER BY total_sales DESC").fetchall()]
    conn.close()
    return jsonify(rows)

@app.route("/api/admin/normalize-sellers", methods=["POST"])
@login_required
def normalize_sellers():
    """기존 sales_data의 real_seller 언더바를 공백으로 정규화"""
    conn = get_db()
    # 언더바가 있는 real_seller 모두 수정
    rows = conn.execute("SELECT DISTINCT real_seller FROM sales_data WHERE real_seller LIKE '%_%'").fetchall()
    updated = 0
    for r in rows:
        old = r[0]
        new = old.replace('_', ' ')
        if old != new:
            conn.execute("UPDATE sales_data SET real_seller=? WHERE real_seller=?", (new, old))
            updated += 1
    conn.commit(); conn.close()
    return jsonify({"ok": True, "normalized": updated})

# ── 주별 세부 품목 API ─────────────────────────
@app.route("/api/sales-data/weekly-detail")
@login_required
def sales_weekly_detail():
    week_key = request.args.get("week_key", "")
    seller   = request.args.get("seller",   "").strip()
    conn     = get_db()

    params = [week_key]
    conds  = ["strftime('%Y-%W', sale_date) = ?", "sale_date != ''"]
    if seller:
        conds.append("real_seller = ?")
        params.append(seller)

    where = " AND ".join(conds)

    items = [dict(r) for r in conn.execute(f"""
        SELECT item_name, item_code, item_group,
               SUM(quantity) qty, AVG(unit_price) avg_price, SUM(total) total, COUNT(*) cnt
        FROM sales_data
        WHERE {where} AND strftime('%w', sale_date) BETWEEN '1' AND '5'
        GROUP BY item_name ORDER BY total DESC""", params).fetchall()]

    summary = conn.execute(f"""
        SELECT COUNT(*) cnt, SUM(quantity) qty, SUM(total) total,
               MIN(sale_date) date_from, MAX(sale_date) date_to
        FROM sales_data WHERE {where}
          AND strftime('%w', sale_date) BETWEEN '1' AND '5'""", params).fetchone()

    conn.close()
    return jsonify({"items": items, "summary": dict(summary), "week_key": week_key, "seller": seller})

# ── 주별 실적 API ──────────────────────────────
@app.route("/api/sales-data/weekly")
@login_required
def sales_data_weekly():
    year   = request.args.get("year",   str(datetime.now().year))
    month  = request.args.get("month",  "").strip()
    seller = request.args.get("seller", "").strip()
    conn   = get_db()

    params = []
    conds  = ["sale_date != ''"]

    if month:
        conds.append("sale_date LIKE ?")
        params.append(f"{year}-{month.zfill(2)}%")
    else:
        conds.append("sale_date LIKE ?")
        params.append(f"{year}%")

    if seller:
        conds.append("real_seller = ?")
        params.append(seller)

    where = " AND ".join(conds)

    rows = [dict(r) for r in conn.execute(f"""
        SELECT
            strftime('%Y-%W', sale_date) AS week_key,
            COUNT(*) cnt,
            SUM(quantity) qty,
            SUM(total) total,
            -- 해당 주의 월요일 계산
            date(MIN(sale_date), 'weekday 0', '-6 days') AS week_mon,
            -- 해당 주의 금요일 계산
            date(MIN(sale_date), 'weekday 0', '-2 days') AS week_fri
        FROM sales_data
        WHERE {where}
          AND strftime('%w', sale_date) BETWEEN '1' AND '5'
        GROUP BY week_key
        ORDER BY week_key""", params).fetchall()]

    # week_mon/week_fri 보정: SQLite weekday 0=일요일이므로 직접 계산
    from datetime import datetime as dt, timedelta
    for r in rows:
        try:
            # week_key: 2026-05 형식
            yr, wk = r['week_key'].split('-')
            # 해당 연도+주차의 월요일 계산
            jan1 = dt.strptime(f"{yr}-01-01", "%Y-%m-%d")
            # ISO 주차와 다를 수 있으므로 MIN(sale_date) 기반으로 계산
            # week_mon을 sale_date의 최소값 기반으로 정확히 계산
            if r.get('week_mon'):
                mon = dt.strptime(r['week_mon'], "%Y-%m-%d")
                # 해당 날짜의 월요일 찾기
                wd = mon.weekday()  # 0=월, 6=일
                actual_mon = mon - timedelta(days=wd)
                actual_fri = actual_mon + timedelta(days=4)
                r['week_start'] = actual_mon.strftime("%Y-%m-%d")
                r['week_end']   = actual_fri.strftime("%Y-%m-%d")
            else:
                r['week_start'] = ''
                r['week_end']   = ''
        except Exception:
            r['week_start'] = r.get('week_mon', '')
            r['week_end']   = r.get('week_fri', '')
    conn.close()
    return jsonify(rows)

# ── 월별 세부 품목 API ─────────────────────────
@app.route("/api/sales-data/monthly-detail")
@login_required
def sales_monthly_detail():
    year   = request.args.get("year",   str(datetime.now().year))
    month  = request.args.get("month",  "")
    seller = request.args.get("seller", "").strip()
    conn   = get_db()

    params = [f"{year}-{month.zfill(2)}%"] if month else [f"{year}%"]
    where  = "sale_date LIKE ?"
    if seller:
        where  += " AND real_seller=?"
        params.append(seller)

    # 품목별 집계
    items = [dict(r) for r in conn.execute(f"""
        SELECT item_name, item_code, item_group,
               SUM(quantity) qty, AVG(unit_price) avg_price, SUM(total) total, COUNT(*) cnt
        FROM sales_data
        WHERE {where} AND sale_date != ''
        GROUP BY item_name ORDER BY total DESC""", params).fetchall()]

    # 요약
    summary = conn.execute(f"""
        SELECT COUNT(*) cnt, SUM(quantity) qty, SUM(total) total
        FROM sales_data WHERE {where} AND sale_date != ''""", params).fetchone()

    conn.close()
    return jsonify({
        "items": items,
        "summary": dict(summary),
        "year": year, "month": month, "seller": seller
    })

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
