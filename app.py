"""
매장 관리 시스템 v1.0
Flask + SQLite | 본사 전용 지사 관리 플랫폼
"""
from flask import Flask, render_template, request, jsonify, session, redirect, url_for, send_file
import sqlite3, json, os, csv, io
from datetime import date, datetime
from functools import wraps
try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

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
        ceo_phone TEXT,
        store_manager TEXT,
        store_manager_phone TEXT,
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
        if 'ceo_phone' not in branch_cols:
            conn.execute("ALTER TABLE branches ADD COLUMN ceo_phone TEXT DEFAULT ''")
        if 'store_manager' not in branch_cols:
            conn.execute("ALTER TABLE branches ADD COLUMN store_manager TEXT DEFAULT ''")
        if 'store_manager_phone' not in branch_cols:
            conn.execute("ALTER TABLE branches ADD COLUMN store_manager_phone TEXT DEFAULT ''")
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
    conn.execute("""INSERT INTO branches(name,ceo,ceo_phone,store_manager,store_manager_phone,
        region,manager,phone,email,address,status,note) VALUES(?,?,?,?,?,?,?,?,?,?,?,?)""",
        (d["name"],d.get("ceo",""),d.get("ceo_phone",""),d.get("store_manager",""),
         d.get("store_manager_phone",""),d.get("region",""),d.get("manager",""),
         d.get("phone",""),d.get("email",""),d.get("address",""),d.get("status","운영중"),d.get("note","")))
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
    conn.execute("""UPDATE branches SET name=?,ceo=?,ceo_phone=?,store_manager=?,
        store_manager_phone=?,region=?,manager=?,phone=?,email=?,address=?,status=?,note=?
        WHERE id=?""",
        (d["name"],d.get("ceo",""),d.get("ceo_phone",""),d.get("store_manager",""),
         d.get("store_manager_phone",""),d.get("region",""),d.get("manager",""),
         d.get("phone",""),d.get("email",""),d.get("address",""),d.get("status","운영중"),
         d.get("note",""),bid))
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

def detect_region_from_name(name):
    """매장명에서 지역 자동 추출"""
    name = name or ''
    region_keywords = [
        # 특별시/광역시
        ('서울', '서울'), ('강남', '서울'), ('강북', '서울'), ('강서', '서울'),
        ('강동', '서울'), ('마포', '서울'), ('용산', '서울'), ('성북', '서울'),
        ('송파', '서울'), ('노원', '서울'), ('은평', '서울'), ('도봉', '서울'),
        ('관악', '서울'), ('동작', '서울'), ('영등포', '서울'), ('구로', '서울'),
        ('금천', '서울'), ('양천', '서울'), ('마곡', '서울'), ('목동', '서울'),
        ('부산', '부산'), ('해운대', '부산'), ('동래', '부산'), ('사하', '부산'),
        ('연제', '부산'), ('수영', '부산'), ('금정', '부산'), ('남구', '부산'),
        ('대구', '대구'), ('달성', '대구'), ('수성', '대구'), ('달서', '대구'),
        ('인천', '인천'), ('부평', '인천'), ('송도', '인천'), ('계양', '인천'),
        ('광주', '광주'), ('북구', '광주'), ('서구', '광주'),
        ('대전', '대전'), ('유성', '대전'), ('서대전', '대전'),
        ('울산', '울산'),
        ('세종', '세종'),
        # 경기
        ('수원', '경기'), ('성남', '경기'), ('고양', '경기'), ('용인', '경기'),
        ('부천', '경기'), ('안산', '경기'), ('안양', '경기'), ('남양주', '경기'),
        ('화성', '경기'), ('평택', '경기'), ('의정부', '경기'), ('시흥', '경기'),
        ('파주', '경기'), ('김포', '경기'), ('광명', '경기'), ('광주', '경기'),
        ('군포', '경기'), ('하남', '경기'), ('오산', '경기'), ('이천', '경기'),
        ('양주', '경기'), ('구리', '경기'), ('안성', '경기'), ('포천', '경기'),
        ('의왕', '경기'), ('여주', '경기'), ('동두천', '경기'), ('과천', '경기'),
        ('가평', '경기'), ('양평', '경기'), ('연천', '경기'), ('영통', '경기'),
        ('동탄', '경기'), ('판교', '경기'), ('분당', '경기'), ('일산', '경기'),
        ('서수원', '경기'), ('다산', '경기'), ('미사', '경기'),
        # 강원
        ('강원', '강원'), ('춘천', '강원'), ('원주', '강원'), ('강릉', '강원'),
        ('속초', '강원'), ('동해', '강원'), ('삼척', '강원'), ('태백', '강원'),
        # 충청
        ('청주', '충북'), ('충주', '충북'), ('제천', '충북'),
        ('천안', '충남'), ('아산', '충남'), ('서산', '충남'), ('당진', '충남'),
        ('홍성', '충남'), ('공주', '충남'), ('보령', '충남'),
        # 전라
        ('전주', '전북'), ('익산', '전북'), ('군산', '전북'), ('완주', '전북'),
        ('목포', '전남'), ('여수', '전남'), ('순천', '전남'), ('나주', '전남'),
        ('광양', '전남'),
        # 경상
        ('포항', '경북'), ('경주', '경북'), ('구미', '경북'), ('안동', '경북'),
        ('영천', '경북'), ('경산', '경북'),
        ('창원', '경남'), ('진주', '경남'), ('김해', '경남'), ('양산', '경남'),
        ('거제', '경남'), ('통영', '경남'), ('밀양', '경남'),
        # 제주
        ('제주', '제주'), ('서귀포', '제주'),
    ]
    for keyword, region in region_keywords:
        if keyword in name:
            return region
    return ''

@app.route("/api/upload/stores", methods=["POST"])
@login_required
def upload_stores():
    """매장 정보 xlsx 업로드
    B열:업체구분, D열:거래처명, E열:실적용거래처명, F열:매장전화,
    G열:사장님이름, H열:사장연락처, I열:점장이름, J열:점장연락처,
    M열:담당자, N열:주소, O열:이메일
    """
    import zipfile as zf2, xml.etree.ElementTree as ET2
    f = request.files.get("file")
    if not f: return jsonify({"error": "파일이 없습니다"}), 400

    file_bytes = f.read()
    stores = []
    try:
        with zf2.ZipFile(io.BytesIO(file_bytes)) as z:
            strings = []
            if 'xl/sharedStrings.xml' in z.namelist():
                sst = z.read('xl/sharedStrings.xml').decode('utf-8')
                sr = ET2.fromstring(sst)
                ns2 = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
                for si in sr.findall(f'{{{ns2}}}si'):
                    strings.append(''.join(t.text or '' for t in si.findall(f'.//{{{ns2}}}t')))

            sheet_xml = z.read('xl/worksheets/sheet1.xml').decode('utf-8')
            root = ET2.fromstring(sheet_xml)
            ns2 = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

            def cell_val(cell, ns2=ns2, strings=strings):
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
                if rnum < 5: continue  # 헤더 스킵 (4행이 헤더)

                vals = {}
                for c in row.findall(f'{{{ns2}}}c'):
                    ref = c.get('r', '')
                    col = ''.join(x for x in ref if x.isalpha())
                    v = cell_val(c)
                    if v: vals[col] = v

                # B열에 업체구분이 있으면 그룹 업데이트
                if 'B' in vals and vals['B'] not in ('업체구분', '※ 오프라인 거래처별 리스트'):
                    current_group = vals['B']

                # E열: 실적용거래처명이 기준 (없으면 D열 사용)
                name = vals.get('E', '').strip() or vals.get('D', '').strip()
                if not name: continue

                # 이름 정제: "이정현사장님" → "이정현", "이준석점장님" → "이준석"
                def clean_name(s):
                    return s.replace('사장님','').replace('점장님','').replace('매니저님','').replace('실장','').replace('과장','').strip()

                address = vals.get('N', '').strip()
                region  = parse_region_from_address(address) or detect_region_from_name(name)

                stores.append({
                    'name':                 name.replace('_', ' '),
                    'group':                current_group,
                    'phone':                vals.get('F', '').strip(),       # 매장 전화
                    'ceo':                  clean_name(vals.get('G', '')),   # 사장님 이름
                    'ceo_phone':            vals.get('H', '').strip(),       # 사장 연락처
                    'store_manager':        clean_name(vals.get('I', '') or vals.get('K', '')),  # 점장
                    'store_manager_phone':  vals.get('J', '').strip() or vals.get('L', '').strip(),  # 점장 연락처
                    'manager':              vals.get('M', '').strip(),       # 담당자(본사)
                    'address':              address,
                    'region':               region,
                    'email':                vals.get('O', '').strip(),
                    'note':                 current_group,
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
            conn.execute("""UPDATE branches SET phone=?,ceo=?,ceo_phone=?,store_manager=?,
                store_manager_phone=?,manager=?,address=?,region=?,email=?,note=?
                WHERE id=?""", (s['phone'],s['ceo'],s['ceo_phone'],s['store_manager'],
                s['store_manager_phone'],s['manager'],s['address'],s['region'],s['email'],
                s['note'],existing['id']))
            updated += 1
        else:
            conn.execute("""INSERT INTO branches(name,region,ceo,ceo_phone,store_manager,
                store_manager_phone,manager,phone,address,email,status,note)
                VALUES(?,?,?,?,?,?,?,?,?,?,?,?)""",
                (s['name'],s['region'],s['ceo'],s['ceo_phone'],s['store_manager'],
                 s['store_manager_phone'],s['manager'],s['phone'],s['address'],
                 s['email'],'운영중',s['note']))
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
        # 전체 매장 요약 — 브랜드별 그룹 정렬 (베이비하우스 → 링크맘 → 기타)
        rows = [dict(r) for r in conn.execute("""
            SELECT real_seller AS seller_name,
                   COUNT(*) cnt, SUM(total) total, SUM(quantity) qty
            FROM sales_data
            WHERE sale_date LIKE ? AND real_seller != '' AND real_seller IS NOT NULL
            GROUP BY real_seller ORDER BY real_seller""", (date_cond,)).fetchall()]
        conn.close()

        def brand_sort_key(r):
            nm = (r['seller_name'] or '').replace('_', ' ').lower()
            if '베이비하우스' in nm: return (0, -r['total'])
            if '링크맘' in nm:      return (1, -r['total'])
            if '베이비파크' in nm:  return (2, -r['total'])
            if '베네피아' in nm:    return (3, -r['total'])
            return (9, -r['total'])

        rows.sort(key=brand_sort_key)
        return jsonify(rows)

# ── 판매현황 — 판매처 수 전체 반환 ──────────────
# ── 판매실적 엑셀 내보내기 ────────────────────────
@app.route("/api/export/sales-monthly")
@login_required
def export_sales_monthly():
    year   = request.args.get("year",   str(datetime.now().year))
    month  = request.args.get("month",  "")
    seller = request.args.get("seller", "").strip()
    date_cond = f"{year}-{month.zfill(2)}%" if month else f"{year}%"
    conn = get_db()
    if seller:
        rows = [dict(r) for r in conn.execute("""
            SELECT ? AS seller_name, CAST(strftime('%m', sale_date) AS INTEGER) AS month,
                   COUNT(*) cnt, SUM(total) total, SUM(quantity) qty
            FROM sales_data WHERE real_seller=? AND sale_date LIKE ? AND sale_date != ''
            GROUP BY month ORDER BY month""", (seller, seller, date_cond)).fetchall()]
    else:
        rows = [dict(r) for r in conn.execute("""
            SELECT real_seller AS seller_name, COUNT(*) cnt, SUM(total) total, SUM(quantity) qty
            FROM sales_data WHERE sale_date LIKE ? AND real_seller != ''
            GROUP BY real_seller ORDER BY total DESC""", (date_cond,)).fetchall()]
    conn.close()
    buf = io.StringIO()
    w = csv.writer(buf)
    if seller:
        w.writerow(['매장명', '월', '판매건수', '판매수량', '판매금액'])
        for r in rows:
            w.writerow([r['seller_name'], f"{r['month']}월", r['cnt'], r['qty'], r['total']])
    else:
        w.writerow(['매장명', '판매건수', '판매수량', '판매금액'])
        for r in rows:
            w.writerow([r['seller_name'], r['cnt'], r['qty'], r['total']])
    buf.seek(0)
    fname = f"월별실적_{year}{'_'+month+'월' if month else ''}.csv"
    return send_file(io.BytesIO(buf.getvalue().encode('utf-8-sig')), mimetype='text/csv',
                     as_attachment=True, download_name=fname)

@app.route("/api/export/sales-weekly")
@login_required
def export_sales_weekly():
    year   = request.args.get("year",   str(datetime.now().year))
    month  = request.args.get("month",  "")
    seller = request.args.get("seller", "").strip()
    from datetime import datetime as dt2, timedelta
    qp = ["sale_date != ''"]
    pp = []
    if month: qp.append("sale_date LIKE ?"); pp.append(f"{year}-{month.zfill(2)}%")
    else:     qp.append("sale_date LIKE ?"); pp.append(f"{year}%")
    if seller: qp.append("real_seller = ?"); pp.append(seller)
    conn = get_db()
    rows = [dict(r) for r in conn.execute(f"""
        SELECT strftime('%Y-%W', sale_date) AS week_key,
               COUNT(*) cnt, SUM(quantity) qty, SUM(total) total, MIN(sale_date) AS min_date
        FROM sales_data WHERE {' AND '.join(qp)} AND sale_date != ''
        GROUP BY week_key ORDER BY week_key""", pp).fetchall()]
    conn.close()
    def wr(ds):
        d = dt2.strptime(ds, "%Y-%m-%d"); wd=d.weekday()
        sun=d-timedelta(days=(wd+1)%7)
        return sun.strftime("%Y-%m-%d"), (sun+timedelta(days=6)).strftime("%Y-%m-%d")
    for r in rows:
        try: r['ws'],r['we']=wr(r['min_date'])
        except: r['ws']=r['we']=''
    buf = io.StringIO(); w = csv.writer(buf)
    w.writerow(['주차','기간 시작','기간 종료','판매건수','판매수량','판매금액'])
    for i,r in enumerate(rows):
        w.writerow([f"{i+1}주차",r['ws'],r['we'],r['cnt'],r['qty'],r['total']])
    buf.seek(0)
    fname = f"주별실적_{year}{'_'+month+'월' if month else ''}.csv"
    return send_file(io.BytesIO(buf.getvalue().encode('utf-8-sig')), mimetype='text/csv',
                     as_attachment=True, download_name=fname)

@app.route("/api/export/sales-ranking")
@login_required
def export_sales_ranking():
    year  = request.args.get("year",  str(datetime.now().year))
    month = request.args.get("month", "")
    date_cond = f"{year}-{month.zfill(2)}%" if month else f"{year}%"
    conn = get_db()
    rows = [dict(r) for r in conn.execute("""
        SELECT real_seller AS seller_name, COUNT(*) cnt, SUM(total) total, SUM(quantity) qty
        FROM sales_data WHERE sale_date LIKE ? AND real_seller != ''
        GROUP BY real_seller ORDER BY total DESC""", (date_cond,)).fetchall()]
    conn.close()
    buf = io.StringIO(); w = csv.writer(buf)
    w.writerow(['순위','매장명','판매금액','판매건수','판매수량'])
    for i,r in enumerate(rows):
        w.writerow([i+1,r['seller_name'],r['total'],r['cnt'],r['qty']])
    buf.seek(0)
    fname = f"매출순위_{year}{'_'+month+'월' if month else ''}.csv"
    return send_file(io.BytesIO(buf.getvalue().encode('utf-8-sig')), mimetype='text/csv',
                     as_attachment=True, download_name=fname)

# ── xlsx 엑셀 내보내기 헬퍼 ──────────────────────

# 브랜드 순서 (엑셀 열 순서)
# ── 매장명 별칭 매핑 ─────────────────────────────
SELLER_ALIAS = {
    '주식회사 위드에이컴퍼니': '베이비하우스 관악점',
    '위드에이컴퍼니':          '베이비하우스 관악점',
}
def resolve_seller(name):
    return SELLER_ALIAS.get(name, name)

BRAND_ORDER = ['줄즈', '레카로', 'ABC디자인', '원더폴드', '카오스', '엔픽스', '타프토이즈']

GROUP_REMAP = {
    '식탁의자':    '카오스',       # [카오스] 제품
    '하이체어':    '엔픽스',       # [엔픽스]비바체 → 엔픽스
    '보행기':      '엔픽스',
    '쏘서':        '엔픽스',
    '점퍼루':      '엔픽스',
    '휴대용부스터': '엔픽스',
    '유아섬유류':   '',            # 제품명으로 파악 (아래 로직)
    'TAFTOYS':     '타프토이즈',
    '컨버터블카시트': '레카로',
    '주니어카시트':  '레카로',
    '유모차':      '',            # 제품명 브랜드로 파악 (아래 로직)
    '웨건':        '원더폴드',
    '카시트':      '레카로',
}

def remap_group(group, item_name=''):
    """품목그룹을 브랜드명으로 정규화"""
    g    = (group or '').strip()
    item = (item_name or '').lower()

    # 제품명에서 브랜드 추출 [브랜드]제품명 형태
    import re
    brand_match = re.match(r'\[([^\]]+)\]', item_name or '')
    brand_tag   = brand_match.group(1) if brand_match else ''

    # 유아섬유류 — 제품명 브랜드로
    if g == '유아섬유류':
        if '줄즈' in brand_tag:       return '줄즈'
        if '레카로' in brand_tag:     return '레카로'
        if 'abc' in brand_tag.lower(): return 'ABC디자인'
        if '원더폴드' in brand_tag:   return '원더폴드'
        if '엔픽스' in brand_tag:     return '엔픽스'
        if '타프' in brand_tag or 'taft' in brand_tag.lower(): return '타프토이즈'
        return 'ABC디자인'  # 기본값

    # 유모차 그룹 — 브랜드 태그로 구분
    if g == '유모차':
        if '줄즈' in brand_tag:                        return '줄즈'
        if 'abc' in brand_tag.lower():                  return 'ABC디자인'
        if '원더폴드' in brand_tag:                    return '원더폴드'
        if '레카로' in brand_tag:                      return '레카로'
        return '줄즈'  # 기본값

    # TAFTOYS 내 엔픽스 제품 예외
    if g == 'TAFTOYS' and '엔픽스' in brand_tag:
        return '엔픽스'

    # 나머지 매핑
    return GROUP_REMAP.get(g, g or '기타')

def normalize_item_name(name):
    """제품명에서 색상/옵션 제거"""
    if not name: return name
    import re
    # [브랜드]모델명_색상 → [브랜드]모델명
    cleaned = re.sub(r'_[가-힣a-zA-Z0-9\-]+$', '', name).strip()
    return cleaned if cleaned else name

def get_group_sort_key(group):
    """브랜드 정렬 순서"""
    try:
        return BRAND_ORDER.index(group)
    except ValueError:
        return 99
# ── 타프토이즈 제품 카탈로그 ─────────────────────
TAFTOYS_CATALOG = {
    '[타프토이즈]드라이브&디스커버트래블토이': {'price':26900,'category':'트래블토이','desc':'이동 중 아이 집중도 UP, 유모차 부착 가능'},
    '[타프토이즈]사바나 어드벤쳐 아치':       {'price':28600,'category':'아치/모빌','desc':'바닥 놀이 필수템, 감각 발달 + 인테리어 효과'},
    '[타프토이즈]라이드타임비지북':            {'price':23600,'category':'비지북','desc':'유모차·카시트 부착, 0-3세 인지발달'},
    '[타프토이즈]트로피컬 오케스트라 아치 모빌':{'price':27800,'category':'아치/모빌','desc':'뮤지컬 아치, 터미타임 필수'},
    '[타프토이즈]코알라 카 휠 토이':          {'price':30900,'category':'카시트 장난감','desc':'카시트 부착, 지루한 이동 시간 해결사'},
    '[타프토이즈]어반가든 팝업 티슈 박스':    {'price':22500,'category':'감각 장난감','desc':'무한 반복 놀이, 소근육 발달 최고'},
    '[타프토이즈]어반가든 유모차 모빌':       {'price':15600,'category':'아치/모빌','desc':'유모차 클립형, 시각 자극 + 휴대성'},
    '[타프토이즈]피크 앤 플레이 큐브':        {'price':18400,'category':'큐브','desc':'6면 다기능, 0-2세 전방위 발달'},
    '[타프토이즈]마이홈비지북':               {'price':18400,'category':'비지북','desc':'집 모양 비지북, 역할놀이 시작'},
    '[타프토이즈]어반가든 액티비티 큐브':     {'price':15300,'category':'큐브','desc':'4면 액티비티, 혼자 놀기 최적'},
    '[타프토이즈]사바나 360 액티비티짐':      {'price':89000,'category':'액티비티짐','desc':'360도 회전 아치, 신생아부터 12개월'},
    '[타프토이즈]사바나 터미타임 북':         {'price':19800,'category':'터미타임','desc':'엎드려 놀기 훈련, 목 근육 강화'},
    '[타프토이즈]어반가든 뮤지컬 버니':       {'price':28900,'category':'인형/뮤지컬','desc':'뮤지컬 봉제 인형, 수면 루틴 도움'},
    '[타프토이즈]어반가든 터미타임 스피닝북': {'price':22000,'category':'터미타임','desc':'스피닝 기능, 아이 주의 집중'},
    '[타프토이즈]사바나 디스커버리 큐브':     {'price':32900,'category':'큐브','desc':'프리미엄 큐브, 1-3세 탐색놀이'},
    '[타프토이즈]북극 액티비티 북':           {'price':21000,'category':'비지북','desc':'천 소재 액티비티 북, 감촉 자극'},
    '[타프토이즈]아이스크림 베어 워터매트':   {'price':35000,'category':'워터매트','desc':'여름 필수템, 터미타임 + 시각자극'},
    '[타프토이즈]팬더 블룸 워터매트':         {'price':35000,'category':'워터매트','desc':'실내 물놀이, 감각 자극 극대화'},
    '[타프토이즈]팝앤플레이스테이션':         {'price':45000,'category':'액티비티짐','desc':'팝업 텐트형, 실내 놀이공간 완성'},
    '[타프토이즈]파멜라 레인스틱':            {'price':16000,'category':'감각 장난감','desc':'청각 자극, 비 소리 감각놀이'},
    '[타프토이즈]코알라 액티비티 스파이럴':   {'price':14500,'category':'트래블토이','desc':'유모차/카시트 나선형, 다양한 질감'},
    '[타프토이즈]미니문 유모차 모빌':         {'price':13500,'category':'아치/모빌','desc':'초소형 모빌, 어디든 클립 부착'},
    '[타프토이즈]베어 허그 스파이럴':         {'price':14500,'category':'트래블토이','desc':'곰돌이 스파이럴, 촉감+색상 자극'},
}

@app.route("/api/script/analysis")
@login_required
def api_script_analysis():
    seller_raw = request.args.get("seller","").strip()
    year       = request.args.get("year", str(datetime.now().year))
    seller     = resolve_seller(seller_raw)
    conn       = get_db()

    sold_items=[dict(r) for r in conn.execute("""
        SELECT item_group,item_name,SUM(quantity) qty,SUM(total) total,COUNT(*) cnt,
               MIN(sale_date) first_sale,MAX(sale_date) last_sale
        FROM sales_data WHERE (real_seller=? OR real_seller=?) AND sale_date LIKE ? AND sale_date!=''
        GROUP BY item_name ORDER BY total DESC""",(seller,seller_raw,f"{year}%")).fetchall()]

    brand_summary={}
    for r in sold_items:
        b=remap_group(r['item_group'],r['item_name'])
        if b not in brand_summary: brand_summary[b]={'qty':0,'total':0}
        brand_summary[b]['qty']+=r['qty']; brand_summary[b]['total']+=r['total']

    sold_taft=set(normalize_item_name(r['item_name']) for r in sold_items
                  if remap_group(r['item_group'],r['item_name'])=='타프토이즈')
    unsold_taft=[{'name':normalize_item_name(k),'category':v['category'],
                  'price':v['price'],'desc':v['desc']}
                 for k,v in TAFTOYS_CATALOG.items()
                 if normalize_item_name(k) not in sold_taft]

    daily=[dict(r) for r in conn.execute("""
        SELECT sale_date,SUM(total) total,SUM(quantity) qty,COUNT(*) cnt
        FROM sales_data WHERE (real_seller=? OR real_seller=?) AND sale_date LIKE ? AND sale_date!=''
        GROUP BY sale_date ORDER BY sale_date""",(seller,seller_raw,f"{year}%")).fetchall()]

    weekly_raw=conn.execute("""
        SELECT strftime('%Y-%W',sale_date) wk,MIN(sale_date) md,SUM(total) total,SUM(quantity) qty
        FROM sales_data WHERE (real_seller=? OR real_seller=?) AND sale_date LIKE ? AND sale_date!=''
        GROUP BY wk ORDER BY wk""",(seller,seller_raw,f"{year}%")).fetchall()

    from datetime import datetime as dt2,timedelta
    weekly=[]
    for r in weekly_raw:
        try:
            d=dt2.strptime(r[1],"%Y-%m-%d"); sun=d-timedelta(days=(d.weekday()+1)%7)
            weekly.append({'week':r[0],'week_start':sun.strftime("%Y-%m-%d"),
                           'week_end':(sun+timedelta(days=6)).strftime("%Y-%m-%d"),'total':r[2],'qty':r[3]})
        except: pass

    total_all=conn.execute(f"SELECT SUM(total) FROM sales_data WHERE sale_date LIKE '{year}%'").fetchone()[0] or 1
    seller_total=sum(r['total'] for r in sold_items)
    conn.close()

    return jsonify({
        'seller':seller,'year':year,'total':seller_total,
        'total_pct':round(seller_total/total_all*100,1),
        'brand_summary':[{'brand':k,'qty':v['qty'],'total':v['total'],
                          'pct':round(v['total']/seller_total*100,1) if seller_total else 0}
                         for k,v in sorted(brand_summary.items(),key=lambda x:-x[1]['total'])],
        'sold_items':sold_items,'top5':sold_items[:5],
        'unsold_taft':unsold_taft,'daily':daily,'weekly':weekly,
    })

@app.route("/api/script/generate",methods=["POST"])
@login_required
def api_script_generate():
    data=request.json; seller=data.get('seller',''); analysis=data.get('analysis',{})
    brand_lines='\n'.join(f"  - {b['brand']}: {b['total']:,}원 ({b['pct']}%)"
                          for b in analysis.get('brand_summary',[])[:6])
    top5_lines='\n'.join(f"  - {normalize_item_name(r['item_name'])}: {r['qty']}개 ({r['total']:,}원)"
                         for r in analysis.get('top5',[]))
    unsold=analysis.get('unsold_taft',[])[:8]
    unsold_lines='\n'.join(f"  - {u['name']} ({u['category']}, {u['price']:,}원): {u['desc']}"
                           for u in unsold)
    prompt=f"""당신은 유아용품 브랜드 ENFIX의 10년 경력 영업 전문가입니다.
아래 매장 실적 데이터를 바탕으로, 매장 사장님과의 실전 영업 미팅 스크립트를 작성해주세요.

【매장명】 {seller}
【분석 연도】 {analysis.get('year','')}년  
【총 매출】 {analysis.get('total',0):,}원 (전체 대비 {analysis.get('total_pct',0)}%)

【브랜드별 실적】
{brand_lines}

【베스트 제품 TOP5】
{top5_lines}

【미취급 타프토이즈 제품】
{unsold_lines}

다음 구조로 실전 영업 스크립트를 작성해주세요:

1. **오프닝** (방문 인사 + 지난 방문 이후 관계 언급)
2. **실적 공유 & 칭찬** (구체적 수치 활용, 사장님 자부심 자극)
3. **베스트 제품 분석** (왜 잘 팔리는지 + 추가 발주 제안 멘트)
4. **타프토이즈 신규 제안** (3개 제품, 왜 이 매장에 맞는지 구체적으로)
5. **시즌 전략 제안** (현재 시기 + 다음 시즌 예측)
6. **클로징** (발주 유도 + 다음 방문 약속)

실제 영업사원이 대화하듯 자연스럽게, 한국어로 작성하세요."""
    return jsonify({'prompt':prompt,'seller':seller})

@app.route("/api/export/xlsx/script")
@login_required
def export_xlsx_script():
    seller_raw=request.args.get("seller","").strip()
    year=request.args.get("year",str(datetime.now().year))
    seller=resolve_seller(seller_raw)
    conn=get_db()
    sold_items=[dict(r) for r in conn.execute("""
        SELECT item_group,item_name,SUM(quantity) qty,SUM(total) total
        FROM sales_data WHERE (real_seller=? OR real_seller=?) AND sale_date LIKE ? AND sale_date!=''
        GROUP BY item_name ORDER BY total DESC""",(seller,seller_raw,f"{year}%")).fetchall()]
    brand_summary={}
    for r in sold_items:
        b=remap_group(r['item_group'],r['item_name'])
        if b not in brand_summary: brand_summary[b]={'qty':0,'total':0}
        brand_summary[b]['qty']+=r['qty']; brand_summary[b]['total']+=r['total']
    sold_taft=set(normalize_item_name(r['item_name']) for r in sold_items
                  if remap_group(r['item_group'],r['item_name'])=='타프토이즈')
    unsold=[{'name':normalize_item_name(k),'category':v['category'],'price':v['price'],'desc':v['desc']}
            for k,v in TAFTOYS_CATALOG.items() if normalize_item_name(k) not in sold_taft]
    weekly_raw=conn.execute("""SELECT strftime('%Y-%W',sale_date) wk,MIN(sale_date) md,SUM(total) total,SUM(quantity) qty
        FROM sales_data WHERE (real_seller=? OR real_seller=?) AND sale_date LIKE ? AND sale_date!=''
        GROUP BY wk ORDER BY wk""",(seller,seller_raw,f"{year}%")).fetchall()
    conn.close()

    from datetime import datetime as dt2,timedelta
    seller_total=sum(v['total'] for v in brand_summary.values())
    wb=openpyxl.Workbook()
    mf=lambda h: PatternFill(start_color=h,end_color=h,fill_type="solid")
    mft=lambda h,b=False,s=10: Font(color=h,bold=b,size=s)
    thin=Side(style='thin',color='E0E0E0')
    bdr=Border(left=thin,right=thin,top=thin,bottom=thin)
    ctr=Alignment(horizontal="center",vertical="center"); rgt=Alignment(horizontal="right")

    # 시트1: 브랜드별 실적
    ws1=wb.active; ws1.title="브랜드별 실적"
    ws1.merge_cells("A1:E1")
    c=ws1.cell(row=1,column=1,value=f"{seller} — {year}년 브랜드별 판매 실적")
    c.fill=mf("1E3A5F"); c.font=mft("FFFFFF",True,13); c.alignment=ctr; ws1.row_dimensions[1].height=28
    for ci,h in enumerate(['브랜드','판매수량','판매금액(원)','비율(%)','등급'],1):
        c=ws1.cell(row=2,column=ci,value=h); c.fill=mf("F2F2F2"); c.font=mft("595959",True,10); c.alignment=ctr; c.border=bdr
    for ri,(brand,v) in enumerate(sorted(brand_summary.items(),key=lambda x:-x[1]['total']),3):
        pct=round(v['total']/seller_total*100,1) if seller_total else 0
        grade="★★★ 핵심" if pct>25 else "★★ 주력" if pct>10 else "★ 보조" if pct>3 else "△ 소량"
        for ci,val in enumerate([brand,v['qty'],v['total'],pct,grade],1):
            c=ws1.cell(row=ri,column=ci,value=val); c.border=bdr
            if ri%2==0: c.fill=mf("FAFAFA")
            if ci==3: c.number_format='#,##0'; c.alignment=rgt
            if ci in (4,5): c.alignment=ctr
    for ci,w in zip('ABCDE',[14,10,16,10,14]): ws1.column_dimensions[ci].width=w

    # 시트2: 제품별 상세
    ws2=wb.create_sheet("제품별 상세")
    ws2.merge_cells("A1:F1")
    c=ws2.cell(row=1,column=1,value=f"{seller} — 제품별 판매 상세")
    c.fill=mf("1E3A5F"); c.font=mft("FFFFFF",True,12); c.alignment=ctr; ws2.row_dimensions[1].height=26
    for ci,h in enumerate(['브랜드','제품명','판매수량','판매금액(원)','비율(%)','등급'],1):
        c=ws2.cell(row=2,column=ci,value=h); c.fill=mf("F2F2F2"); c.font=mft("595959",True,10); c.alignment=ctr; c.border=bdr
    for ri,r in enumerate(sold_items,3):
        brand=remap_group(r['item_group'],r['item_name']); norm=normalize_item_name(r['item_name'])
        pct=round(r['total']/seller_total*100,1) if seller_total else 0
        grade="◎ 인기" if pct>10 else "○ 판매중" if pct>3 else "△ 소량"
        for ci,val in enumerate([brand,norm,r['qty'],r['total'],pct,grade],1):
            c=ws2.cell(row=ri,column=ci,value=val); c.border=bdr
            if ri%2==0: c.fill=mf("FAFAFA")
            if ci==4: c.number_format='#,##0'; c.alignment=rgt
            if ci in (5,6): c.alignment=ctr
    ws2.column_dimensions['A'].width=14; ws2.column_dimensions['B'].width=32
    ws2.column_dimensions['C'].width=10; ws2.column_dimensions['D'].width=16
    ws2.column_dimensions['E'].width=10; ws2.column_dimensions['F'].width=12

    # 시트3: 타프토이즈 추천
    ws3=wb.create_sheet("타프토이즈 추천")
    ws3.merge_cells("A1:E1")
    c=ws3.cell(row=1,column=1,value=f"타프토이즈 미취급 제품 추천 — {seller}")
    c.fill=mf("7C3AED"); c.font=mft("FFFFFF",True,12); c.alignment=ctr; ws3.row_dimensions[1].height=26
    SP={'아치/모빌':'인스타 감성↑, 구매 결정 빠름','트래블토이':'유모차/카시트 필수, 재구매율 높음',
        '비지북':'교육적 가치, 선물용 인기','큐브':'다기능 가성비, 1+1 구성 가능',
        '워터매트':'계절성 높음, 여름 전 선주문','액티비티짐':'고마진, 출산선물 1순위',
        '터미타임':'소아과 추천, 안전 강조','감각 장난감':'6개월부터 사용, 반복구매',
        '인형/뮤지컬':'수면 루틴, 감성 구매','카시트 장난감':'카시트 구매 시 ADD-ON'}
    for ci,h in enumerate(['제품명','카테고리','소비자가(원)','제품 특징','영업 포인트'],1):
        c=ws3.cell(row=2,column=ci,value=h); c.fill=mf("F2F2F2"); c.font=mft("595959",True,10); c.alignment=ctr; c.border=bdr
    for ri,u in enumerate(unsold,3):
        for ci,val in enumerate([u['name'],u.get('category',''),u.get('price',0),u.get('desc',''),SP.get(u.get('category',''),'')],1):
            c=ws3.cell(row=ri,column=ci,value=val); c.border=bdr
            if ri%2==0: c.fill=mf("FAF5FF")
            if ci==3: c.number_format='#,##0'; c.alignment=rgt
    ws3.column_dimensions['A'].width=30; ws3.column_dimensions['B'].width=16
    ws3.column_dimensions['C'].width=14; ws3.column_dimensions['D'].width=42
    ws3.column_dimensions['E'].width=32

    # 시트4: 주별 추이
    ws4=wb.create_sheet("주별 추이")
    ws4.merge_cells("A1:D1")
    c=ws4.cell(row=1,column=1,value=f"{seller} — 주별 판매 추이")
    c.fill=mf("1E3A5F"); c.font=mft("FFFFFF",True,12); c.alignment=ctr; ws4.row_dimensions[1].height=26
    for ci,h in enumerate(['주차','기간','판매금액(원)','판매수량'],1):
        c=ws4.cell(row=2,column=ci,value=h); c.fill=mf("F2F2F2"); c.font=mft("595959",True,10); c.alignment=ctr; c.border=bdr
    for ri,r in enumerate(weekly_raw,3):
        try:
            d=dt2.strptime(r[1],"%Y-%m-%d"); sun=d-timedelta(days=(d.weekday()+1)%7)
            sat=sun+timedelta(days=6); period=f"{sun.strftime('%m/%d')}~{sat.strftime('%m/%d')}"
        except: period=r[0]
        for ci,val in enumerate([f"{ri-2}주차",period,r[2],r[3]],1):
            c=ws4.cell(row=ri,column=ci,value=val); c.border=bdr
            if ri%2==0: c.fill=mf("FAFAFA")
            if ci==3: c.number_format='#,##0'; c.alignment=rgt
            if ci in(1,2): c.alignment=ctr
    for ci,w in zip('ABCD',[10,16,16,10]): ws4.column_dimensions[ci].width=w

    buf=io.BytesIO(); wb.save(buf); buf.seek(0)
    return send_file(buf,mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                     as_attachment=True,download_name=f"영업스크립트_{seller}_{year}.xlsx")

def make_xlsx(headers, rows_data, sheet_name="데이터"):
    wb = openpyxl.Workbook()
    ws = wb.active; ws.title = sheet_name
    hdr_fill = PatternFill(start_color="4F46E5", end_color="4F46E5", fill_type="solid")
    hdr_font = Font(color="FFFFFF", bold=True, size=11)
    thin = Side(style='thin', color='E5E7EB')
    bdr  = Border(left=thin, right=thin, top=thin, bottom=thin)
    for col, hdr in enumerate(headers, 1):
        c = ws.cell(row=1, column=col, value=hdr)
        c.fill=hdr_fill; c.font=hdr_font
        c.alignment=Alignment(horizontal="center",vertical="center"); c.border=bdr
    ws.row_dimensions[1].height = 24
    even_fill = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")
    for ri, row in enumerate(rows_data, 2):
        for ci, val in enumerate(row, 1):
            c = ws.cell(row=ri, column=ci, value=val)
            c.border = bdr
            if ri % 2 == 0: c.fill = even_fill
            if isinstance(val,(int,float)) and ci > 1:
                c.alignment = Alignment(horizontal="right")
                if '금액' in headers[ci-1]: c.number_format = '#,##0'
    for col in ws.columns:
        ml = max((len(str(c.value or '')) for c in col), default=8)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(ml+4, 40)
    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf

@app.route("/api/export/xlsx/monthly")
@login_required
def export_xlsx_monthly():
    """브랜드별 정리 엑셀 형식 — 매장×월별×브랜드별 금액/수량"""
    year   = request.args.get("year",   str(datetime.now().year))
    month  = request.args.get("month",  "")
    seller = request.args.get("seller", "").strip()
    conn   = get_db()

    # 브랜드 고정 순서 사용
    brands = BRAND_ORDER  # ['줄즈','레카로','ABC디자인','원더폴드','카오스','엔픽스','타프토이즈']

    # 월 목록
    if month:
        months = [int(month)]
    else:
        months_raw = conn.execute(
            f"SELECT DISTINCT CAST(strftime('%m',sale_date) AS INTEGER) m "
            f"FROM sales_data WHERE sale_date LIKE '{year}%' AND sale_date!='' "
            f"ORDER BY m").fetchall()
        months = [r[0] for r in months_raw] or list(range(1,13))

    # 매장 목록
    seller_cond   = "AND real_seller=?" if seller else ""
    seller_params = [seller] if seller else []
    sellers_raw = conn.execute(
        f"SELECT DISTINCT real_seller FROM sales_data "
        f"WHERE real_seller!='' AND sale_date LIKE '{year}%' {seller_cond} "
        f"ORDER BY real_seller", seller_params).fetchall()
    sellers_list = [r[0] for r in sellers_raw]

    def brand_key(nm):
        nm=(nm or '').replace('_',' ').lower()
        if '베이비하우스' in nm: return (0,nm)
        if '링크맘' in nm:       return (1,nm)
        if '베이비파크' in nm:   return (2,nm)
        if '베네피아' in nm:     return (3,nm)
        return (9,nm)
    sellers_list.sort(key=brand_key)

    # 데이터 조회 — 매장×월×품목그룹 (item_group, item_name 모두 가져와 remap)
    data_rows = conn.execute(
        f"""SELECT real_seller, CAST(strftime('%m',sale_date) AS INTEGER) mo,
            item_group, item_name, SUM(total) total, SUM(quantity) qty
            FROM sales_data
            WHERE real_seller!='' AND sale_date LIKE '{year}%' AND sale_date!=''
            {seller_cond}
            GROUP BY real_seller, mo, item_group, item_name""",
        seller_params).fetchall()

    # 인덱스: {(seller, month, brand): {total, qty}} — remap_group 적용
    idx = {}
    for r in data_rows:
        brand = remap_group(r[2], r[3])  # item_group, item_name
        if not brand or brand == '기타': continue
        key = (r[0], r[1], brand)
        if key not in idx:
            idx[key] = {'total': 0, 'qty': 0}
        idx[key]['total'] += r[4] or 0
        idx[key]['qty']   += r[5] or 0

    # ── openpyxl 빌드 ──
    wb = openpyxl.Workbook()

    # ── 스타일 팔레트 ──
    WHITE       = "FFFFFF"
    GRAY_LIGHT  = "F2F2F2"   # 행 2~3 (업체구분, 헤더)
    GRAY_HDR    = "E8E8E8"   # 데이터 헤더 행
    BORDER_CLR  = "BFBFBF"   # 매장정보 열 테두리
    FONT_BLACK  = "000000"
    FONT_GRAY   = "595959"

    def mf(h):  return PatternFill(start_color=h,end_color=h,fill_type="solid")
    def mft(h,bold=False,sz=10): return Font(color=h,bold=bold,size=sz)
    thin_bdr = Side(style='thin', color=BORDER_CLR)
    no_bdr   = Side(style=None)
    bdr_left  = Border(left=thin_bdr,right=thin_bdr,top=thin_bdr,bottom=thin_bdr)  # 매장정보 열
    bdr_none  = Border(left=no_bdr,right=no_bdr,top=no_bdr,bottom=no_bdr)          # 브랜드 데이터 열
    center   = Alignment(horizontal="center",vertical="center")
    right    = Alignment(horizontal="right",  vertical="center")
    num_fmt  = '#,##0'

    col_start = 4  # A=업체구분, B=거래처명, C=실적용, D부터 데이터

    def build_sheet(wb_ref, title, field):
        """금액 또는 수량 시트 생성"""
        ws = wb_ref.create_sheet(title) if title != "브랜드별 금액" else wb_ref.active
        if title == "브랜드별 금액": ws.title = title
        total_cols = 3 + len(months)*(len(brands)+1)

        # 행1: 타이틀 — 흰색 배경
        ws.merge_cells(f"A1:{get_column_letter(total_cols)}1")
        c = ws.cell(row=1,column=1,value=f"오프라인 {'판매금액' if field=='total' else '판매수량'} 브랜드별 정리_{year}")
        c.fill=mf(WHITE); c.font=mft(FONT_BLACK,True,12); c.alignment=center
        ws.row_dimensions[1].height=26

        # 행2: 월 헤더 — 연한 회색
        ws.cell(row=2,column=1,value="업체구분").fill=mf(GRAY_LIGHT)
        ws.cell(row=2,column=2,value="거래처명").fill=mf(GRAY_LIGHT)
        ws.cell(row=2,column=3,value="실적용거래처명").fill=mf(GRAY_LIGHT)
        for ci in range(1,4):
            c=ws.cell(row=2,column=ci)
            c.font=mft(FONT_GRAY,True,10); c.alignment=center; c.border=bdr_left
        ws.merge_cells("A2:A3"); ws.merge_cells("B2:B3"); ws.merge_cells("C2:C3")

        col=col_start
        for mo in months:
            span=len(brands)+1; end_col=col+span-1
            ws.merge_cells(f"{get_column_letter(col)}2:{get_column_letter(end_col)}2")
            c=ws.cell(row=2,column=col,value=f"{year}_{mo:02d}")
            c.fill=mf(GRAY_LIGHT); c.font=mft(FONT_GRAY,True,11); c.alignment=center
            c.border=bdr_left  # 월 구분선만
            col+=span

        # 행3: 브랜드 헤더 — 연한 회색
        col=col_start
        for mo in months:
            for b in brands:
                c=ws.cell(row=3,column=col,value=f"{b}{'금액' if field=='total' else '수량'}")
                c.fill=mf(GRAY_LIGHT); c.font=mft(FONT_GRAY,False,9); c.alignment=center
                c.border=bdr_none; col+=1
            c=ws.cell(row=3,column=col,value="합계")
            c.fill=mf(GRAY_LIGHT); c.font=mft(FONT_GRAY,True,9); c.alignment=center
            c.border=bdr_left; col+=1
        ws.row_dimensions[2].height=18; ws.row_dimensions[3].height=16

        # 행4: 빈 헤더행 (흰색 구분)
        for ci in range(1,total_cols+1):
            c=ws.cell(row=4,column=ci,value="")
            c.fill=mf(WHITE); c.border=bdr_none
        ws.row_dimensions[4].height=4

        # 컬럼 너비
        ws.column_dimensions['A'].width=12
        ws.column_dimensions['B'].width=22
        ws.column_dimensions['C'].width=24
        for mo_i in range(len(months)):
            for b_i in range(len(brands)+1):
                ws.column_dimensions[get_column_letter(col_start+mo_i*(len(brands)+1)+b_i)].width=11

        # 데이터 행 (5행부터)
        branch_group = {}
        try:
            bg=get_db()
            for r in bg.execute("SELECT name,note FROM branches").fetchall():
                branch_group[r[0]]=r[1] or ''
            bg.close()
        except: pass

        prev_grp=None
        for ri,s in enumerate(sellers_list,5):
            grp=branch_group.get(s,''); gv=grp if grp!=prev_grp else ''; prev_grp=grp
            # 매장정보 열 (A~C) — 테두리 있음
            for ci,val in enumerate([gv,s,s],1):
                c=ws.cell(row=ri,column=ci,value=val)
                c.fill=mf(WHITE); c.border=bdr_left; c.font=mft(FONT_BLACK,False,10)
                if ci==1: c.font=mft(FONT_GRAY,False,10)

            # 브랜드 데이터 열 — 테두리 없음
            col=col_start
            for mo in months:
                month_total=0
                for b in brands:
                    val=idx.get((s,mo,b),{}).get(field,0)
                    month_total+=val
                    c=ws.cell(row=ri,column=col,value=val if val else 0)
                    c.fill=mf(WHITE); c.border=bdr_none; c.alignment=right
                    c.number_format=num_fmt; c.font=mft(FONT_BLACK,False,10); col+=1
                # 합계 열 — 왼쪽에 세로선 (월 구분)
                c=ws.cell(row=ri,column=col,value=month_total)
                c.fill=mf(WHITE); c.font=mft(FONT_BLACK,True,10)
                c.border=Border(left=thin_bdr,right=no_bdr,top=no_bdr,bottom=no_bdr)
                c.alignment=right; c.number_format=num_fmt; col+=1

        # 합계 행 — 흰색
        tot_row=len(sellers_list)+5
        for ci,val in enumerate(["합계","",""],1):
            c=ws.cell(row=tot_row,column=ci,value=val)
            c.fill=mf(WHITE); c.border=bdr_left; c.font=mft(FONT_BLACK,True,10)
        col=col_start
        for mo in months:
            for b in brands:
                total_b=sum(idx.get((s,mo,b),{}).get(field,0) for s in sellers_list)
                c=ws.cell(row=tot_row,column=col,value=total_b)
                c.fill=mf(WHITE); c.border=bdr_none; c.alignment=right
                c.number_format=num_fmt; c.font=mft(FONT_BLACK,True,10); col+=1
            grand=sum(idx.get((s,mo,b),{}).get(field,0) for s in sellers_list for b in brands)
            c=ws.cell(row=tot_row,column=col,value=grand)
            c.fill=mf(WHITE); c.font=mft(FONT_BLACK,True,10)
            c.border=Border(left=thin_bdr,right=no_bdr,top=no_bdr,bottom=no_bdr)
            c.alignment=right; c.number_format=num_fmt; col+=1

        ws.freeze_panes="D5"
        return ws

    build_sheet(wb, "브랜드별 금액", "total")
    build_sheet(wb, "브랜드별 수량", "qty")

    # ── 시트3: 제품별 상세 ──
    ws3=wb.create_sheet("제품별 상세")
    params2=[f"{year}%"]; conds2=["sale_date LIKE ?","sale_date!=''"]
    if seller: conds2.append("real_seller=?"); params2.append(seller)
    if month:  conds2.append(f"strftime('%m',sale_date)='{month.zfill(2)}'")
    raw_items=[dict(r) for r in conn.execute(f"""
        SELECT item_group,item_name,SUM(quantity) qty,SUM(total) total,COUNT(*) cnt
        FROM sales_data WHERE {' AND '.join(conds2)}
        GROUP BY item_name ORDER BY item_group,total DESC""",params2).fetchall()]
    conn.close()

    merged={}
    for r in raw_items:
        nn=normalize_item_name(r['item_name']); ng=remap_group(r['item_group'],r['item_name'])
        if not ng or ng=='기타': continue
        key=(ng,nn)
        if key not in merged: merged[key]={'item_group':ng,'item_name':nn,'qty':0,'total':0,'cnt':0}
        merged[key]['qty']+=r['qty']; merged[key]['total']+=r['total']; merged[key]['cnt']+=r['cnt']
    sorted_items=sorted(merged.values(),key=lambda x:(get_group_sort_key(x['item_group']),-x['total']))

    thin3=Side(style='thin',color='E0E0E0')
    bdr3=Border(left=thin3,right=thin3,top=thin3,bottom=thin3)
    for ci,h in enumerate(['품목그룹','제품명','판매건수','판매수량','합계금액(원)'],1):
        c=ws3.cell(row=1,column=ci,value=h)
        c.fill=mf(GRAY_LIGHT); c.font=mft(FONT_GRAY,True,10); c.alignment=center; c.border=bdr3
    ws3.row_dimensions[1].height=20
    for ri,r in enumerate(sorted_items,2):
        for ci,val in enumerate([r['item_group'],r['item_name'],r['cnt'],r['qty'],r['total']],1):
            c=ws3.cell(row=ri,column=ci,value=val); c.border=bdr3
            if ri%2==0: c.fill=mf("FAFAFA")
            if ci>2: c.alignment=right
            if ci==5 and isinstance(val,int): c.number_format=num_fmt
    for col in ws3.columns:
        ml=max((len(str(c.value or '')) for c in col),default=8)
        ws3.column_dimensions[get_column_letter(col[0].column)].width=min(ml+3,35)

    buf=io.BytesIO(); wb.save(buf); buf.seek(0)
    fname=f"오프라인_브랜드별정리_{year}{'_'+month+'월' if month else ''}.xlsx"
    return send_file(buf,mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                     as_attachment=True,download_name=fname)

    # 월 목록
    if month:
        months = [int(month)]
    else:
        months_raw = conn.execute(
            f"SELECT DISTINCT CAST(strftime('%m',sale_date) AS INTEGER) m "
            f"FROM sales_data WHERE sale_date LIKE '{year}%' AND sale_date!='' "
            f"ORDER BY m").fetchall()
        months = [r[0] for r in months_raw] or list(range(1,13))

    # 매장 목록 (브랜드 정렬)
    seller_cond = "AND real_seller=?" if seller else ""
    seller_params = [seller] if seller else []
    sellers_raw = conn.execute(
        f"SELECT DISTINCT real_seller FROM sales_data "
        f"WHERE real_seller!='' AND sale_date LIKE '{year}%' {seller_cond} "
        f"ORDER BY real_seller", seller_params).fetchall()
    sellers_list = [r[0] for r in sellers_raw]

    def brand_key(nm):
        nm = (nm or '').replace('_',' ').lower()
        if '베이비하우스' in nm: return (0, nm)
        if '링크맘' in nm: return (1, nm)
        if '베이비파크' in nm: return (2, nm)
        if '베네피아' in nm: return (3, nm)
        return (9, nm)
    sellers_list.sort(key=brand_key)

    # 데이터 조회 — 매장×월×품목그룹 집계
    data_rows = conn.execute(
        f"""SELECT real_seller, CAST(strftime('%m',sale_date) AS INTEGER) mo,
            item_group, SUM(total) total, SUM(quantity) qty
            FROM sales_data
            WHERE real_seller!='' AND sale_date LIKE '{year}%' AND sale_date!=''
            {seller_cond}
            GROUP BY real_seller, mo, item_group""",
        seller_params).fetchall()

    # 인덱스: {(seller, month, group): {total, qty}}
    idx = {}
    for r in data_rows:
        key = (r[0], r[1], r[2])
        idx[key] = {'total': r[3] or 0, 'qty': r[4] or 0}

    # ── openpyxl 빌드 ──
    wb = openpyxl.Workbook()

    # 스타일
    def mk_fill(hex): return PatternFill(start_color=hex,end_color=hex,fill_type="solid")
    def mk_font(hex, bold=True, sz=10): return Font(color=hex,bold=bold,size=sz)
    thin = Side(style='thin', color='D1D5DB')
    bdr  = Border(left=thin,right=thin,top=thin,bottom=thin)
    num_fmt = '#,##0'
    center = Alignment(horizontal="center",vertical="center")
    right  = Alignment(horizontal="right",vertical="center")

    fill_main  = mk_fill("4F46E5")   # 진한 인디고 — 메인 헤더
    fill_month = mk_fill("818CF8")   # 연한 인디고 — 월 헤더
    fill_brand = mk_fill("C7D2FE")   # 더 연한 — 브랜드 헤더
    fill_total = mk_fill("EEF2FF")   # 합계 열
    fill_even  = mk_fill("F9FAFB")
    fill_group_bh  = mk_fill("FFF7ED")  # 베이비하우스
    fill_group_lm  = mk_fill("F0FDF4")  # 링크맘
    fill_group_etc = mk_fill("FFFFFF")

    def group_fill(nm):
        nm=(nm or '').replace('_',' ').lower()
        if '베이비하우스' in nm: return fill_group_bh
        if '링크맘' in nm: return fill_group_lm
        return fill_group_etc

    # ── 시트1: 브랜드별 금액 ──
    ws1 = wb.active; ws1.title = "브랜드별 금액"

    # 행1: 타이틀
    ws1.merge_cells(f"A1:{get_column_letter(4 + len(months)*(len(brands)+1))}1")
    c=ws1.cell(row=1,column=1,value=f"오프라인 판매금액 브랜드별 정리_{year}")
    c.fill=fill_main; c.font=mk_font("FFFFFF",True,12)
    c.alignment=center; ws1.row_dimensions[1].height=28

    # 행2: 월 헤더 (span 브랜드+합계)
    fixed_cols = 3  # 업체구분, 거래처명, 실적용거래처명
    col_start = fixed_cols + 1
    ws1.cell(row=2,column=1,value="업체구분").fill=fill_brand
    ws1.cell(row=2,column=2,value="거래처명").fill=fill_brand
    ws1.cell(row=2,column=3,value="실적용거래처명").fill=fill_brand
    for ci in range(1,4):
        ws1.cell(row=2,column=ci).font=mk_font("374151",True,10)
        ws1.cell(row=2,column=ci).alignment=center
        ws1.cell(row=2,column=ci).border=bdr
    ws1.merge_cells(f"A2:A3"); ws1.merge_cells(f"B2:B3"); ws1.merge_cells(f"C2:C3")

    col = col_start
    for mo in months:
        span = len(brands)+1
        end_col = col+span-1
        ws1.merge_cells(f"{get_column_letter(col)}2:{get_column_letter(end_col)}2")
        c=ws1.cell(row=2,column=col,value=f"{year}_{mo:02d}")
        c.fill=fill_month; c.font=mk_font("FFFFFF",True,11); c.alignment=center; c.border=bdr
        col += span

    # 행3: 브랜드명 헤더
    col = col_start
    for mo in months:
        for b in brands:
            c=ws1.cell(row=3,column=col,value=f"{b}금액")
            c.fill=fill_brand; c.font=mk_font("374151",True,9); c.alignment=center; c.border=bdr
            col+=1
        c=ws1.cell(row=3,column=col,value="합계")
        c.fill=mk_fill("A5B4FC"); c.font=mk_font("1E1B4B",True,9); c.alignment=center; c.border=bdr
        col+=1
    ws1.row_dimensions[2].height=20; ws1.row_dimensions[3].height=18

    # 고정 컬럼 너비
    ws1.column_dimensions['A'].width = 12
    ws1.column_dimensions['B'].width = 20
    ws1.column_dimensions['C'].width = 22

    # 데이터 행
    # 업체구분 파악 (branches 테이블)
    branch_group = {}
    try:
        bg_conn = get_db()
        for r in bg_conn.execute("SELECT name, note FROM branches").fetchall():
            branch_group[r[0]] = r[1] or ''
        bg_conn.close()
    except: pass

    prev_group = None
    for ri, s in enumerate(sellers_list, 4):
        grp = branch_group.get(s,'')
        row_fill = group_fill(s)

        # 업체구분 — 그룹 변경 시만 표시
        grp_val = grp if grp != prev_group else ''
        prev_group = grp

        ws1.cell(row=ri,column=1,value=grp_val).fill=row_fill
        ws1.cell(row=ri,column=2,value=s).fill=row_fill
        ws1.cell(row=ri,column=3,value=s).fill=row_fill
        for ci in range(1,4):
            ws1.cell(row=ri,column=ci).border=bdr
            ws1.cell(row=ri,column=ci).font=Font(size=10)

        col = col_start
        for mo in months:
            month_total = 0
            for b in brands:
                val = idx.get((s,mo,b),{}).get('total',0)
                month_total += val
                c=ws1.cell(row=ri,column=col,value=val if val else 0)
                c.fill=row_fill; c.border=bdr; c.alignment=right
                c.number_format=num_fmt; c.font=Font(size=10)
                col+=1
            # 합계
            c=ws1.cell(row=ri,column=col,value=month_total)
            c.fill=fill_total; c.border=bdr; c.alignment=right
            c.number_format=num_fmt; c.font=Font(bold=True,size=10)
            col+=1

    # 합계 행
    tot_row = len(sellers_list)+4
    ws1.cell(row=tot_row,column=1,value="합계").fill=fill_main
    ws1.cell(row=tot_row,column=2,value="").fill=fill_main
    ws1.cell(row=tot_row,column=3,value="").fill=fill_main
    for ci in range(1,4): ws1.cell(row=tot_row,column=ci).font=mk_font("FFFFFF",True,10)
    col = col_start
    for mo in months:
        for b in brands:
            total_b = sum(idx.get((s,mo,b),{}).get('total',0) for s in sellers_list)
            c=ws1.cell(row=tot_row,column=col,value=total_b)
            c.fill=fill_main; c.font=mk_font("FFFFFF",True,10)
            c.border=bdr; c.alignment=right; c.number_format=num_fmt; col+=1
        grand = sum(idx.get((s,mo,b),{}).get('total',0) for s in sellers_list for b in brands)
        c=ws1.cell(row=tot_row,column=col,value=grand)
        c.fill=mk_fill("312E81"); c.font=mk_font("FFFFFF",True,10)
        c.border=bdr; c.alignment=right; c.number_format=num_fmt; col+=1

    # 데이터 컬럼 너비
    brand_col_width = 11
    for mo_i in range(len(months)):
        for b_i in range(len(brands)+1):
            col_idx = col_start + mo_i*(len(brands)+1) + b_i
            ws1.column_dimensions[get_column_letter(col_idx)].width = brand_col_width

    ws1.freeze_panes = "D4"

    # ── 시트2: 브랜드별 수량 ──
    ws2 = wb.create_sheet("브랜드별 수량")
    # 동일 구조, qty
    ws2.merge_cells(f"A1:{get_column_letter(4 + len(months)*(len(brands)+1))}1")
    c=ws2.cell(row=1,column=1,value=f"오프라인 판매수량 브랜드별 정리_{year}")
    c.fill=mk_fill("065F46"); c.font=mk_font("FFFFFF",True,12); c.alignment=center
    ws2.row_dimensions[1].height=28

    for ci in range(1,4):
        ws2.cell(row=2,column=ci,value=["업체구분","거래처명","실적용거래처명"][ci-1])
        ws2.cell(row=2,column=ci).fill=fill_brand; ws2.cell(row=2,column=ci).font=mk_font("374151",True,10)
        ws2.cell(row=2,column=ci).alignment=center; ws2.cell(row=2,column=ci).border=bdr
    ws2.merge_cells(f"A2:A3"); ws2.merge_cells(f"B2:B3"); ws2.merge_cells(f"C2:C3")

    col=col_start
    for mo in months:
        span=len(brands)+1; end_col=col+span-1
        ws2.merge_cells(f"{get_column_letter(col)}2:{get_column_letter(end_col)}2")
        c=ws2.cell(row=2,column=col,value=f"{year}_{mo:02d}")
        c.fill=mk_fill("065F46"); c.font=mk_font("FFFFFF",True,11); c.alignment=center; c.border=bdr
        for b in brands:
            c2=ws2.cell(row=3,column=col,value=f"{b}수량")
            c2.fill=mk_fill("D1FAE5"); c2.font=mk_font("374151",True,9); c2.alignment=center; c2.border=bdr; col+=1
        c2=ws2.cell(row=3,column=col,value="수량합계")
        c2.fill=mk_fill("6EE7B7"); c2.font=mk_font("065F46",True,9); c2.alignment=center; c2.border=bdr; col+=1
    ws2.row_dimensions[2].height=20; ws2.row_dimensions[3].height=18
    ws2.column_dimensions['A'].width=12; ws2.column_dimensions['B'].width=20; ws2.column_dimensions['C'].width=22

    prev_group=None
    for ri,s in enumerate(sellers_list,4):
        grp=branch_group.get(s,''); gv=grp if grp!=prev_group else ''; prev_group=grp
        ws2.cell(row=ri,column=1,value=gv); ws2.cell(row=ri,column=2,value=s); ws2.cell(row=ri,column=3,value=s)
        for ci in range(1,4):
            ws2.cell(row=ri,column=ci).border=bdr; ws2.cell(row=ri,column=ci).font=Font(size=10)
        col=col_start
        for mo in months:
            mt=0
            for b in brands:
                val=idx.get((s,mo,b),{}).get('qty',0); mt+=val
                c=ws2.cell(row=ri,column=col,value=val if val else 0)
                c.border=bdr; c.alignment=right; c.number_format=num_fmt; c.font=Font(size=10); col+=1
            c=ws2.cell(row=ri,column=col,value=mt)
            c.fill=mk_fill("ECFDF5"); c.border=bdr; c.alignment=right
            c.number_format=num_fmt; c.font=Font(bold=True,size=10); col+=1

    for mo_i in range(len(months)):
        for b_i in range(len(brands)+1):
            ws2.column_dimensions[get_column_letter(col_start+mo_i*(len(brands)+1)+b_i)].width=brand_col_width
    ws2.freeze_panes="D4"

    # ── 시트3: 제품별 상세 ──
    ws3 = wb.create_sheet("제품별 상세")
    params2=[f"{year}%"]; conds2=["sale_date LIKE ?","sale_date!=''"]
    if seller: conds2.append("real_seller=?"); params2.append(seller)
    if month:  conds2.append(f"strftime('%m',sale_date)='{month.zfill(2)}'")
    raw_items=[dict(r) for r in conn.execute(f"""
        SELECT item_group,item_name,SUM(quantity) qty,SUM(total) total,COUNT(*) cnt
        FROM sales_data WHERE {' AND '.join(conds2)}
        GROUP BY item_name ORDER BY item_group,total DESC""",params2).fetchall()]
    conn.close()  # 모든 쿼리 완료 후 닫기

    merged={}
    for r in raw_items:
        nn=normalize_item_name(r['item_name']); ng=remap_group(r['item_group'],r['item_name'])
        key=(ng,nn)
        if key not in merged: merged[key]={'item_group':ng,'item_name':nn,'qty':0,'total':0,'cnt':0}
        merged[key]['qty']+=r['qty']; merged[key]['total']+=r['total']; merged[key]['cnt']+=r['cnt']
    sorted_items=sorted(merged.values(),key=lambda x:(get_group_sort_key(x['item_group']),-x['total']))

    hdr3_fill=mk_fill("6366F1")
    item_hdrs=['품목그룹','제품명','판매건수','판매수량','합계금액(원)']
    for ci,h in enumerate(item_hdrs,1):
        c=ws3.cell(row=1,column=ci,value=h); c.fill=hdr3_fill
        c.font=mk_font("FFFFFF",True,10); c.alignment=center; c.border=bdr
    ws3.row_dimensions[1].height=22
    for ri,r in enumerate(sorted_items,2):
        vals=[r['item_group'],r['item_name'],r['cnt'],r['qty'],r['total']]
        for ci,val in enumerate(vals,1):
            c=ws3.cell(row=ri,column=ci,value=val); c.border=bdr
            if ri%2==0: c.fill=mk_fill("F9FAFB")
            if ci>2: c.alignment=right
            if ci==5 and isinstance(val,int): c.number_format=num_fmt
    for col in ws3.columns:
        ml=max((len(str(c.value or '')) for c in col),default=8)
        ws3.column_dimensions[get_column_letter(col[0].column)].width=min(ml+3,35)

    buf=io.BytesIO(); wb.save(buf); buf.seek(0)
    fname=f"오프라인_브랜드별정리_{year}{'_'+month+'월' if month else ''}.xlsx"
    return send_file(buf,mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                     as_attachment=True,download_name=fname)

@app.route("/api/export/xlsx/weekly")
@login_required
def export_xlsx_weekly():
    from datetime import datetime as dt2, timedelta
    year=request.args.get("year",str(datetime.now().year))
    month=request.args.get("month",""); seller=request.args.get("seller","").strip()
    qp=["sale_date != ''"];pp=[]
    if month: qp.append("sale_date LIKE ?");pp.append(f"{year}-{month.zfill(2)}%")
    else:     qp.append("sale_date LIKE ?");pp.append(f"{year}%")
    if seller: qp.append("real_seller = ?");pp.append(seller)
    conn=get_db()
    rows=[dict(r) for r in conn.execute(f"""
        SELECT strftime('%Y-%W',sale_date) AS week_key,
               COUNT(*) cnt,SUM(quantity) qty,SUM(total) total,MIN(sale_date) AS min_date
        FROM sales_data WHERE {' AND '.join(qp)} AND sale_date!=''
        GROUP BY week_key ORDER BY week_key""",pp).fetchall()]

    def wr(ds):
        d=dt2.strptime(ds,"%Y-%m-%d"); sun=d-timedelta(days=(d.weekday()+1)%7)
        return sun.strftime("%Y-%m-%d"),(sun+timedelta(days=6)).strftime("%Y-%m-%d")

    for r in rows:
        try: r['ws'],r['we']=wr(r['min_date'])
        except: r['ws']=r['we']=''

    # 제품별 상세 (주차별 — 색상 통합, 그룹 정규화)
    items_by_week={}
    for r in rows:
        wk=r['week_key']
        w_conds=["sale_date!=''", "strftime('%Y-%W',sale_date)=?"]
        w_params=[wk]
        if seller: w_conds.append("real_seller=?"); w_params.append(seller)
        irows=conn.execute(f"""SELECT item_group, item_name,
            SUM(quantity) qty, SUM(total) total, COUNT(*) cnt
            FROM sales_data WHERE {' AND '.join(w_conds)}
            GROUP BY item_name ORDER BY total DESC""", w_params).fetchall()
        # 색상 통합
        merged_w = {}
        for ir in irows:
            d=dict(ir)
            norm_name  = normalize_item_name(d['item_name'])
            norm_group = remap_group(d['item_group'], d['item_name'])
            key=(norm_group,norm_name)
            if key not in merged_w:
                merged_w[key]={'item_group':norm_group,'item_name':norm_name,'qty':0,'total':0,'cnt':0}
            merged_w[key]['qty']   += d['qty']
            merged_w[key]['total'] += d['total']
            merged_w[key]['cnt']   += d['cnt']
        items_by_week[wk]=sorted(merged_w.values(),
            key=lambda x:(get_group_sort_key(x['item_group']),-x['total']))
    conn.close()

    wb=openpyxl.Workbook()
    thin=Side(style='thin',color='E5E7EB'); bdr=Border(left=thin,right=thin,top=thin,bottom=thin)
    even=PatternFill(start_color="F9FAFB",end_color="F9FAFB",fill_type="solid")

    def hc(cell,color="4F46E5"):
        cell.fill=PatternFill(start_color=color,end_color=color,fill_type="solid")
        cell.font=Font(color="FFFFFF",bold=True,size=10)
        cell.alignment=Alignment(horizontal="center",vertical="center"); cell.border=bdr

    # 시트1: 주별 요약
    ws1=wb.active; ws1.title="주별 요약"
    hdrs=['주차','시작일(일)','종료일(토)','판매건수','판매수량','판매금액(원)']
    for ci,h in enumerate(hdrs,1): hc(ws1.cell(row=1,column=ci,value=h))
    for i,r in enumerate(rows,2):
        for ci,v in enumerate([f"{i-1}주차",r['ws'],r['we'],r['cnt'],r['qty'],r['total']],1):
            c=ws1.cell(row=i,column=ci,value=v); c.border=bdr
            if i%2==0: c.fill=even
            if ci>3: c.alignment=Alignment(horizontal="right")
            if ci==6 and isinstance(v,int): c.number_format='#,##0'

    # 시트2: 제품별 상세 (색상 통합, 평균단가/공급가/부가세 제외)
    ws2=wb.create_sheet("제품별 상세")
    item_hdrs=['주차','기간','품목그룹','제품명','판매건수','판매수량','판매금액(원)']
    for ci,h in enumerate(item_hdrs,1): hc(ws2.cell(row=1,column=ci,value=h),"6366F1")
    ri=2
    for i,r in enumerate(rows):
        for item in items_by_week.get(r['week_key'],[]):
            vals=[f"{i+1}주차",f"{r['ws']}~{r['we']}",item.get('item_group',''),
                  item.get('item_name',''),item.get('cnt',0),item.get('qty',0),item.get('total',0)]
            for ci,v in enumerate(vals,1):
                c=ws2.cell(row=ri,column=ci,value=v); c.border=bdr
                if ri%2==0: c.fill=even
                if ci>4: c.alignment=Alignment(horizontal="right")
                if ci==7 and isinstance(v,int): c.number_format='#,##0'
            ri+=1

    for ws in [ws1,ws2]:
        for col in ws.columns:
            ml=max((len(str(c.value or '')) for c in col),default=8)
            ws.column_dimensions[get_column_letter(col[0].column)].width=min(ml+3,35)

    buf=io.BytesIO(); wb.save(buf); buf.seek(0)
    fname=f"주별실적_{year}{'_'+month+'월' if month else ''}.xlsx"
    return send_file(buf,mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                     as_attachment=True,download_name=fname)

@app.route("/api/export/xlsx/ranking")
@login_required
def export_xlsx_ranking():
    year=request.args.get("year",str(datetime.now().year)); month=request.args.get("month","")
    date_cond=f"{year}-{month.zfill(2)}%" if month else f"{year}%"
    conn=get_db()
    rows=[dict(r) for r in conn.execute("""SELECT real_seller AS seller_name,
        COUNT(*) cnt,SUM(total) total,SUM(quantity) qty
        FROM sales_data WHERE sale_date LIKE ? AND real_seller!=''
        GROUP BY real_seller ORDER BY total DESC""",(date_cond,)).fetchall()]
    conn.close()
    hdrs=['순위','매장명','판매금액(원)','판매건수','판매수량']
    data=[[i+1,r['seller_name'],r['total'],r['cnt'],r['qty']] for i,r in enumerate(rows)]
    buf=make_xlsx(hdrs,data,"매출순위")
    fname=f"매출순위_{year}{'_'+month+'월' if month else ''}.xlsx"
    return send_file(buf,mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                     as_attachment=True,download_name=fname)

# ── 제품별 관리 API ────────────────────────────────
@app.route("/api/products/groups")
@login_required
def api_product_groups():
    conn = get_db()
    groups = [dict(r) for r in conn.execute("""
        SELECT item_group, COUNT(DISTINCT item_name) item_cnt,
               SUM(quantity) qty, SUM(total) total
        FROM sales_data WHERE item_group != '' AND item_group IS NOT NULL
        GROUP BY item_group ORDER BY total DESC""").fetchall()]
    conn.close()
    return jsonify(groups)

@app.route("/api/products/items")
@login_required
def api_product_items():
    group=request.args.get("group",""); seller=request.args.get("seller","").strip()
    year=request.args.get("year",str(datetime.now().year)); month=request.args.get("month","")
    conn=get_db()
    date_cond=f"{year}-{month.zfill(2)}%" if month else f"{year}%"
    params=[date_cond]; conds=["sale_date LIKE ?","sale_date != ''"]
    if group:  conds.append("item_group=?");  params.append(group)
    if seller: conds.append("real_seller=?"); params.append(seller)
    rows=[dict(r) for r in conn.execute(f"""SELECT item_name,item_group,item_code,
        SUM(quantity) qty,AVG(unit_price) avg_price,SUM(total) total,COUNT(*) cnt
        FROM sales_data WHERE {' AND '.join(conds)}
        GROUP BY item_name ORDER BY total DESC""",params).fetchall()]
    conn.close()
    return jsonify(rows)

@app.route("/api/products/by-seller")
@login_required
def api_product_by_seller():
    group=request.args.get("group",""); item=request.args.get("item","")
    year=request.args.get("year",str(datetime.now().year)); month=request.args.get("month","")
    conn=get_db()
    date_cond=f"{year}-{month.zfill(2)}%" if month else f"{year}%"
    params=[date_cond]; conds=["sale_date LIKE ?","sale_date != ''","real_seller != ''"]
    if group: conds.append("item_group=?"); params.append(group)
    if item:  conds.append("item_name=?");  params.append(item)
    rows=[dict(r) for r in conn.execute(f"""SELECT real_seller seller_name,
        SUM(quantity) qty,SUM(total) total,COUNT(*) cnt
        FROM sales_data WHERE {' AND '.join(conds)}
        GROUP BY real_seller ORDER BY total DESC""",params).fetchall()]
    conn.close()
    return jsonify(rows)

@app.route("/api/export/xlsx/branches")
@login_required
def export_xlsx_branches():
    """거래처별 브랜드 입점 리스트 형식으로 판매처 내보내기"""
    year = request.args.get("year", str(datetime.now().year))
    conn = get_db()
    BRANDS_ORDER = ['줄즈','카오스','원더폴드','레카로','엔픽스','타프토이즈','ABC디자인']
    actual_groups = set(r[0] for r in conn.execute(
        "SELECT DISTINCT item_group FROM sales_data WHERE item_group!=''").fetchall())
    brands = [b for b in BRANDS_ORDER if b in actual_groups]
    for g in sorted(actual_groups):
        if g not in brands and g: brands.append(g)

    branches = [dict(r) for r in conn.execute("""
        SELECT id,name,ceo,ceo_phone,store_manager,store_manager_phone,
               manager,phone,address,email,status,note,region
        FROM branches ORDER BY note,name""").fetchall()]

    # 매장별 취급 브랜드
    brand_sold = {}
    for r in conn.execute(f"""SELECT real_seller,item_group FROM sales_data
        WHERE sale_date LIKE '{year}%' AND real_seller!='' AND item_group!=''
        GROUP BY real_seller,item_group""").fetchall():
        if r[0] not in brand_sold: brand_sold[r[0]] = set()
        brand_sold[r[0]].add(r[1])

    # 연간 실적
    year_sales_map = {r[0]:r[1] for r in conn.execute(f"""
        SELECT real_seller, SUM(total) FROM sales_data
        WHERE sale_date LIKE '{year}%' AND real_seller!=''
        GROUP BY real_seller""").fetchall()}
    conn.close()

    wb = openpyxl.Workbook()
    ws = wb.active; ws.title = "오프라인 거래처별 리스트"
    def mf(h): return PatternFill(start_color=h,end_color=h,fill_type="solid")
    def mft(h,b=True,s=10): return Font(color=h,bold=b,size=s)
    thin=Side(style='thin',color='D1D5DB')
    bdr=Border(left=thin,right=thin,top=thin,bottom=thin)
    ctr=Alignment(horizontal="center",vertical="center")
    rgt=Alignment(horizontal="right",vertical="center")

    total_cols=15+len(brands)
    ws.merge_cells(f"A1:{get_column_letter(total_cols)}1")
    c=ws.cell(row=1,column=1,value=f"오프라인 거래처별 브랜드 입점 리스트_{year}")
    c.fill=mf("1E3A5F"); c.font=mft("FFFFFF",True,12); c.alignment=ctr; ws.row_dimensions[1].height=28
    ws.row_dimensions[2].height=6

    hdrs=['업체구분','거래처명','실적용거래처명','전화번호','사장님','사장연락처',
          '점장','점장연락처','담당자','주소','Email','지역','상태','연간실적(원)',''] + brands
    for ci,h in enumerate(hdrs,1):
        c=ws.cell(row=3,column=ci,value=h)
        c.fill=mf("7C3AED") if ci>15 else mf("2563EB")
        c.font=mft("FFFFFF",True,10); c.alignment=ctr; c.border=bdr
    ws.row_dimensions[3].height=22

    cws=[12,24,26,14,12,14,12,14,12,45,26,8,8,14,4]+[7]*len(brands)
    for ci,w in enumerate(cws,1): ws.column_dimensions[get_column_letter(ci)].width=w

    prev_grp=None
    for ri,b in enumerate(branches,4):
        nm=b['name'] or ''; nml=nm.replace('_',' ').lower()
        if '베이비하우스' in nml: rf=mf("FFF7ED")
        elif '링크맘' in nml: rf=mf("F0FDF4")
        elif ri%2==0: rf=mf("F8FAFC")
        else: rf=mf("FFFFFF")
        grp=b.get('note','') or ''
        gv=grp if grp!=prev_grp else ''; prev_grp=grp
        sold=brand_sold.get(nm,set())
        yr_sales=year_sales_map.get(nm,0)
        row_vals=[gv,nm,nm,b.get('phone',''),b.get('ceo',''),b.get('ceo_phone',''),
                  b.get('store_manager',''),b.get('store_manager_phone',''),
                  b.get('manager',''),b.get('address',''),b.get('email',''),
                  b.get('region',''),b.get('status',''),yr_sales,''] + \
                 ['○' if br in sold else '' for br in brands]
        for ci,val in enumerate(row_vals,1):
            c=ws.cell(row=ri,column=ci,value=val); c.fill=rf; c.border=bdr; c.font=Font(size=10)
            if ci==14: c.number_format='#,##0'; c.alignment=rgt
            if ci>15: c.alignment=ctr
    ws.freeze_panes="A4"

    buf=io.BytesIO(); wb.save(buf); buf.seek(0)
    fname=f"오프라인_거래처별_브랜드_입점_리스트_{year}.xlsx"
    return send_file(buf,mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                     as_attachment=True,download_name=fname)

@app.route("/api/sales-data/summary")
@login_required
def sales_data_summary():
    year  = request.args.get("year", "")
    conn  = get_db()
    where = f"AND sale_date LIKE '{year}%'" if year else ""
    total = conn.execute(f"SELECT COUNT(*) c, SUM(total) t, SUM(quantity) q FROM sales_data WHERE 1=1 {where}").fetchone()
    by_seller = [dict(r) for r in conn.execute(f"""
        SELECT real_seller seller_name, COUNT(*) cnt, SUM(quantity) qty, SUM(total) total
        FROM sales_data WHERE real_seller != '' {where}
        GROUP BY real_seller ORDER BY total DESC""").fetchall()]
    by_group = [dict(r) for r in conn.execute(f"""
        SELECT item_group, COUNT(*) cnt, SUM(quantity) qty, SUM(total) total
        FROM sales_data WHERE item_group != '' {where} GROUP BY item_group ORDER BY total DESC""").fetchall()]
    by_date = [dict(r) for r in conn.execute(f"""
        SELECT sale_date, COUNT(*) cnt, SUM(total) total
        FROM sales_data WHERE sale_date != '' {where} GROUP BY sale_date ORDER BY sale_date""").fetchall()]
    by_item = [dict(r) for r in conn.execute(f"""
        SELECT item_name, SUM(quantity) qty, SUM(total) total
        FROM sales_data WHERE 1=1 {where} GROUP BY item_name ORDER BY total DESC LIMIT 20""").fetchall()]
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
        region = detect_region_from_name(name)
        existing = conn.execute("SELECT id FROM branches WHERE name=?", (name,)).fetchone()
        if not existing:
            conn.execute("""INSERT INTO branches(name,region,manager,phone,address,status,note)
                VALUES(?,?,?,?,?,?,?)""", (name, region,"","","","운영중",""))
            added += 1
        else:
            # 지역이 비어있으면 자동 채우기
            if region:
                conn.execute("UPDATE branches SET region=? WHERE id=? AND (region='' OR region IS NULL)",
                             (region, existing["id"]))
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
            real_seller = real_seller.replace('_', ' ')
            # 별칭 처리: 위드에이컴퍼니 → 베이비하우스 관악점
            real_seller = resolve_seller(real_seller)

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
    """sales_data의 real_seller 언더바→공백 정규화 + 지역 자동 배정"""
    conn = get_db()
    rows = conn.execute("SELECT DISTINCT real_seller FROM sales_data WHERE real_seller != ''").fetchall()
    updated = 0
    for r in rows:
        old = r[0]
        new = old.replace('_', ' ')
        if old != new:
            conn.execute("UPDATE sales_data SET real_seller=? WHERE real_seller=?", (new, old))
            updated += 1
    # branches 지역 자동 배정
    branches = conn.execute("SELECT id, name FROM branches WHERE region='' OR region IS NULL").fetchall()
    region_updated = 0
    for b in branches:
        region = detect_region_from_name(b["name"])
        if region:
            conn.execute("UPDATE branches SET region=? WHERE id=?", (region, b["id"]))
            region_updated += 1
    conn.commit(); conn.close()
    return jsonify({"ok": True, "normalized": updated, "region_updated": region_updated})

@app.route("/api/admin/merge-branches", methods=["POST"])
@login_required
def merge_branches():
    """띄어쓰기 차이로 중복된 판매처 통합 (실적 있는 쪽 기준)"""
    conn = get_db()
    branches = [dict(r) for r in conn.execute("SELECT id, name FROM branches ORDER BY name").fetchall()]

    def normalize(name):
        return name.replace('_', '').replace(' ', '').lower()

    # 정규화된 이름으로 그룹화
    groups = {}
    for b in branches:
        key = normalize(b['name'])
        if key not in groups:
            groups[key] = []
        groups[key].append(b)

    merged = 0
    for key, group in groups.items():
        if len(group) < 2:
            continue
        # 실적이 있는 쪽 선택 (year_actual 기준)
        y = datetime.now().year
        best = None
        best_sales = -1
        for b in group:
            sales = conn.execute("SELECT COALESCE(SUM(total),0) FROM sales_data WHERE real_seller=? AND sale_date LIKE ?",
                                 (b['name'], f"{y}%")).fetchone()[0]
            if sales > best_sales:
                best_sales = sales
                best = b
        # 나머지를 best로 리다이렉트
        for b in group:
            if b['id'] == best['id']:
                continue
            # sales_data의 real_seller 업데이트
            conn.execute("UPDATE sales_data SET real_seller=? WHERE real_seller=?",
                         (best['name'], b['name']))
            # branches 삭제
            conn.execute("DELETE FROM branches WHERE id=?", (b['id'],))
            merged += 1
    conn.commit(); conn.close()
    # 지역 자동 배정
    conn2 = get_db()
    branches_no_region = conn2.execute("SELECT id, name FROM branches WHERE region='' OR region IS NULL").fetchall()
    region_updated = 0
    for b in branches_no_region:
        region = detect_region_from_name(b["name"])
        if region:
            conn2.execute("UPDATE branches SET region=? WHERE id=?", (region, b["id"]))
            region_updated += 1
    conn2.commit(); conn2.close()
    return jsonify({"ok": True, "merged": merged, "region_updated": region_updated})

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
        WHERE {where}
        GROUP BY item_name ORDER BY total DESC""", params).fetchall()]

    summary = conn.execute(f"""
        SELECT COUNT(*) cnt, SUM(quantity) qty, SUM(total) total,
               MIN(sale_date) date_from, MAX(sale_date) date_to
        FROM sales_data WHERE {where}""", params).fetchone()

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
            MIN(sale_date) AS min_date
        FROM sales_data
        WHERE {where} AND sale_date != ''
        GROUP BY week_key
        ORDER BY week_key""", params).fetchall()]
    conn.close()

    # 주차별 일요일~토요일 범위 계산
    from datetime import datetime as dt, timedelta

    def get_week_range(date_str):
        d = dt.strptime(date_str, "%Y-%m-%d")
        wd = d.weekday()  # 0=월
        days_to_sun = (wd + 1) % 7
        sun = d - timedelta(days=days_to_sun)
        sat = sun + timedelta(days=6)
        return sun.strftime("%Y-%m-%d"), sat.strftime("%Y-%m-%d")

    for r in rows:
        try:
            r['week_start'], r['week_end'] = get_week_range(r['min_date'])
        except Exception:
            r['week_start'] = r.get('min_date', '')
            r['week_end']   = ''

    # 선택 월이 있으면 해당 월의 모든 주차를 채움 (데이터 없는 주도 표시)
    if month and rows:
        import calendar
        yr_int = int(year)
        mo_int = int(month)
        # 해당 월의 첫날~마지막날
        first_day = dt(yr_int, mo_int, 1)
        last_day  = dt(yr_int, mo_int, calendar.monthrange(yr_int, mo_int)[1])

        # 해당 월에 포함된 모든 주(일~토) 목록 생성
        all_weeks = {}
        cur = first_day
        while cur <= last_day:
            wk_start, wk_end = get_week_range(cur.strftime("%Y-%m-%d"))
            wk_key = cur.strftime("%Y-%W")
            if wk_key not in all_weeks:
                all_weeks[wk_key] = {'week_key': wk_key, 'week_start': wk_start, 'week_end': wk_end,
                                     'cnt': 0, 'qty': 0, 'total': 0, 'min_date': cur.strftime("%Y-%m-%d")}
            cur += timedelta(days=1)

        # 실제 데이터로 채우기
        data_map = {r['week_key']: r for r in rows}
        for wk_key in all_weeks:
            if wk_key in data_map:
                all_weeks[wk_key] = data_map[wk_key]

        rows = sorted(all_weeks.values(), key=lambda x: x['week_key'])

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
