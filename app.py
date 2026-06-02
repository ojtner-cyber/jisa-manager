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

    # SNS 정보 테이블 (블로그 중심)
    conn.execute("""CREATE TABLE IF NOT EXISTS sns_info (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        seller_name TEXT UNIQUE,
        blog_url TEXT DEFAULT '',
        blog_name TEXT DEFAULT '',
        blog_platform TEXT DEFAULT '',
        blog_total_posts INTEGER DEFAULT 0,
        blog_latest_date TEXT DEFAULT '',
        blog_recent_30d INTEGER DEFAULT 0,
        blog_recent_titles TEXT DEFAULT '',
        blog_keywords TEXT DEFAULT '',
        blog_has_product_post INTEGER DEFAULT 0,
        blog_score INTEGER DEFAULT 0,
        blog_grade TEXT DEFAULT '',
        last_searched TEXT DEFAULT '',
        memo TEXT DEFAULT '',
        updated_at TEXT DEFAULT ''
    )""")
    # 마이그레이션: 기존 테이블에 컬럼 추가
    try:
        sns_cols = [r[1] for r in conn.execute("PRAGMA table_info(sns_info)").fetchall()]
        new_cols = {
            'blog_platform': 'TEXT DEFAULT ""',
            'blog_total_posts': 'INTEGER DEFAULT 0',
            'blog_latest_date': 'TEXT DEFAULT ""',
            'blog_recent_30d': 'INTEGER DEFAULT 0',
            'blog_recent_titles': 'TEXT DEFAULT ""',
            'blog_keywords': 'TEXT DEFAULT ""',
            'blog_has_product_post': 'INTEGER DEFAULT 0',
            'blog_grade': 'TEXT DEFAULT ""',
            'last_searched': 'TEXT DEFAULT ""',
        }
        for col, typ in new_cols.items():
            if col not in sns_cols:
                conn.execute(f"ALTER TABLE sns_info ADD COLUMN {col} {typ}")
    except: pass

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

def remap_group(group, item_name=''):
    """품목그룹을 브랜드명으로 정규화 — 브랜드 태그 우선 적용"""
    import re
    g    = (group or '').strip()
    item = (item_name or '')

    # 제품명에서 [브랜드] 태그 추출
    brand_match = re.match(r'\[([^\]]+)\]', item)
    brand_tag   = brand_match.group(1) if brand_match else ''
    bt_lower    = brand_tag.lower()

    # ── 브랜드 태그 우선 판단 ──────────────────────
    if '줄즈' in brand_tag:                          return '줄즈'
    if '레카로' in brand_tag:                        return '레카로'
    if 'abc' in bt_lower:                            return 'ABC디자인'
    if '원더폴드' in brand_tag:                      return '원더폴드'
    if '카오스' in brand_tag:                        return '카오스'
    if '엔픽스' in brand_tag:                        return '엔픽스'
    if '타프토이즈' in brand_tag or 'taft' in bt_lower: return '타프토이즈'

    # ── 그룹명 기반 매핑 (태그 없는 경우) ──────────
    GROUP_MAP = {
        '유모차':          '줄즈',      # 태그 없으면 줄즈 기본
        '웨건':            '원더폴드',
        '컨버터블카시트':  '레카로',
        '주니어카시트':    '레카로',
        '토들러카시트':    '레카로',
        '카시트':          '레카로',
        '식탁의자':        '카오스',
        '하이체어':        '엔픽스',
        '보행기':          '엔픽스',
        '쏘서':            '엔픽스',
        '점퍼루':          '엔픽스',
        '휴대용부스터':    '엔픽스',
        'TAFTOYS':         '타프토이즈',
        '유아섬유류':      'ABC디자인',
    }
    return GROUP_MAP.get(g, g or '기타')

def normalize_item_name(name):
    """제품명에서 색상/옵션 완전 제거
    [줄즈]에어2_샌디타프 → [줄즈]에어2
    [줄즈]에어2_네이비블루(다크) → [줄즈]에어2
    [레카로]제논1_엘레강트베이지_캐노피 → [레카로]제논1
    """
    if not name: return name
    import re
    # 언더바 이후 모든 내용 제거 (색상, 옵션, 한정판 등)
    cleaned = re.sub(r'_.*$', '', name).strip()
    # 괄호 내용도 제거 (예: "다크", "한정판")
    cleaned = re.sub(r'\s*\([^)]*\)\s*$', '', cleaned).strip()
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

    # 색상 통합 — 같은 제품 합산
    norm_items = {}
    for r in sold_items:
        brand = remap_group(r['item_group'], r['item_name'])
        norm  = normalize_item_name(r['item_name'])
        key   = (brand, norm)
        if key not in norm_items:
            norm_items[key] = dict(r); norm_items[key]['item_name'] = norm; norm_items[key]['item_group'] = brand
        else:
            norm_items[key]['qty']   += r['qty']
            norm_items[key]['total'] += r['total']
            norm_items[key]['cnt']   += r['cnt']
    sold_items = sorted(norm_items.values(), key=lambda x: -x['total'])

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

@app.route("/api/script/generate", methods=["POST"])
@login_required
def api_script_generate():
    """데이터 기반 영업 스크립트 — 매장 패턴별 분기 + 매번 다른 각도"""
    import random, hashlib
    from datetime import datetime as dt2

    data        = request.json or {}
    seller      = data.get('seller', '')
    analysis    = data.get('analysis', {})
    gen_count   = data.get('gen_count', 0)

    year        = analysis.get('year', str(dt2.now().year))
    total       = analysis.get('total', 0)
    total_pct   = analysis.get('total_pct', 0.0)
    brands      = analysis.get('brand_summary', [])
    top5        = analysis.get('top5', [])
    sold_items  = analysis.get('sold_items', [])
    unsold_taft = analysis.get('unsold_taft', [])
    weekly      = analysis.get('weekly', [])

    # 시드: 매번 다른 결과
    seed_str = f"{seller}{gen_count}{dt2.now().strftime('%H%M%S')}"
    rng = random.Random(int(hashlib.md5(seed_str.encode()).hexdigest()[:8], 16))
    def w(n): return f"{n:,}"
    def pick(lst): return rng.choice(lst)

    # ── 데이터 분석 ────────────────────────────────
    top_brand   = brands[0]['brand'] if brands else ''
    top_pct     = brands[0]['pct']   if brands else 0
    top2_brand  = brands[1]['brand'] if len(brands) > 1 else ''
    top2_pct    = brands[1]['pct']   if len(brands) > 1 else 0
    top_item    = normalize_item_name(top5[0].get('item_name','')) if top5 else ''
    top_item_qty= top5[0].get('qty',0) if top5 else 0
    top_item_tot= top5[0].get('total',0) if top5 else 0

    all_brands  = set(b['brand'] for b in brands)
    missing_brs = [b for b in BRAND_ORDER if b not in all_brands and b != '타프토이즈']
    weak_brs    = [b for b in brands if b['pct'] < 5 and b['brand'] != '타프토이즈']

    taft_sold   = [r for r in sold_items if remap_group(r.get('item_group',''), r.get('item_name',''))=='타프토이즈']
    taft_total  = sum(r.get('total',0) for r in taft_sold)
    taft_cnt    = len(set(normalize_item_name(r.get('item_name','')) for r in taft_sold))
    taft_pct    = round(taft_total/total*100,1) if total else 0

    week_avg   = int(sum(wk.get('total',0) for wk in weekly)/len(weekly)) if weekly else 0
    week_trend = ''
    if len(weekly) >= 3:
        recent = [wk.get('total',0) for wk in weekly[-3:]]
        if recent[-1] > recent[0]*1.15:   week_trend = '강한상승'
        elif recent[-1] > recent[0]*1.05: week_trend = '상승'
        elif recent[-1] < recent[0]*0.85: week_trend = '하락'
        elif recent[-1] < recent[0]*0.95: week_trend = '약한하락'
        else: week_trend = '안정'

    CAT_PRI = {'아치/모빌':1,'액티비티짐':2,'트래블토이':3,'비지북':4,'큐브':5,'워터매트':6,'터미타임':7}
    rec_taft = sorted(unsold_taft, key=lambda x: CAT_PRI.get(x.get('category',''),9))

    month_now = dt2.now().month
    season_map = [(range_k, v) for range_k, v in [((3,4,5),'봄'),((6,7,8),'여름'),((9,10,11),'가을'),((12,1,2),'겨울')]]
    season = next((v for k,v in season_map if month_now in k), '봄')

    # 매장 등급 / 패턴
    if total_pct >= 10:    store_tier = 'VIP'
    elif total_pct >= 5:   store_tier = 'A'
    elif total_pct >= 2:   store_tier = 'B'
    else:                  store_tier = 'C'

    if top_pct >= 70:          store_pattern = 'mono'
    elif top_pct >= 45:        store_pattern = 'dominant'
    elif len(brands) >= 4 and top_pct < 40: store_pattern = 'balanced'
    elif len(brands) <= 2:     store_pattern = 'narrow'
    else:                      store_pattern = 'duo'

    taft_pattern = 'none' if taft_pct==0 else ('low' if taft_pct<5 else ('mid' if taft_pct<15 else 'high'))

    def rank_expr(tier, pct):
        if tier=='VIP': return f"상위 {pct}% 핵심 거래처"
        if tier=='A':   return f"상위권 거래처 (전체 대비 {pct}%)"
        if tier=='B':   return f"중요 거래처 (전체의 {pct}%)"
        return f"성장 가능성 높은 거래처 ({pct}%)"

    # ── 섹션 1: 오프닝 ────────────────────────────
    opening_pool = [
        f'''영업사원: "사장님, 안녕하세요! 오늘 오기 전에 {seller} 데이터 뽑아봤는데 숫자가 좋아서 오는 길에 기분이 좋았어요.\n잠깐 같이 보실 수 있으세요?"\n\n(태블릿/자료 꺼내며)\n\n영업사원: "저희 전체 거래처 중에서 {rank_expr(store_tier, total_pct)}이에요. {top_brand} 비중이 {top_pct}%로 탄탄하게 잡혀 있어요."''',
        f'''영업사원: "사장님, 들어오면서 {top_item} 진열이 눈에 잘 띄더라고요. 역시 매장 동선을 잘 잡고 계신 것 같아요."\n\n사장님: (반응)\n\n영업사원: "실제로 {top_item}이 {w(top_item_qty)}개 나갔거든요. 저희 거래처 중에서도 상위권이에요. {year}년 데이터 정리해서 가져왔는데, 같이 보실까요?"''',
        f'''영업사원: "사장님! 요즘 {top_brand} 어떠세요? 저희 다른 매장들이 {top_brand} 문의가 {pick(['많이 늘었다','꾸준하다','올해 특히 좋다'])}고 하더라고요."\n\n사장님: (반응)\n\n영업사원: "{seller}도 비슷한 흐름이에요. {year}년 데이터 분석해서 가져왔어요. 꼭 공유드리고 싶었습니다."''',
        f'''영업사원: "사장님, {{"봄":"신학기 시즌이라","여름":"여름이라","가을":"가을 나들이 시즌이라","겨울":"연말이라"}}[season] 매장 분위기 어떠세요?"\n\n사장님: (반응)\n\n영업사원: "맞아요. 저도 이 시즌에 딱 맞는 제안 드리려고 왔어요. {year}년 데이터 기반으로 준비했거든요."''',
    ]
    s1 = pick(opening_pool)

    # ── 섹션 2: 실적 공유 ────────────────────────
    brand_lines = '\n'.join(f"  · {b['brand']}: {w(b['total'])}원 ({b['pct']}%)" for b in brands[:5])
    trend_map = {'강한상승':'최근 3주 추이가 강하게 올라가고 있어요! 이 흐름 놓치면 안 됩니다.',
                 '상승':'최근 추이도 상승 중이라 지금이 발주 타이밍이에요.',
                 '안정':'판매가 꾸준히 안정적으로 유지되고 있어요. 탄탄한 베이스가 있는 거예요.',
                 '약한하락':'최근 3주가 살짝 내려갔는데, 진열 변화로 충분히 잡을 수 있어요.',
                 '하락':'최근 흐름이 빠졌는데, 오늘 원인 같이 찾아봐요. 해결책이 있어요.','':''}
    trend_comment = trend_map.get(week_trend,'')

    pattern_comments = {
        'mono': f"{top_brand} 하나에 {top_pct}% 집중하고 계신데, 이걸로 {w(total)}원을 만드신 게 대단해요. 근데 한 브랜드 의존도가 높으면 리스크가 있어요. 오늘 그 다음 전략 얘기해봐요.",
        'dominant': f"{top_brand}({top_pct}%)가 압도적이고, {top2_brand}({top2_pct}%)가 받쳐주는 구조예요. {top2_brand} 비중을 더 키우면 전체 매출이 쑥 올라가요.",
        'balanced': "브랜드 구성이 다양하게 잡혀 있어요. 각 브랜드가 역할 분담하는 구조인데, 조금만 최적화하면 같은 방문객 수로 매출을 더 올릴 수 있어요.",
        'narrow': f"지금 {len(brands)}개 브랜드 취급하고 계신데, 1-2개 추가하면 고객 이탈을 줄일 수 있어요.",
        'duo': f"{top_brand}와 {top2_brand}의 2강 구도예요. 이 구조 자체는 좋은데, 세 번째 기둥이 생기면 더 안정적이에요.",
    }
    real_talk = pattern_comments.get(store_pattern, '')

    motivation = pick([
        "분명히 고객들이 사장님 매장을 신뢰한다는 거예요. 추천을 잘 해주시니까요.",
        "제품을 그냥 파는 게 아니라 제대로 설명해서 파신다는 게 느껴져요.",
        "이 매출은 그냥 나오는 게 아니에요. 사장님이 만들어 낸 거예요.",
    ])
    pushback = pick([
        "다른 매장들이랑 비교하면 여기가 훨씬 잘하고 있어요. 체감이 안 될 뿐이에요.",
        "힘들다고 느끼실 때가 도약 직전인 경우가 많아요. 오늘 같이 방법 찾아봐요.",
        "그래도 이 숫자는 시장 평균보다 위에 있어요. 기반이 탄탄하다는 뜻이에요.",
    ])

    s2 = f'''영업사원: "사장님, {year}년 {seller} 전체 데이터예요.\n\n총 {w(total)}원 — {rank_expr(store_tier, total_pct)}입니다.\n{"주간 평균 " + w(week_avg) + "원으로 " if week_avg else ""}꾸준히 판매되고 있고요.\n\n브랜드별로 보면:\n{brand_lines}\n\n{real_talk}\n{"  ※ " + trend_comment if trend_comment else ""}\n\n{motivation}\n\n💡 힘들다 하시면:\n→ "{pushback}"'''

    # ── 섹션 3: 베스트 제품 ───────────────────────
    brand_insights = {
        '레카로': {'reason':'카시트는 안전 민감 제품 → 전문점 신뢰도 핵심. 설명 잘 해주시니까 팔림', 'upsell':'카시트 구매자에게 타프토이즈 카시트 장난감 추가 제안', 'risk':'재고 소진 시 이탈 위험 높음 — 안전재고 3개 권장'},
        '줄즈': {'reason':'SNS 바이럴 강함 + 색상 다양성 → 엄마 커뮤니티 추천 1위', 'upsell':'에어2 구매자에게 데이5 신색상 미리 예약 유도', 'risk':'시즌별 신색상 → 구색 부족 시 기회 손실'},
        '원더폴드': {'reason':'웨건 카테고리 독점적 포지션 → 비교 구매 없이 결정', 'upsell':'웨건 구매자에게 타프토이즈 트래블토이 번들 제안', 'risk':'전시 필수 — 실물 못 보면 구매 주저'},
        '엔픽스': {'reason':'국내 브랜드 신뢰 + 합리적 가격 → 재구매율 높음', 'upsell':'보행기 구매자에게 비바체(하이체어) 또는 쏘서 연계', 'risk':'시즌 수요 집중 — 봄여름 전 선발주 중요'},
        '카오스': {'reason':'하이체어 프리미엄 포지션 + 디자인 감성 → 인테리어 중시 부모층 강함', 'upsell':'하이체어 구매 후 이유식 용품 연계', 'risk':'높은 단가 → 충분한 설명과 체험 필수'},
        'ABC디자인': {'reason':'유럽 감성 디자인 → 20-30대 부모층 강함', 'upsell':'유모차+카시트 패밀리 세트 구성 제안', 'risk':'인지도 낮음 → 설명력이 판매 좌우'},
        '타프토이즈': {'reason':'완구 시장 최고 성장 브랜드 + 선물 수요 높음', 'upsell':'아치→모빌→비지북→큐브 시리즈 업셀', 'risk':'전시 위치가 판매 좌우 — 눈에 잘 띄는 곳 배치'},
    }

    best_blocks = []
    for i, r in enumerate(top5[:3]):
        nm    = normalize_item_name(r.get('item_name',''))
        qty   = r.get('qty', 0)
        tot   = r.get('total', 0)
        brand = remap_group(r.get('item_group',''), r.get('item_name',''))
        ins   = brand_insights.get(brand, {'reason':'검증된 베스트셀러','upsell':'','risk':''})
        if qty >= 20:   qty_comment = f"{w(qty)}개는 전국 상위권 판매량이에요."
        elif qty >= 10: qty_comment = f"{w(qty)}개, 이 브랜드 기준 잘 나가는 편이에요."
        else:           qty_comment = f"{w(qty)}개인데, 여기서 더 올릴 여지가 충분해요."
        block = f"  {i+1}위. {nm} — {w(qty)}개 / {w(tot)}원\n  {qty_comment}\n  잘 팔리는 이유: {ins['reason']}\n  연계 제안: {ins['upsell']}\n  주의: {ins['risk']}\n"
        best_blocks.append(block)

    stock_q = pick([
        f'"{top_item} 재고 지금 몇 개 남아계세요? 이 제품은 품절 나면 고객이 바로 온라인으로 가거든요."',
        f'"{top_item} 다음 발주 언제 생각하고 계세요? 제가 미리 물량 잡아드릴게요."',
        f'"{top_item} 재고 체크해보실 수 있어요? 이번 달 소진 속도 보고 발주량 같이 정해드릴게요."',
    ])
    pushback2 = pick([
        "혹시 최근에 진열 위치가 바뀌셨어요? 위치가 판매량에 정말 크게 영향 주거든요.",
        "고객 문의는 있는데 구매로 안 이어지나요? 어떤 제품과 비교하시는지 여쭤봐도 될까요?",
        "그럴 때일수록 재고 줄이고 다른 제품 비중 늘리는 게 맞을 수 있어요. 같이 봐요.",
    ])
    s3 = f'''영업사원: "이 매장 베스트 TOP3 분석해봤어요.\n\n{chr(10).join(best_blocks)}\n{stock_q}\n\n💡 \"요즘 그 제품 잘 안 나가요\" 하시면:\n→ "{pushback2}"'''

    # ── 섹션 4: 타프토이즈 ───────────────────────
    if taft_pattern == 'none':
        taft_approaches = [
            f'''영업사원: "사장님, 타프토이즈 아세요? 유아 완구 브랜드인데 베이비페어에서 카시트보다 줄 서는 브랜드가 됐어요.\n\n이 매장에 없는 이유가 있을 것 같아서요. 혹시 완구는 취급 안 하시는 정책인가요?\n\n사실 완구가 매장 객단가 올리는 데 효과적이에요. 카시트 하나 사러 온 고객이 {w(rec_taft[0].get("price",25000) if rec_taft else 25000)}원짜리 완구 하나 더 집어가거든요.\n\n💡 \"마진이 낮지 않나요?\" 하시면:\n→ \"오히려 반대예요. 카시트보다 완구 마진이 높아요. 재방문 효과도 있어요.\""''',
            f'''영업사원: "사장님, 솔직하게 여쭤볼게요. 지금 고객 한 분당 평균 구매 금액이 얼마인 것 같으세요?"\n\n사장님: (반응)\n\n영업사원: "저희 데이터로 {seller} 평균 객단가가 {w(int(total/len(sold_items)) if sold_items else 0)}원 정도예요. 근데 타프토이즈 취급 매장들은 평균 25,000원씩 더 나와요.\n카시트 사면서 완구 하나 더 집어가는 거거든요. 그 역할 할 제품이 지금 이 매장엔 없어요."''',
        ]
        s4 = pick(taft_approaches)
        if rec_taft:
            r0 = rec_taft[0]; nm0 = r0.get("name","").replace("[타프토이즈]","").strip()
            s4 += f'\n\n제가 이 매장에 맞는 제품 골라봤어요:\n  ◆ {nm0} ({r0.get("category","")}) — {w(r0.get("price",0))}원\n    "{r0.get("desc","")}"\n    처음엔 3종 소량으로 시작해보세요. 한 달 후에 반응 보고 확대해드릴게요.'

    elif taft_pattern == 'low':
        taft_names = [normalize_item_name(r.get('item_name','')) for r in taft_sold[:2]]
        untracked = [u for u in rec_taft if normalize_item_name(u.get('name','')) not in taft_names][:2]
        ut_txt = ""
        if untracked:
            u = untracked[0]
            ut_txt = f'\n  ◆ {u.get("name","").replace("[타프토이즈]","").strip()} ({u.get("category","")}) — {w(u.get("price",0))}원\n    "{u.get("desc","")}"'
        s4 = f'''영업사원: "타프토이즈 {taft_cnt}종에서 {w(taft_total)}원 나왔어요. 비중이 {taft_pct}%인데, 이걸 10%로만 올려도 전체 매출이 달라져요.\n\n지금 취급 중인 제품 고객들에게 시리즈 연결이 잘 안 되고 있을 가능성이 높아요.\n아치 산 고객한테 3주 후 \"아이가 자라면 이 제품이 딱이에요\" 연락하면 재방문이 돼요.\n\n이번에 추가 추천 제품:{ut_txt}\n\n💡 \"관리하기 어려워요\" 하시면:\n→ \"이 브랜드는 팔고 나면 고객이 알아서 찾아와요. 설명이 필요 없는 브랜드예요.\""''' 

    elif taft_pattern == 'mid':
        s4 = f'''영업사원: "타프토이즈가 {taft_pct}%까지 올라왔는데, 여기가 중간 고비예요. 이 브랜드가 20% 이상 되면 매장 이미지 자체가 바뀌거든요.\n\n{pick(["트래블토이는 카시트/유모차 옆에 두면 번들 구매가 자연스럽게 일어나요.","액티비티짐 하나만 놔도 인스타 감성 사진이 나와서 매장이 SNS에 올라가요.","비지북 시리즈는 선물용 수요가 강해서 돌잔치 코너 옆에 두면 효과적이에요."])}\n\n이번에 2종 추가해보시고, 한 달 후에 반응 체크해드릴게요."'''

    else:
        top_taft_item = normalize_item_name(taft_sold[0].get('item_name','')) if taft_sold else ''
        s4 = f'''영업사원: "타프토이즈 {taft_pct}%면 저희 거래처 중 최상위권이에요. {top_taft_item}을 중심으로 정말 잘 운영하고 계세요.\n\n이제 다음 레벨 얘기를 해도 될 것 같아요. 신상 독점 전시를 해보시는 건 어때요?\n전시한 매장들은 한 달 만에 타프 매출이 평균 {pick(["40%","35%","28%"])} 올랐어요."'''

    # ── 섹션 5: 구조 개선 ────────────────────────
    if missing_brs:
        miss1 = missing_brs[0]
        miss_advice = {'원더폴드':'웨건은 유모차와 겹치지 않아요. 오히려 유모차 사고 웨건도 사는 가정이 많아요.','ABC디자인':'ABC는 유럽 감성이라 줄즈와 다른 고객층이에요. 경쟁이 아니라 보완이에요.','카오스':'하이체어는 이유식 시작 6개월 필수품이에요. 카시트 구매 후 타이밍 맞게 제안하면 돼요.','엔픽스':'보행기/쏘서는 6-12개월 집중 수요예요. 카시트 구매 3-4개월 후 제안하면 재방문이 돼요.'}.get(miss1, f'{miss1}은 이 매장 고객층에 맞는 브랜드예요.')
        s5 = f'''영업사원: "솔직히 아쉬운 게 있어요. {", ".join(missing_brs[:2])} 쪽이 빠져있거든요.\n\n{miss_advice}\n\n지금 오는 고객들이 {miss1} 때문에 다른 매장 가는 경우가 있을 수 있어요.\n처음에 전시용 1개만 두고 반응 보세요.\n\n💡 \"그 브랜드 잘 모르는데요\" 하시면:\n→ \"제가 직접 설명 드리고, 첫 고객 상담도 같이 해드릴 수 있어요.\""''' 

    elif weak_brs:
        wb1 = weak_brs[0]; wb_name = wb1['brand']; wb_pct = wb1['pct']
        advice = pick([f'{wb_name}이 {wb_pct}%인데, 진열 위치만 바꿔도 달라져요.',f'{wb_name}은 {top_brand} 구매 고객에게 추가 제안하는 방식이 더 효과적이에요.',f'{wb_name} 단독보다 세트 구성으로 팔면 부담이 줄어요.'])
        s5 = f'''영업사원: "브랜드 구성은 좋은데, {wb_name} 비중이 {wb_pct}%로 낮아요.\n\n{advice}\n\n{pick(["제가 다음 방문 때 진열 레이아웃 같이 봐드릴게요.","이 브랜드 잘 파는 다른 매장 사례 공유해드릴게요.","한 달에 2배 올린 매장도 있어요. 비결 알려드릴게요."])}"''' 

    else:
        deep = pick([f'{top_brand}에서 {top_item} 잘 파시는데, 같은 브랜드 2-3종 더 깊이 파는 라인업 확장 전략이 있어요. 7종 이상 취급하면 전문 매장 이미지가 생겨요.','브랜드 구성은 완성 단계예요. 이제 각 브랜드에서 프리미엄 라인 하나씩 추가하면 객단가가 올라가요.','이 정도면 다음 스텝은 고객 관계 관리 시스템화예요. 매출이 한 단계 더 올라갈 수 있어요.'])
        s5 = f'''영업사원: "브랜드 구성은 정말 잘 잡혀 있어요. 진심으로 칭찬이에요.\n\n{deep}"''' 

    # ── 섹션 6: 시즌 전략 ────────────────────────
    season_data = {
        '봄':  {'items':['줄즈 에어2 봄 신색상 (3-4월 출시)','타프토이즈 어반가든 아치 (야외 테마)','엔픽스 보행기 (신학기 선물)'],'insight':'3-5월은 출생아 수 피크 + 어린이날 선물 수요 집중. 이 시기 발주가 1년 매출을 좌우해요.','action':'어린이날 전 선물 포장 세트 구성하면 객단가가 올라가요.'},
        '여름':{'items':['타프토이즈 워터매트 (6-8월 한정)','타프토이즈 팝앤플레이스테이션 (실내 놀이)','레카로 카시트 (여름 휴가 이동 수요)'],'insight':'워터매트는 7월까지가 발주 골든타임. 8월엔 재고 소진 빠르고 보충 어려워요.','action':'에어컨 켠 여름에 실내 액티비티짐이 의외로 잘 나가요.'},
        '가을':{'items':['레카로 카시트 (추석 선물 수요)','원더폴드 웨건 (가을 나들이)','타프토이즈 비지북 (독서의 계절)'],'insight':'추석 전후 2주가 선물 수요 피크. 재고 부족은 기회 손실이 커요.','action':'선물 포장 서비스 앞에 내세우면 입소문이 나요.'},
        '겨울':{'items':['타프토이즈 실내놀이 세트 (겨울 실내)','줄즈 크리스마스 에디션','엔픽스 점퍼루 (실내 활동)'],'insight':'12월 크리스마스 + 1월 설 선물로 더블 피크. 11월 말까지 발주 완료가 핵심이에요.','action':'크리스마스 패키지 구성이 있으면 인스타 바이럴이 잘 돼요.'},
    }
    sd = season_data.get(season, season_data['봄'])
    wait_pushback = pick(['사장님, 지켜보다가 타이밍 놓치면 다음 시즌까지 기다려야 해요.','주변 매장들이 지금 발주 넣고 있어요. 같이 움직이시는 게 유리해요.','소량으로라도 먼저 들여놓고 반응 보세요. 안 팔리면 제가 어떻게든 해결해드릴게요.'])
    s6 = f'''영업사원: "지금 {season}이잖아요. {sd["insight"]}\n\n이 시기 집중 제품:\n{chr(10).join(f"  · {p}" for p in sd["items"])}\n\n{sd["action"]}\n\n지금 발주 넣으시면 이번 주 안으로 납품 가능해요.\n시즌 물량은 한정이라 이번에 같이 넣어두시죠."\n\n💡 \"일단 지켜볼게요\" 하시면:\n→ "{wait_pushback}"'''

    # ── 섹션 7: 클로징 ───────────────────────────
    checklist = []
    if top5: checklist.append(f"{normalize_item_name(top5[0].get('item_name',''))} 재고 확보")
    if taft_pattern in ('none','low'): checklist.append("타프토이즈 3종 소량 시작")
    elif taft_pattern == 'mid': checklist.append("타프토이즈 2종 추가")
    if missing_brs: checklist.append(f"{missing_brs[0]} 전시용 1종 시작")
    checklist.append(f"{season} 시즌 집중 제품 발주")

    closing_variants = [
        f'''영업사원: "오늘 이야기 나눈 거 정리할게요:\n\n{chr(10).join(f"  ✓ {item}" for item in checklist)}\n\n다 한꺼번에 하시기 부담스러우시면, 오늘은 {checklist[0]}만 먼저 해도 돼요.\n어느 쪽부터 시작하실래요?"''',
        f'''영업사원: "오늘 제안드린 것들 다 하시면 {w(int(total*0.15))}~{w(int(total*0.25))}원 추가 매출이 가능해요.\n\n한 번에 다 하실 필요 없고요, 오늘 {checklist[0]}부터 시작해볼까요?\n발주서 바로 뽑아드릴게요."''',
        f'''영업사원: "사장님, 6개월 후 이 매장 그림을 그려봤어요. {top_brand}는 지금보다 {pick(["20%","15%","25%"])} 더 올리고, 타프토이즈가 10%를 차지하면 연간 {w(int(total*1.3))}원이 충분히 가능해요.\n\n그 첫 걸음을 오늘 같이 내딛어볼까요? 발주 구성 최적화해서 바로 올려드릴게요."''',
    ]

    next_visit = pick([
        f'3주 후에 오늘 발주한 제품들 반응 들으러 올게요. 사장님 목소리 기다려요.',
        f'다음 달 초에 다시 방문드릴게요. 그때 신규 제품 판매 현황 같이 봐요.',
        f'2주 후에 들를게요. 그때까지 신규 제품 첫 반응 꼭 알려주세요.',
    ])
    s7 = pick(closing_variants) + f'''\n\n────────────────────────\n다음 방문 약속:\n"{next_visit}"'''

    # ── 최종 조합 ─────────────────────────────────
    def section(title, content):
        return f"{'━'*52}\n【{title}】\n{'━'*52}\n{content}\n"

    now_str = dt2.now().strftime('%Y.%m.%d %H:%M')
    script = f"""{'='*57}
  매장 영업 방문 스크립트 — {seller}
  분석: {year}년 / 생성: {now_str} / 유형: [{store_tier}/{store_pattern}/타프{taft_pattern}]
{'='*57}

{section('1. 오프닝 — 첫 60초가 전체를 결정한다', s1)}
{section('2. 실적 공유 — 숫자로 신뢰를 만든다', s2)}
{section('3. 베스트 제품 심층 분석 — 왜 팔리는가', s3)}
{section('4. 타프토이즈 전략 — 성장 레버 잡기', s4)}
{section('5. 구조 개선 — 빈틈을 기회로', s5)}
{section('6. 시즌 전략 — 지금이 골든타임', s6)}
{section('7. 클로징 — 오늘 결정을 이끌어낸다', s7)}

{'─'*57}
  ※ {seller} / {year}년 실판매 데이터 기반 자동 생성
  ※ 재생성 시마다 다른 각도의 스크립트가 나옵니다
{'─'*57}"""

    return jsonify({'text': script, 'ok': True, 'seller': seller})

@app.route("/api/script/report", methods=["POST"])
@login_required
def api_script_report():
    """매장 분석 리포트 생성"""
    import random, hashlib
    from datetime import datetime as dt2

    data     = request.json or {}
    seller   = data.get('seller', '')
    analysis = data.get('analysis', {})

    year        = analysis.get('year', str(dt2.now().year))
    total       = analysis.get('total', 0)
    total_pct   = analysis.get('total_pct', 0.0)
    brands      = analysis.get('brand_summary', [])
    top5        = analysis.get('top5', [])
    sold_items  = analysis.get('sold_items', [])
    unsold_taft = analysis.get('unsold_taft', [])
    weekly      = analysis.get('weekly', [])

    now    = dt2.now()
    rng    = random.Random(int(hashlib.md5(f"{seller}{now.strftime('%H%M')}".encode()).hexdigest()[:8],16))
    def pick(lst): return rng.choice(lst)
    def w(n): return f"{int(n):,}" if n else '0'

    # ── 기본 지표 ─────────────────────────────────────
    total_qty   = sum(r.get('qty',0) for r in sold_items)
    total_cnt   = sum(r.get('cnt',0) for r in sold_items)
    item_cnt    = len(sold_items)
    brand_cnt   = len(brands)
    avg_per_tx  = int(total/total_cnt) if total_cnt else 0

    # 타프토이즈
    taft_items  = [r for r in sold_items if remap_group(r.get('item_group',''),r.get('item_name',''))=='타프토이즈']
    taft_total  = sum(r.get('total',0) for r in taft_items)
    taft_cnt_k  = len(set(normalize_item_name(r.get('item_name','')) for r in taft_items))
    taft_pct    = round(taft_total/total*100,1) if total else 0

    # 주별 분석
    week_avg    = int(sum(wk.get('total',0) for wk in weekly)/len(weekly)) if weekly else 0
    week_max    = max(weekly, key=lambda x:x.get('total',0)) if weekly else {}
    week_min    = min(weekly, key=lambda x:x.get('total',0)) if weekly else {}
    week_range  = week_max.get('total',0)-week_min.get('total',0)

    trend_label = ''; trend_detail = ''; growth_rate = 0
    if len(weekly) >= 3:
        recent = [wk.get('total',0) for wk in weekly[-4:]]
        growth_rate = (recent[-1]-recent[0])/recent[0]*100 if recent[0] else 0
        if   growth_rate > 20:  trend_label='강한 상승세'; trend_detail=f"최근 {len(recent)}주 {growth_rate:.1f}% 증가"
        elif growth_rate > 8:   trend_label='상승세';       trend_detail=f"최근 {len(recent)}주 {growth_rate:.1f}% 증가"
        elif growth_rate > 2:   trend_label='완만한 상승';  trend_detail=f"최근 {len(recent)}주 {growth_rate:.1f}% 증가"
        elif growth_rate < -20: trend_label='급격한 하락';  trend_detail=f"최근 {len(recent)}주 {abs(growth_rate):.1f}% 감소"
        elif growth_rate < -8:  trend_label='하락세';       trend_detail=f"최근 {len(recent)}주 {abs(growth_rate):.1f}% 감소"
        elif growth_rate < -2:  trend_label='완만한 하락';  trend_detail=f"최근 {len(recent)}주 {abs(growth_rate):.1f}% 감소"
        else:                   trend_label='안정';         trend_detail=f"최근 {len(recent)}주 ±2% 내외 유지"

    # ── 등급 산정 (매출 기준 — DB에서 전체 매장 분포 조회) ──
    try:
        conn_g = get_db()
        all_totals = sorted([r[0] for r in conn_g.execute(
            "SELECT SUM(total) FROM sales_data WHERE real_seller!='' GROUP BY real_seller").fetchall()
        ], reverse=True)
        conn_g.close()
        n = len(all_totals)
        thresh_a = all_totals[max(0,int(n*0.10)-1)] if n >= 10 else 0
        thresh_b = all_totals[max(0,int(n*0.30)-1)] if n >= 4  else 0
        thresh_c = all_totals[max(0,int(n*0.70)-1)] if n >= 2  else 0
        if total >= thresh_a:   grade='A'; grade_basis='전체 거래처 상위 10% 이내'
        elif total >= thresh_b: grade='B'; grade_basis='전체 거래처 상위 30% 이내'
        elif total >= thresh_c: grade='C'; grade_basis='전체 거래처 상위 70% 이내'
        else:                   grade='D'; grade_basis='전체 거래처 하위 30%'
    except:
        if   total_pct >= 10: grade='A'; grade_basis='전체 비중 10% 이상'
        elif total_pct >= 5:  grade='B'; grade_basis='전체 비중 5~10%'
        elif total_pct >= 2:  grade='C'; grade_basis='전체 비중 2~5%'
        else:                  grade='D'; grade_basis='전체 비중 2% 미만'

    # 브랜드 패턴
    top_brand = brands[0]['brand'] if brands else '-'
    top_pct_v = brands[0]['pct']   if brands else 0
    top2_brand= brands[1]['brand'] if len(brands)>1 else ''
    top2_pct_v= brands[1]['pct']   if len(brands)>1 else 0
    all_brand_set = set(b['brand'] for b in brands)
    missing_brs   = [b for b in BRAND_ORDER if b not in all_brand_set and b!='타프토이즈']
    weak_brs      = [b for b in brands if b['pct']<5 and b['brand']!='타프토이즈']

    if top_pct_v >= 70:          concentration='단일 브랜드 집중형'
    elif top_pct_v >= 45:        concentration='1강 중심형'
    elif brand_cnt >= 4 and top_pct_v < 40: concentration='다브랜드 균형형'
    elif brand_cnt <= 2:         concentration='소수 브랜드형'
    else:                         concentration='2강 구도형'

    # ── 총괄 현황 — 수백 가지 경우의 수 ──────────────
    # A등급별 코멘트 풀
    grade_comments = {
        'A': [
            f"저희 전체 거래처 {len([1])}개 중 상위 10%에 해당하는 핵심 거래처입니다. {top_brand}를 중심으로 안정적인 매출 구조를 갖추고 있으며, 지속적인 관리가 중요합니다.",
            f"매출 규모와 브랜드 구성 면에서 우수한 성과를 내고 있는 거래처입니다. {top_brand} 판매 역량이 특히 두드러지며, 추가 브랜드 확대 시 더 큰 성과가 기대됩니다.",
            f"연간 {w(total)}원의 매출을 기록한 상위권 거래처로, 사장님의 적극적인 영업 활동이 실적에 반영된 결과로 판단됩니다.",
        ],
        'B': [
            f"전체 거래처 중 상위 30% 이내에 위치한 성장형 거래처입니다. 현재의 판매 흐름이 지속된다면 A등급 진입도 충분히 가능합니다.",
            f"{w(total)}원 매출로 중상위권을 유지하고 있습니다. {top_brand} 비중({top_pct_v}%)이 높아 해당 브랜드 재고 관리가 실적에 직접 영향을 미칩니다.",
            f"안정적인 판매 기반을 갖추고 있으나, 취급 브랜드 다양화를 통해 매출 성장의 여지가 있습니다.",
        ],
        'C': [
            f"중간 수준의 매출을 유지하고 있으며, 집중 관리를 통한 성장이 기대되는 거래처입니다. 방문 빈도를 높이고 제품 구성 개선이 우선 과제입니다.",
            f"현재 매출 수준에서 브랜드 구성을 보완하고 핵심 제품의 재고 관리를 강화한다면 단기간 내 성과 개선이 가능합니다.",
            f"{top_brand} 중심의 판매 구조를 갖추고 있으나, 추가 브랜드 도입과 타프토이즈 확대를 통해 객단가 향상이 필요합니다.",
        ],
        'D': [
            f"매출 성장이 필요한 거래처로, 기본적인 제품 구성 점검과 함께 사장님과의 밀착 소통이 요구됩니다. 방문 주기를 단축하고 원인 분석이 선행되어야 합니다.",
            f"현재 매출이 목표 대비 낮은 수준입니다. 취급 제품 라인업 검토, 진열 방식 개선, 사장님 교육을 통한 판매 역량 강화가 필요합니다.",
            f"거래처 유지를 위한 집중 지원이 필요한 시점입니다. 방문 시 애로사항을 파악하고 단기 실행 가능한 개선 방안을 함께 마련해야 합니다.",
        ],
    }

    # 추이별 코멘트
    trend_comments = {
        '강한 상승세': [
            f"최근 판매 추이가 가파르게 상승하고 있어 매우 고무적입니다. 이 흐름을 유지하기 위한 재고 선제 확보가 중요합니다.",
            f"주간 매출이 {growth_rate:.0f}% 이상 증가하는 강한 성장세를 보이고 있습니다. 현 상황을 적극 활용해야 합니다.",
        ],
        '상승세': [
            f"꾸준한 상승 흐름이 확인됩니다. 성장 모멘텀을 유지하면서 취약 브랜드 보완을 병행하는 것이 효과적입니다.",
            f"안정적인 성장세를 보이고 있으며, 방문 시 성장 요인을 파악하여 타 거래처에도 적용할 수 있는 사례 발굴이 필요합니다.",
        ],
        '완만한 상승': [
            f"소폭의 성장이 지속되고 있습니다. 계절 수요와 신제품 도입을 통해 성장 속도를 높일 수 있을 것으로 판단됩니다.",
        ],
        '안정': [
            f"판매가 안정적으로 유지되고 있습니다. 현 수준의 유지와 함께 새로운 성장 동력 발굴이 필요합니다.",
            f"고른 판매 흐름이 지속되고 있으나, 안정세가 장기화되면 성장 둔화로 이어질 수 있어 신제품 도입을 검토할 시점입니다.",
        ],
        '완만한 하락': [
            f"소폭의 매출 감소가 감지됩니다. 진열 위치 점검, 재고 상황 확인, 경쟁 매장 동향 파악이 필요합니다.",
            f"완만한 하락 추세로, 현 시점에서 원인을 파악하고 조기 대응하는 것이 중요합니다.",
        ],
        '하락세': [
            f"판매 하락이 지속되고 있어 즉각적인 원인 분석과 대응이 필요합니다. 방문 빈도를 높이고 사장님과 심층 면담을 권고합니다.",
        ],
        '급격한 하락': [
            f"단기간 급격한 매출 감소가 확인되어 긴급 점검이 필요합니다. 경쟁사 진입, 매장 운영 변화, 재고 문제 등 원인을 즉시 파악해야 합니다.",
        ],
        '': ["판매 추이 분석을 위한 추가 데이터 축적이 필요합니다."],
    }

    # 브랜드 집중도 코멘트
    concentration_comments = {
        '단일 브랜드 집중형': [
            f"{top_brand} 의존도({top_pct_v}%)가 매우 높아 단일 브랜드 리스크가 존재합니다. 보완 브랜드 도입을 통한 포트폴리오 다변화가 권고됩니다.",
            f"{top_brand} 한 브랜드로 전체 매출의 {top_pct_v}%를 차지하고 있습니다. 해당 브랜드의 재고 관리가 전체 실적에 직결됩니다.",
        ],
        '1강 중심형': [
            f"{top_brand}({top_pct_v}%)가 주력이며 {top2_brand}({top2_pct_v}%)가 보조 역할을 하는 구조입니다. 2위 브랜드 성장이 전체 매출 확대의 핵심입니다.",
        ],
        '다브랜드 균형형': [
            f"다양한 브랜드를 고르게 취급하고 있어 안정적인 매출 구조를 갖추고 있습니다. 각 브랜드의 시너지 효과를 극대화하는 전략이 필요합니다.",
        ],
        '소수 브랜드형': [
            f"취급 브랜드가 {brand_cnt}개로 적어 고객 선택의 폭이 제한됩니다. 1-2개 브랜드 추가를 통한 구색 확장이 시급합니다.",
        ],
        '2강 구도형': [
            f"{top_brand}({top_pct_v}%)와 {top2_brand}({top2_pct_v}%)의 2강 구도가 형성되어 있습니다. 세 번째 핵심 브랜드를 육성하면 더 탄탄한 매출 기반을 만들 수 있습니다.",
        ],
    }

    # 한줄 평 — 매장 상황별 다양한 표현
    one_liners = {
        ('A','강한 상승세'): f"핵심 거래처로서 성장 가속도가 붙어 있는 이상적인 상태로, 적극적인 지원과 재고 확보로 상승세를 극대화해야 할 시점이다.",
        ('A','상승세'):      f"상위권 거래처로서 성장이 이어지고 있으며, 현재의 판매 방식과 구성을 유지하면서 추가 브랜드 도입을 검토할 적기다.",
        ('A','안정'):        f"핵심 거래처의 안정적인 매출을 유지하고 있으나, 성장 정체를 극복하기 위한 새로운 모멘텀 발굴이 필요하다.",
        ('A','하락세'):      f"핵심 거래처임에도 하락세가 감지되어 원인 파악과 즉각적인 대응이 요구되는 상황이다.",
        ('B','강한 상승세'): f"중상위권에서 강한 성장세를 보이고 있어 A등급 진입 가능성이 매우 높다. 지금이 집중 지원의 적기다.",
        ('B','상승세'):      f"안정적인 성장을 거듭하고 있는 거래처로, 지속적인 지원과 브랜드 다양화를 통해 상위 등급으로 도약이 기대된다.",
        ('B','안정'):        f"중상위권의 안정적인 매출을 유지하고 있으며, 집중 관리를 통해 A등급 진입을 목표로 해야 한다.",
        ('B','완만한 하락'): f"중상위권이지만 소폭 하락세가 감지되어 원인 파악과 함께 회복 전략이 필요하다.",
        ('B','하락세'):      f"잠재력 있는 거래처에서 하락이 이어지고 있어 즉각적인 현장 점검과 사장님 밀착 소통이 필요하다.",
        ('C','강한 상승세'): f"중간 수준이지만 강한 성장세가 확인되어 집중 지원 시 단기간 내 B등급 이상으로 성장이 가능하다.",
        ('C','안정'):        f"중간 수준의 매출을 유지하고 있으며, 브랜드 구성 개선과 방문 빈도 강화를 통한 성장 전략이 필요하다.",
        ('C','하락세'):      f"매출 수준과 하락세가 동시에 나타나 집중 관리가 시급한 거래처다.",
        ('D','강한 상승세'): f"저매출이지만 성장 신호가 감지되고 있어, 집중 지원을 통해 빠른 회복이 기대된다.",
        ('D','안정'):        f"매출 성장이 정체된 거래처로, 근본적인 판매 환경 개선과 사장님과의 긴밀한 협력이 필요하다.",
        ('D','하락세'):      f"즉각적인 원인 파악과 집중 지원이 필요한 거래처다. 이탈 방지를 위한 적극적인 관계 관리가 요구된다.",
    }
    one_liner_key = (grade, trend_label)
    one_liner = one_liners.get(one_liner_key,
        one_liners.get((grade,'안정'),
        f"{seller}은(는) {year}년 기준 [{grade}등급] 거래처로, {concentration}이며 매출 추이는 {trend_label if trend_label else '분석 중'}이다."))

    # 코멘트 선택
    grade_comment     = pick(grade_comments.get(grade, grade_comments['C']))
    trend_comment     = pick(trend_comments.get(trend_label, trend_comments['']))
    conc_comment      = pick(concentration_comments.get(concentration, ['포트폴리오 검토가 필요합니다.']))

    # ── 브랜드별 실적 표 (정렬 고정) ─────────────────
    # 숫자 오른쪽 정렬, 고정폭 폰트 기준
    brand_table = f"  {'브랜드':<10}  {'비율':>6}  {'판매금액':>15}  {'수량':>6}  {'평가'}\n"
    brand_table += f"  {'─'*10}  {'─'*6}  {'─'*15}  {'─'*6}  {'─'*8}\n"
    for b in brands:
        bar    = '■'*int(b['pct']/5) + '□'*(20-int(b['pct']/5))
        eval_k = '◎ 핵심' if b['pct']>=35 else ('○ 주력' if b['pct']>=15 else ('△ 보조' if b['pct']>=5 else '▽ 소량'))
        brand_table += f"  {b['brand']:<10}  {b['pct']:>5.1f}%  {w(b['total']):>15}원  {b['qty']:>5}개  {eval_k}\n"
        brand_table += f"  {'':10}  {bar}\n"
    brand_table += f"  {'─'*10}  {'─'*6}  {'─'*15}  {'─'*6}\n"
    brand_table += f"  {'합 계':<10}  {'100.0':>5}%  {w(total):>15}원  {w(total_qty):>5}개\n"

    # ── TOP5 제품 ──────────────────────────────────
    top_table = ''
    brand_insights_short = {
        '레카로':'안전 신뢰 → 전문점 강점', '줄즈':'SNS 바이럴 + 재구매', '원더폴드':'웨건 독점 포지션',
        '엔픽스':'국내 신뢰 + 가성비', '카오스':'하이체어 프리미엄', 'ABC디자인':'유럽 감성 + 패밀리',
        '타프토이즈':'완구 성장 + 선물 수요',
    }
    for i, r in enumerate(top5[:5], 1):
        nm    = normalize_item_name(r.get('item_name',''))
        br    = remap_group(r.get('item_group',''), r.get('item_name',''))
        qty   = r.get('qty',0)
        tot   = r.get('total',0)
        share = round(tot/total*100,1) if total else 0
        insight = brand_insights_short.get(br,'')
        top_table += f"  {i}위. [{br}] {nm}\n"
        top_table += f"       판매 {w(qty)}개 / {w(tot)}원 / 비중 {share}% / {insight}\n\n"

    # ── 주별 추이 + 품목 상세 ────────────────────────
    weekly_table = ''
    conn_w = get_db()
    for i, wk in enumerate(weekly, 1):
        wk_key = wk.get('week','')
        ws_    = wk.get('week_start','')[:10]
        we_    = wk.get('week_end','')[:10]
        tot_w  = wk.get('total',0)
        max_t  = week_max.get('total',1) or 1
        bar_len= int(tot_w/max_t*15)
        bar    = '▮'*bar_len + '▯'*(15-bar_len)
        weekly_table += f"\n  {i:2}주차 ({ws_}~{we_})\n"
        weekly_table += f"  매출: {bar}  {w(tot_w)}원  ({wk.get('qty',0)}개)\n"
        # 해당 주·해당 매장의 브랜드별 판매 상세
        if wk_key:
            try:
                # 해당 매장 + 해당 주 데이터만
                wk_items = conn_w.execute("""
                    SELECT item_group, item_name, SUM(quantity) qty, SUM(total) total
                    FROM sales_data
                    WHERE strftime('%Y-%W',sale_date)=? AND sale_date!=''
                      AND (real_seller=? OR real_seller=?)
                    GROUP BY item_name ORDER BY total DESC LIMIT 20""",
                    (wk_key, seller, seller_raw)).fetchall()

                # 브랜드별 집계
                brand_wi = {}
                for wi in wk_items:
                    br_n = remap_group(wi[0], wi[1])
                    nm_n = normalize_item_name(wi[1])
                    if br_n not in brand_wi:
                        brand_wi[br_n] = {'qty': 0, 'total': 0, 'items': []}
                    brand_wi[br_n]['qty']   += wi[2]
                    brand_wi[br_n]['total'] += wi[3]
                    brand_wi[br_n]['items'].append(nm_n)

                if brand_wi:
                    weekly_table += f"  ┌ 브랜드별 현황:\n"
                    for br_n, bv in sorted(brand_wi.items(), key=lambda x:-x[1]['total']):
                        pct_w = round(bv['total']/tot_w*100,1) if tot_w else 0
                        top_item = bv['items'][0].replace(f'[{br_n}]','').strip() if bv['items'] else ''
                        weekly_table += f"  │ {br_n:<10} {w(bv['qty']):>5}개  {w(bv['total']):>12}원  ({pct_w}%)  주력:{top_item}\n"
                    weekly_table += f"  └\n"
            except: pass
    conn_w.close()

    # ── 개선 포인트 ────────────────────────────────
    improvements = []
    imp_details  = []

    if missing_brs:
        brs_str = ', '.join(missing_brs[:3])
        improvements.append(f"미취급 브랜드 도입 검토 ({brs_str})")
        imp_details.append(pick([
            f"{missing_brs[0]}는 현재 매장 고객층과 부합하는 브랜드입니다. 전시용 1개부터 시작하여 반응을 확인하는 방식을 권고합니다.",
            f"{brs_str} 도입 시 현재 취급 브랜드와의 시너지 효과가 기대되며, 객단가 향상에도 기여할 수 있습니다.",
        ]))

    if weak_brs:
        wb_str = ', '.join(b['brand'] for b in weak_brs[:2])
        wb1_name = weak_brs[0]['brand']; wb1_pct = weak_brs[0]['pct']
        improvements.append(f"저비중 브랜드 활성화 ({wb_str})")
        imp_details.append(pick([
            f"{wb_str} 진열 위치를 주력 제품 옆으로 변경하고, 함께 구매 시 효과적인 조합을 사장님과 논의하는 것을 권고합니다.",
            f"{wb_str}은 단독 판매보다 기존 인기 제품과의 세트 구성으로 접근하면 판매 활성화에 도움이 됩니다.",
        ]))

    if taft_pct < 3:
        improvements.append(f"타프토이즈 신규 도입 필요 (현재 {taft_pct}%)")
        improvements.append(f"완구 카테고리 공백으로 인한 기회 손실 발생 가능")
        imp_details.append(pick([
            "카시트·유모차 구매 고객에게 타프토이즈를 번들 제안하면 추가 구매를 유도할 수 있습니다. 초기 3~5종으로 시작을 권고합니다.",
            "타프토이즈 미취급으로 인해 고객이 완구 구매 시 타 채널로 이탈하는 상황입니다. 진입 장벽이 낮은 20,000원대 제품부터 시작을 권고합니다.",
        ]))
    elif taft_pct < 8:
        improvements.append(f"타프토이즈 비중 확대 필요 (현재 {taft_pct}% → 목표 10%)")
        imp_details.append(pick([
            f"현재 {taft_cnt_k}종을 취급하고 있습니다. 미취급 카테고리(아치, 비지북, 큐브 등)를 보완하면 타프 매출이 2배 이상 성장 가능합니다.",
            "타프토이즈 판매 고객의 재방문율이 높습니다. 시리즈 구성으로 연속 구매를 유도하는 전략이 효과적입니다.",
        ]))

    if trend_label in ('하락세','급격한 하락'):
        improvements.append("판매 감소 원인 분석 및 즉각적 대응 필요")
        imp_details.append(pick([
            "방문 시 경쟁 매장 동향, 재고 상황, 진열 변화 여부를 파악하고 단기 실행 가능한 개선책을 제시해야 합니다.",
            "판매 하락의 주요 원인(재고 부족, 진열 문제, 경쟁사 진입 등)을 현장에서 직접 확인하고 즉시 대응이 필요합니다.",
        ]))

    if avg_per_tx < 100000:
        improvements.append(f"건당 매출 향상 필요 (현재 {w(avg_per_tx)}원)")
        imp_details.append("고가 제품(레카로 카시트, 원더폴드 웨건 등)과 타프토이즈 번들 구성을 통해 건당 구매액을 높이는 전략이 필요합니다.")

    if not improvements:
        improvements.append("현재 지표 전반 양호 — 현 수준 유지 및 점진적 확대 권고")
        imp_details.append("모든 핵심 지표가 양호한 수준입니다. 현재의 운영 방식을 유지하면서 신규 시즌 제품 선주문을 통한 기회 선점을 권고합니다.")

    imp_block = ''
    for i, (item, detail) in enumerate(zip(improvements, imp_details), 1):
        imp_block += f"\n  {i}. {item}\n     → {detail}\n"

    # ── 향후 관리 방향 — 등급·추이·거리·패턴별 ─────
    # 방문 주기 (100개 매장 관리 현실 반영)
    if grade == 'A' and trend_label in ('하락세','급격한 하락'):  visit_cycle = '1~2주'
    elif grade == 'A':                                              visit_cycle = '3~4주'
    elif grade == 'B' and trend_label in ('강한 상승세','상승세'): visit_cycle = '2~3주'
    elif grade == 'B':                                              visit_cycle = '3~4주'
    elif grade == 'C' and trend_label in ('하락세','급격한 하락'): visit_cycle = '1~2주'
    elif grade == 'C':                                              visit_cycle = '4~6주'
    elif grade == 'D' and trend_label in ('강한 상승세','상승세'): visit_cycle = '2~3주'
    elif grade == 'D':                                              visit_cycle = '2~4주'
    else:                                                           visit_cycle = '4주'

    # 목표 매출 (현실적 수치)
    if trend_label in ('강한 상승세','상승세'):  growth_target = 1.20
    elif trend_label in ('완만한 상승','안정'):   growth_target = 1.10
    elif trend_label in ('완만한 하락'):          growth_target = 1.05
    elif trend_label in ('하락세','급격한 하락'): growth_target = 1.00
    else:                                          growth_target = 1.10
    target_total = int(total * growth_target)

    # 핵심 액션 아이템 (상황별 자동 생성)
    action_items = []
    if grade == 'A':
        action_items.append(pick([
            f"{top_brand} 안전 재고 유지 (2~3주치 상시 확보) — 품절은 고객 이탈로 직결",
            f"시즌별 선주문 체계 구축 — {top_brand} 수요 예측 기반 선제적 재고 관리",
        ]))
        if taft_pct < 5:
            action_items.append("타프토이즈 카테고리 도입으로 객단가 추가 향상")
    elif grade == 'B':
        action_items.append(pick([
            f"{top_brand} 판매량 20% 확대 목표 수립 — 구체적 실행 방안 현장에서 협의",
            "A등급 진입을 위한 분기별 성과 점검 체계 수립",
        ]))
    elif grade == 'C':
        action_items.append(pick([
            "방문 빈도 강화 및 진열 개선 지원 — 현장 점검을 통한 즉각적 개선",
            "핵심 제품 집중 전략 수립 — 잘 팔리는 제품 라인 강화부터 시작",
        ]))
    else:
        action_items.append(pick([
            "긴급 현장 점검 및 사장님 면담 — 운영 현황과 애로사항 파악 선행",
            "단기 성과 목표 설정 — 1개월 내 가시적 개선 지표 공동 설정",
        ]))

    if missing_brs:
        action_items.append(f"{missing_brs[0]} 도입 제안 및 초기 교육 지원")
    if taft_pct > 0 and taft_pct < 10:
        action_items.append("타프토이즈 시리즈 판매 가이드 제공 및 전시 레이아웃 개선")
    if trend_label in ('강한 상승세','상승세'):
        action_items.append("성공 사례 문서화 — 타 거래처 적용 방안 검토")

    # 관리 방향 코멘트
    mgmt_comments = {
        ('A','강한 상승세'): "현재의 상승세를 최대한 활용할 수 있도록 재고 충분 공급과 신제품 우선 배정을 지원합니다.",
        ('A','안정'):        "핵심 거래처로서의 안정적 관계 유지를 최우선으로 하며, 정기 방문을 통한 신뢰 강화에 집중합니다.",
        ('A','하락세'):      "즉각적인 현장 방문을 통해 하락 원인을 파악하고, 핵심 거래처 이탈을 방지하기 위한 적극적 지원이 필요합니다.",
        ('B','상승세'):      "성장 모멘텀을 유지하면서 A등급 진입을 위한 중점 관리 대상으로 선정하여 집중 지원합니다.",
        ('B','안정'):        "안정적 성과를 인정하고 다음 단계 성장을 위한 구체적 방안을 함께 수립합니다.",
        ('C','강한 상승세'): "성장 신호를 놓치지 않도록 방문 빈도를 높이고, 성장 가속화를 위한 집중 지원을 시작합니다.",
        ('C','안정'):        "현상 유지에서 벗어나 성장 단계로 전환하기 위한 체계적인 지원 계획을 수립합니다.",
        ('D','상승세'):      "하위 등급에서의 성장 신호는 중요한 기회입니다. 즉각적인 지원으로 성장세를 가속화합니다.",
        ('D','하락세'):      "최하위 등급의 하락세는 거래처 이탈 위험을 의미합니다. 관계 유지를 최우선으로 집중 지원합니다.",
    }
    mgmt_key = (grade, trend_label)
    mgmt_comment = mgmt_comments.get(mgmt_key,
        mgmt_comments.get((grade,'안정'),
        f"[{grade}등급] 거래처로서 {visit_cycle} 주기의 정기 방문과 지속적인 관계 관리를 권고합니다."))

    action_block = '\n'.join(f"  {i+1}. {a}" for i,a in enumerate(action_items))

    # ── 최종 보고서 ────────────────────────────────
    sep1 = '─'*60; sep2 = '━'*60; sep3 = '·'*60

    report = f"""{sep2}
  매장 분석 리포트
{sep2}
  거래처명   : {seller}
  분석 기간  : {year}년 / 작성 일자 : {now.strftime('%Y년 %m월 %d일')}
  담당자     :                   제출처 :
{sep2}

{sep1}
  1. 총괄 현황
{sep1}

  [ 실적 요약 ]
  연간 매출      : {w(total)}원
  거래처 등급    : {grade}등급  ({grade_basis})
  매출 추이      : {trend_label if trend_label else '-'}  ({trend_detail if trend_detail else '-'})

  판매 건수      : {w(total_cnt)}건
  판매 수량      : {w(total_qty)}개
  건당 평균 매출 : {w(avg_per_tx)}원
  취급 브랜드 수 : {brand_cnt}개 / 취급 제품 종류 : {item_cnt}종
  타프토이즈     : {taft_pct}% ({taft_cnt_k}종 / {w(taft_total)}원)

{sep3}
  [ 종합 평가 ]
  {grade_comment}

  [ 매출 추이 분석 ]
  {trend_comment}

  [ 브랜드 구성 분석 ]
  {conc_comment}

  [ 한 줄 평 ]
  {one_liner}
{sep3}

{sep1}
  2. 브랜드별 판매 실적
{sep1}

{brand_table}
{sep1}
  3. 주요 판매 제품 (TOP 5)
{sep1}

{top_table}
{sep1}
  4. 주별 판매 추이
{sep1}

  주간 평균 : {w(week_avg)}원  /  최고 주 : {w(week_max.get('total',0))}원  /  최저 주 : {w(week_min.get('total',0))}원  /  편차 : {w(week_range)}원
{sep3}
{weekly_table if weekly_table else chr(10)+'  (데이터 없음)'+chr(10)}
{sep3}

{sep1}
  5. 개선 필요 사항
{sep1}
{imp_block}
{sep1}
  6. 향후 관리 방향
{sep1}

  {mgmt_comment}

  방문 권고 주기  : {visit_cycle}
  연간 목표 매출  : {w(target_total)}원  (현재 대비 +{int((growth_target-1)*100)}%)

  핵심 실행 항목:
{action_block}

{sep1}
  7. 담당자 기록란
{sep1}

  방문일 :                         방문 유형 : □ 정기  □ 긴급  □ 기타
  사장님 반응 :

  주요 논의 내용 :


  발주 내역 :


  특이사항 및 메모 :


  다음 방문 예정일 :               담당자 서명 :

{sep2}
  ※ 본 리포트는 {year}년 실판매 데이터를 기반으로 작성되었습니다.
  ※ 제출 전 담당자 확인 및 서명 필수.
{sep2}"""

    return jsonify({'report': report, 'ok': True})

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

    # ── 시트4: 월별 브랜드 요약 (세로형 — 출력 최적화) ──
    ws4 = wb.create_sheet("월별 브랜드 요약")
    ws4.merge_cells("A1:G1")
    c=ws4.cell(row=1,column=1,value=f"월별 브랜드 판매 요약_{year}")
    c.fill=mf(WHITE); c.font=mft(FONT_BLACK,True,12); c.alignment=center
    ws4.row_dimensions[1].height=26
    m_hdrs=['월','브랜드','판매금액(원)','판매수량','비율(%)','월합계(원)','누계(원)']
    for ci,h in enumerate(m_hdrs,1):
        c=ws4.cell(row=2,column=ci,value=h)
        c.fill=mf(GRAY_LIGHT); c.font=mft(FONT_GRAY,True,10); c.alignment=center; c.border=bdr_left
    ws4.row_dimensions[2].height=20
    ri4=3; cum_m=0
    for mo in months:
        mo_total = sum(idx.get((s,mo,b),{}).get('total',0) for s in sellers_list for b in brands)
        mo_qty   = sum(idx.get((s,mo,b),{}).get('qty',0)   for s in sellers_list for b in brands)
        cum_m += mo_total
        ws4.cell(row=ri4,column=1,value=f"{mo}월").fill=mf(GRAY_LIGHT)
        ws4.cell(row=ri4,column=2,value="전체 합계").fill=mf(GRAY_LIGHT)
        for ci,val in [(3,mo_total),(4,mo_qty),(5,100.0),(6,mo_total),(7,cum_m)]:
            c4=ws4.cell(row=ri4,column=ci,value=val)
            c4.fill=mf(GRAY_LIGHT); c4.font=mft(FONT_GRAY,True,10)
            c4.border=bdr_none; c4.alignment=right
            if ci in (3,6,7): c4.number_format=num_fmt
            if ci==5: c4.number_format='0.0'
        for ci in range(1,3):
            ws4.cell(row=ri4,column=ci).font=mft(FONT_GRAY,True,10)
            ws4.cell(row=ri4,column=ci).border=bdr_left
        ws4.row_dimensions[ri4].height=18; ri4+=1
        for b in brands:
            bv=sum(idx.get((s,mo,b),{}).get('total',0) for s in sellers_list)
            bq=sum(idx.get((s,mo,b),{}).get('qty',0)   for s in sellers_list)
            if bv==0: continue
            pct_b=round(bv/mo_total*100,1) if mo_total else 0
            ws4.cell(row=ri4,column=1,value=""); ws4.cell(row=ri4,column=2,value=f"  └ {b}")
            for ci,val in [(3,bv),(4,bq),(5,pct_b),(6,""),(7,"")]:
                c4=ws4.cell(row=ri4,column=ci,value=val); c4.border=bdr_none; c4.alignment=right
                if ci==3 and isinstance(val,int): c4.number_format=num_fmt
                if ci==5: c4.number_format='0.0'
            ws4.row_dimensions[ri4].height=16; ri4+=1
        ws4.row_dimensions[ri4].height=6; ri4+=1
    for ci,ww in enumerate([8,18,16,10,10,16,16],1):
        ws4.column_dimensions[get_column_letter(ci)].width=ww

    buf=io.BytesIO(); wb.save(buf); buf.seek(0)
    fname=f"오프라인_브랜드별정리_{year}{'_'+month+'월' if month else ''}.xlsx"
    return send_file(buf,mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                     as_attachment=True,download_name=fname)
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
    year   = request.args.get("year",   str(datetime.now().year))
    month  = request.args.get("month",  "")
    seller = request.args.get("seller", "").strip()
    conn   = get_db()

    qp = ["sale_date != ''"]
    pp = []
    if month: qp.append("sale_date LIKE ?"); pp.append(f"{year}-{month.zfill(2)}%")
    else:     qp.append("sale_date LIKE ?"); pp.append(f"{year}%")
    if seller: qp.append("real_seller = ?"); pp.append(seller)

    # 주차 목록
    week_rows = [dict(r) for r in conn.execute(f"""
        SELECT strftime('%Y-%W',sale_date) wk, MIN(sale_date) md,
               COUNT(*) cnt, SUM(quantity) qty, SUM(total) total
        FROM sales_data WHERE {' AND '.join(qp)} AND sale_date!=''
        GROUP BY wk ORDER BY wk""", pp).fetchall()]

    def wr(ds):
        d = dt2.strptime(ds, "%Y-%m-%d")
        sun = d - timedelta(days=(d.weekday()+1) % 7)
        return sun.strftime("%Y-%m-%d"), (sun+timedelta(days=6)).strftime("%Y-%m-%d")

    for r in week_rows:
        try: r['ws'], r['we'] = wr(r['md'])
        except: r['ws'] = r['we'] = ''

    weeks = week_rows
    brands = BRAND_ORDER

    # 주차 × 브랜드 × 매장 인덱스 조회
    raw = conn.execute(f"""
        SELECT strftime('%Y-%W',sale_date) wk, item_group, item_name,
               SUM(total) total, SUM(quantity) qty, real_seller
        FROM sales_data WHERE {' AND '.join(qp)} AND sale_date!=''
        GROUP BY wk, item_name, real_seller""", pp).fetchall()

    # 매장 목록 (업체구분 순)
    seller_cond = "AND real_seller=?" if seller else ""
    seller_params = [seller] if seller else []
    sellers_raw = conn.execute(
        f"SELECT DISTINCT real_seller FROM sales_data WHERE real_seller!='' AND sale_date LIKE ? {seller_cond}",
        [pp[0]] + seller_params).fetchall()
    sellers_list = [r[0] for r in sellers_raw]
    def bk(nm):
        nm=(nm or '').lower()
        if '베이비하우스' in nm: return (0,nm)
        if '링크맘' in nm: return (1,nm)
        if '베이비파크' in nm: return (2,nm)
        return (9,nm)
    sellers_list.sort(key=bk)

    # 업체구분 파악
    branch_group = {}
    for r in conn.execute("SELECT name,note FROM branches").fetchall():
        branch_group[r[0]] = r[1] or ''

    # ── idx: {(wk, brand, seller): {total, qty}} — 매장별 브랜드별 주차별 집계 ──
    idx_seller = {}  # (wk, brand, seller) → {total, qty}
    for r in raw:
        brand = remap_group(r[1], r[2])
        if not brand or brand == '기타': continue
        # r = (wk, item_group, item_name, total, qty, real_seller)
        s_nm = r[5] if len(r) > 5 else ''
        key3 = (r[0], brand, s_nm)
        if key3 not in idx_seller: idx_seller[key3] = {'total':0,'qty':0}
        idx_seller[key3]['total'] += r[3] or 0
        idx_seller[key3]['qty']   += r[4] or 0

    conn.close()  # 모든 쿼리 완료 후 닫기

    # {(wk, brand): {total, qty}}
    idx = {}
    for r in raw:
        brand = remap_group(r[1], r[2])
        if not brand or brand == '기타': continue
        key = (r[0], brand)
        if key not in idx: idx[key] = {'total':0,'qty':0}
        idx[key]['total'] += r[3] or 0
        idx[key]['qty']   += r[4] or 0

    # 주차별 제품 상세 (색상 통합)
    items_by_week = {}
    for r in raw:
        wk = r[0]
        brand = remap_group(r[1], r[2])
        norm  = normalize_item_name(r[2])
        k = (brand, norm)
        if wk not in items_by_week: items_by_week[wk] = {}
        if k not in items_by_week[wk]: items_by_week[wk][k] = {'item_group':brand,'item_name':norm,'qty':0,'total':0}
        items_by_week[wk][k]['qty']   += r[4] or 0
        items_by_week[wk][k]['total'] += r[3] or 0

    # ── 스타일 팔레트 (월별과 동일) ──
    WHITE      = "FFFFFF"
    GRAY_LIGHT = "F2F2F2"
    FONT_BLACK = "000000"
    FONT_GRAY  = "595959"
    thin_bdr   = Side(style='thin', color='BFBFBF')
    no_bdr     = Side(style=None)
    bdr_left   = Border(left=thin_bdr,right=thin_bdr,top=thin_bdr,bottom=thin_bdr)
    bdr_none   = Border(left=no_bdr,right=no_bdr,top=no_bdr,bottom=no_bdr)
    center     = Alignment(horizontal="center",vertical="center")
    right      = Alignment(horizontal="right",vertical="center")
    num_fmt    = '#,##0'
    mf  = lambda h: PatternFill(start_color=h,end_color=h,fill_type="solid")
    mft = lambda h,b=False,s=10: Font(color=h,bold=b,size=s)

    wb  = openpyxl.Workbook()
    col_start = 4  # A=업체구분, B=거래처명, C=실적용, D부터 데이터

    def build_brand_sheet(wb_ref, title, field, is_first=False):
        ws = wb_ref.active if is_first else wb_ref.create_sheet(title)
        if is_first: ws.title = title
        total_cols = 3 + len(weeks) * (len(brands)+1)

        # 행1: 타이틀 — 흰색
        ws.merge_cells(f"A1:{get_column_letter(total_cols)}1")
        c = ws.cell(row=1,column=1,value=f"오프라인 주별 {'판매금액' if field=='total' else '판매수량'} 브랜드별 정리_{year}")
        c.fill=mf(WHITE); c.font=mft(FONT_BLACK,True,12); c.alignment=center
        ws.row_dimensions[1].height=26

        # 행2: 고정 헤더 — 연한 회색
        for ci,h in enumerate(["업체구분","거래처명","실적용거래처명"],1):
            c = ws.cell(row=2,column=ci,value=h)
            c.fill=mf(GRAY_LIGHT); c.font=mft(FONT_GRAY,True,10); c.alignment=center; c.border=bdr_left
        ws.merge_cells("A2:A3"); ws.merge_cells("B2:B3"); ws.merge_cells("C2:C3")

        # 주차 헤더
        col = col_start
        for i,r in enumerate(weeks):
            span = len(brands)+1
            end_col = col+span-1
            ws.merge_cells(f"{get_column_letter(col)}2:{get_column_letter(end_col)}2")
            label = f"{i+1}주차 ({r['ws']}~{r['we']})"
            c = ws.cell(row=2,column=col,value=label)
            c.fill=mf(GRAY_LIGHT); c.font=mft(FONT_GRAY,True,10); c.alignment=center; c.border=bdr_left
            for b in brands:
                c2 = ws.cell(row=3,column=col,value=f"{b}{'금액' if field=='total' else '수량'}")
                c2.fill=mf(GRAY_LIGHT); c2.font=mft(FONT_GRAY,False,9); c2.alignment=center; c2.border=bdr_none
                col+=1
            c2 = ws.cell(row=3,column=col,value="합계")
            c2.fill=mf(GRAY_LIGHT); c2.font=mft(FONT_GRAY,True,9); c2.alignment=center; c2.border=bdr_left
            col+=1
        ws.row_dimensions[2].height=18; ws.row_dimensions[3].height=16

        # 행4: 빈 구분행
        for ci in range(1,total_cols+1):
            ws.cell(row=4,column=ci,value="").fill=mf(WHITE)
            ws.cell(row=4,column=ci).border=bdr_none
        ws.row_dimensions[4].height=4

        # 컬럼 너비
        ws.column_dimensions['A'].width=12; ws.column_dimensions['B'].width=22; ws.column_dimensions['C'].width=24
        for mo_i in range(len(weeks)):
            for b_i in range(len(brands)+1):
                ws.column_dimensions[get_column_letter(col_start+mo_i*(len(brands)+1)+b_i)].width=10

        # 데이터 행 (5행~)
        prev_grp = None
        for ri,s in enumerate(sellers_list, 5):
            grp = branch_group.get(s,''); gv = grp if grp!=prev_grp else ''; prev_grp=grp
            for ci,val in enumerate([gv,s,s],1):
                c=ws.cell(row=ri,column=ci,value=val)
                c.fill=mf(WHITE); c.border=bdr_left; c.font=mft(FONT_BLACK if ci>1 else FONT_GRAY,False,10)
            col=col_start
            for wk_r in weeks:
                wk = wk_r['wk']
                mt = 0
                for b in brands:
                    # 매장별+브랜드별 값: idx_seller 사용
                    val = idx_seller.get((wk, b, s), {}).get(field, 0)
                    mt += val
                    c=ws.cell(row=ri,column=col,value=val if val else 0)
                    c.fill=mf(WHITE); c.border=bdr_none; c.alignment=right
                    c.number_format=num_fmt; c.font=mft(FONT_BLACK,False,10); col+=1
                c=ws.cell(row=ri,column=col,value=mt)
                c.fill=mf(WHITE); c.font=mft(FONT_BLACK,True,10)
                c.border=Border(left=thin_bdr,right=no_bdr,top=no_bdr,bottom=no_bdr)
                c.alignment=right; c.number_format=num_fmt; col+=1

        # 합계 행
        tot_row = len(sellers_list)+5
        for ci,val in enumerate(["합계","",""],1):
            c=ws.cell(row=tot_row,column=ci,value=val)
            c.fill=mf(WHITE); c.border=bdr_left; c.font=mft(FONT_BLACK,True,10)
        col=col_start
        for wk_r in weeks:
            wk=wk_r['wk']
            for b in brands:
                tv = sum(idx_seller.get((wk,b,s),{}).get(field,0) for s in sellers_list)
                c=ws.cell(row=tot_row,column=col,value=tv)
                c.fill=mf(WHITE); c.border=bdr_none; c.alignment=right
                c.number_format=num_fmt; c.font=mft(FONT_BLACK,True,10); col+=1
            grand = sum(idx_seller.get((wk,b,s),{}).get(field,0) for s in sellers_list for b in brands)
            c=ws.cell(row=tot_row,column=col,value=grand)
            c.fill=mf(WHITE); c.font=mft(FONT_BLACK,True,10)
            c.border=Border(left=thin_bdr,right=no_bdr,top=no_bdr,bottom=no_bdr)
            c.alignment=right; c.number_format=num_fmt; col+=1
        ws.freeze_panes="D5"

    # ── 시트1: 주별 요약 (세로형 — 출력 최적화) ──
    ws_sum = wb.active; ws_sum.title="주별 요약"

    # 타이틀
    ws_sum.merge_cells("A1:G1")
    c=ws_sum.cell(row=1,column=1,value=f"주별 판매 실적 요약_{year}")
    c.fill=mf(WHITE); c.font=mft(FONT_BLACK,True,12); c.alignment=center
    ws_sum.row_dimensions[1].height=26

    # 헤더행
    sum_hdrs=['주차','기간','브랜드','판매금액(원)','판매수량','비율(%)','누계금액(원)']
    for ci,h in enumerate(sum_hdrs,1):
        c=ws_sum.cell(row=2,column=ci,value=h)
        c.fill=mf(GRAY_LIGHT); c.font=mft(FONT_GRAY,True,10); c.alignment=center; c.border=bdr_left
    ws_sum.row_dimensions[2].height=20

    ri=3
    cumulative=0
    for i,r in enumerate(weeks):
        wk=r['wk']
        wk_total=r.get('total',0); wk_qty=r.get('qty',0); wk_cnt=r.get('cnt',0)
        cumulative+=wk_total

        # 주차 소계 행
        ws_sum.cell(row=ri,column=1,value=f"{i+1}주차").fill=mf(GRAY_LIGHT)
        ws_sum.cell(row=ri,column=2,value=f"{r['ws']}~{r['we']}").fill=mf(GRAY_LIGHT)
        ws_sum.cell(row=ri,column=3,value="전체 합계").fill=mf(GRAY_LIGHT)
        for ci,val in [(4,wk_total),(5,wk_qty),(6,100.0),(7,cumulative)]:
            c=ws_sum.cell(row=ri,column=ci,value=val)
            c.fill=mf(GRAY_LIGHT); c.font=mft(FONT_GRAY,True,10)
            c.border=bdr_none; c.alignment=right
            if ci in (4,7): c.number_format=num_fmt
            if ci==6: c.number_format='0.0'
        for ci in range(1,3):
            ws_sum.cell(row=ri,column=ci).font=mft(FONT_GRAY,True,10)
            ws_sum.cell(row=ri,column=ci).border=bdr_left
        ws_sum.row_dimensions[ri].height=18
        ri+=1

        # 브랜드별 세부 행 (해당 주차)
        for b in brands:
            bv = sum(idx_seller.get((wk,b,s),{}).get('total',0) for s in sellers_list)
            bq = sum(idx_seller.get((wk,b,s),{}).get('qty',0) for s in sellers_list)
            if bv == 0: continue
            pct = round(bv/wk_total*100,1) if wk_total else 0
            ws_sum.cell(row=ri,column=1,value="")
            ws_sum.cell(row=ri,column=2,value="")
            ws_sum.cell(row=ri,column=3,value=f"  └ {b}")
            for ci,val in [(4,bv),(5,bq),(6,pct),(7,"")]:
                c=ws_sum.cell(row=ri,column=ci,value=val)
                c.border=bdr_none; c.alignment=right
                if ci==4 and isinstance(val,int): c.number_format=num_fmt
                if ci==6: c.number_format='0.0'
            ws_sum.row_dimensions[ri].height=16
            ri+=1

        # 구분 공백행
        ws_sum.row_dimensions[ri].height=6; ri+=1

    # 열 너비
    for ci,w in enumerate([10,28,18,16,10,10,16],1):
        ws_sum.column_dimensions[get_column_letter(ci)].width=w

    # ── 시트2: 브랜드별 금액 ──
    build_brand_sheet(wb, "브랜드별 금액", "total", False)
    # ── 시트3: 브랜드별 수량 ──
    build_brand_sheet(wb, "브랜드별 수량", "qty", False)

    # ── 시트4: 제품별 상세 ──
    ws4=wb.create_sheet("제품별 상세")
    item_hdrs=['주차','기간','브랜드','제품명','판매수량','판매금액(원)']
    for ci,h in enumerate(item_hdrs,1):
        c=ws4.cell(row=1,column=ci,value=h)
        c.fill=mf(GRAY_LIGHT); c.font=mft(FONT_GRAY,True,10); c.alignment=center; c.border=bdr_left if ci<=2 else bdr_none
    ws4.row_dimensions[1].height=20
    ri=2
    for i,r in enumerate(weeks):
        for k,item in sorted(items_by_week.get(r['wk'],{}).items(), key=lambda x:-x[1]['total']):
            for ci,v in enumerate([f"{i+1}주차",f"{r['ws']}~{r['we']}",item['item_group'],item['item_name'],item['qty'],item['total']],1):
                c=ws4.cell(row=ri,column=ci,value=v); c.border=bdr_left if ci<=2 else bdr_none
                if ci>=5: c.alignment=right
                if ci==6 and isinstance(v,int): c.number_format=num_fmt
            ri+=1
    for ci,w in enumerate([10,26,14,36,12,16],1): ws4.column_dimensions[get_column_letter(ci)].width=w

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
    raw = conn.execute("""
        SELECT item_group, item_name,
               SUM(quantity) qty, SUM(total) total
        FROM sales_data WHERE item_group != '' AND item_group IS NOT NULL
        GROUP BY item_group, item_name""").fetchall()
    conn.close()

    # remap 후 재집계 (색상 통합하여 실제 종류 수 계산)
    brand_items = {}  # brand → set of normalized names
    brand_totals = {}
    for r in raw:
        brand = remap_group(r[0], r[1])
        if not brand or brand == '기타': continue
        norm  = normalize_item_name(r[1])
        if brand not in brand_items:
            brand_items[brand]  = set()
            brand_totals[brand] = {'qty': 0, 'total': 0}
        brand_items[brand].add(norm)
        brand_totals[brand]['qty']   += r[2] or 0
        brand_totals[brand]['total'] += r[3] or 0

    result = [
        {'item_group': b, 'item_cnt': len(brand_items[b]),
         'qty': brand_totals[b]['qty'], 'total': brand_totals[b]['total']}
        for b in brand_items
    ]
    result.sort(key=lambda x: get_group_sort_key(x['item_group']))
    return jsonify(result)

@app.route("/api/products/items")
@login_required
def api_product_items():
    group  = request.args.get("group",  "")
    seller = request.args.get("seller", "").strip()
    year   = request.args.get("year",   str(datetime.now().year))
    month  = request.args.get("month",  "")
    conn   = get_db()
    date_cond = f"{year}-{month.zfill(2)}%" if month else f"{year}%"
    params = [date_cond]; conds = ["sale_date LIKE ?", "sale_date != ''"]
    if seller: conds.append("real_seller=?"); params.append(seller)
    raw = [dict(r) for r in conn.execute(f"""
        SELECT item_name, item_group, SUM(quantity) qty,
               AVG(unit_price) avg_price, SUM(total) total, COUNT(*) cnt
        FROM sales_data WHERE {' AND '.join(conds)}
        GROUP BY item_name ORDER BY total DESC""", params).fetchall()]
    conn.close()
    # 브랜드 필터 + 정규화 + 재집계 (색상 통합)
    merged = {}
    for r in raw:
        brand = remap_group(r['item_group'], r['item_name'])
        if not brand or brand == '기타': continue
        if group and brand != group: continue
        norm = normalize_item_name(r['item_name'])
        key  = (brand, norm)
        if key not in merged:
            merged[key] = {'item_name': norm, 'item_group': brand,
                           'qty': 0, 'avg_price': r['avg_price'], 'total': 0, 'cnt': 0}
        merged[key]['qty']   += r['qty']
        merged[key]['total'] += r['total']
        merged[key]['cnt']   += r['cnt']
    return jsonify(sorted(merged.values(), key=lambda x: -x['total']))

@app.route("/api/products/by-seller")
@login_required
def api_product_by_seller():
    """특정 브랜드/품목의 매장별 판매 현황"""
    group  = request.args.get("group",  "")
    item   = request.args.get("item",   "")
    year   = request.args.get("year",   str(datetime.now().year))
    month  = request.args.get("month",  "")
    conn   = get_db()
    date_cond = f"{year}-{month.zfill(2)}%" if month else f"{year}%"
    params = [date_cond]; conds = ["sale_date LIKE ?", "sale_date != ''", "real_seller != ''"]
    raw = [dict(r) for r in conn.execute(f"""
        SELECT real_seller seller_name, item_group, item_name,
               SUM(quantity) qty, SUM(total) total, COUNT(*) cnt
        FROM sales_data WHERE {' AND '.join(conds)}
        GROUP BY real_seller, item_name ORDER BY total DESC""", params).fetchall()]
    conn.close()
    # 브랜드/아이템 필터 + 매장별 재집계
    merged = {}
    for r in raw:
        brand = remap_group(r['item_group'], r['item_name'])
        if group and brand != group: continue
        if item and normalize_item_name(r['item_name']) != item: continue
        nm = r['seller_name']
        if nm not in merged:
            merged[nm] = {'seller_name': nm, 'qty': 0, 'total': 0, 'cnt': 0}
        merged[nm]['qty']   += r['qty']
        merged[nm]['total'] += r['total']
        merged[nm]['cnt']   += r['cnt']
    return jsonify(sorted(merged.values(), key=lambda x: -x['total']))

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

@app.route("/api/products/trend")
@login_required
def api_product_trend():
    """제품별 일별·주별 판매 추이"""
    item   = request.args.get("item",   "")
    group  = request.args.get("group",  "")
    year   = request.args.get("year",   str(datetime.now().year))
    month  = request.args.get("month",  "")
    conn   = get_db()
    date_cond = f"{year}-{month.zfill(2)}%" if month else f"{year}%"

    # item_name 정규화 역매핑 — normalize된 이름으로 like 쿼리
    conds  = ["sale_date LIKE ?", "sale_date != ''"]
    params = [date_cond]
    if group: conds.append("item_group=?"); params.append(group)
    # item은 정규화명이므로 LIKE로 포함 검색
    if item:
        # [브랜드] 부분 제거하고 모델명만 검색
        clean = item.replace('[','').replace(']','')
        import re
        brand_m = re.match(r'^([^\]]+)\]', item[1:]) if item.startswith('[') else None
        model   = re.sub(r'^\[[^\]]+\]', '', item).strip()
        if model:
            conds.append("item_name LIKE ?"); params.append(f"%{model}%")

    # 일별 추이
    daily = [dict(r) for r in conn.execute(f"""
        SELECT sale_date, SUM(quantity) qty, SUM(total) total, COUNT(*) cnt
        FROM sales_data WHERE {' AND '.join(conds)}
        GROUP BY sale_date ORDER BY sale_date""", params).fetchall()]

    # 주별 추이
    weekly_raw = conn.execute(f"""
        SELECT strftime('%Y-%W', sale_date) wk, MIN(sale_date) md,
               SUM(quantity) qty, SUM(total) total
        FROM sales_data WHERE {' AND '.join(conds)} AND sale_date!=''
        GROUP BY wk ORDER BY wk""", params).fetchall()

    from datetime import datetime as dt2, timedelta
    weekly = []
    for r in weekly_raw:
        try:
            d = dt2.strptime(r[1], "%Y-%m-%d")
            sun = d - timedelta(days=(d.weekday()+1)%7)
            label = sun.strftime("%m/%d")
        except: label = r[0]
        weekly.append({'wk': r[0], 'label': label, 'qty': r[2], 'total': r[3]})

    # 매장별 판매 현황 (기존 by-seller와 동일)
    by_seller = [dict(r) for r in conn.execute(f"""
        SELECT real_seller seller_name, SUM(quantity) qty, SUM(total) total, COUNT(*) cnt
        FROM sales_data WHERE {' AND '.join(conds)} AND real_seller!=''
        GROUP BY real_seller ORDER BY total DESC LIMIT 20""", params).fetchall()]

    conn.close()
    return jsonify({'daily': daily, 'weekly': weekly, 'by_seller': by_seller, 'item': item})

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
    """띄어쓰기/언더바 차이로 중복된 판매처 통합 — 연간 실적 기준, 모든 연락처 정보 병합"""
    conn = get_db()
    year = str(datetime.now().year)
    branches = [dict(r) for r in conn.execute(
        "SELECT id,name,ceo,ceo_phone,store_manager,store_manager_phone,manager,phone,address,email,region,note,status FROM branches ORDER BY name").fetchall()]

    def normalize_nm(name):
        return name.replace('_','').replace(' ','').replace('(','').replace(')','').lower()

    # 정규화된 이름으로 그룹화
    groups = {}
    for b in branches:
        key = normalize_nm(b['name'])
        groups.setdefault(key, []).append(b)

    merged = 0
    for key, group in groups.items():
        if len(group) < 2: continue
        # 연간 실적 기준으로 대표 선정
        best = None; best_sales = -1
        for b in group:
            sales = conn.execute(
                "SELECT COALESCE(SUM(total),0) FROM sales_data WHERE real_seller=? AND sale_date LIKE ?",
                (b['name'], f"{year}%")).fetchone()[0]
            if sales > best_sales: best_sales = sales; best = b

        # 나머지 브랜치에서 정보 수집하여 best에 병합
        def pick(vals): return next((v for v in vals if v and v.strip()), '')
        for b in group:
            if b['id'] == best['id']: continue
            # 연락처 정보 없는 쪽에서 있는 쪽으로 채우기
            updates = {}
            for field in ['ceo','ceo_phone','store_manager','store_manager_phone','manager','phone','address','email','region']:
                if not best.get(field) and b.get(field):
                    updates[field] = b[field]
            if updates:
                set_clause = ', '.join(f"{k}=?" for k in updates)
                conn.execute(f"UPDATE branches SET {set_clause} WHERE id=?",
                             list(updates.values()) + [best['id']])
                best.update(updates)
            # sales_data real_seller 업데이트
            conn.execute("UPDATE sales_data SET real_seller=? WHERE real_seller=?",
                         (best['name'], b['name']))
            conn.execute("DELETE FROM branches WHERE id=?", (b['id'],))
            merged += 1

    # 지역 자동 배정
    branches_no_region = conn.execute("SELECT id,name FROM branches WHERE region='' OR region IS NULL").fetchall()
    region_updated = 0
    for b in branches_no_region:
        region = detect_region_from_name(b["name"])
        if region:
            conn.execute("UPDATE branches SET region=? WHERE id=?", (region, b["id"]))
            region_updated += 1
    conn.commit(); conn.close()
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

# ── SNS 활용 매장 API ────────────────────────────────
@app.route("/api/sns/search", methods=["POST"])
@login_required
def api_sns_search():
    """네이버 블로그 검색으로 매장별 블로그 현황 자동 분석"""
    import urllib.request, urllib.parse, os, re
    from datetime import datetime as dt2

    sellers = request.json.get('sellers', [])
    if not sellers:
        return jsonify({'ok': False, 'msg': '매장명 없음'}), 400

    client_id     = os.environ.get('NAVER_CLIENT_ID',     'InqUUQfvWZN1rAZM4whk')
    client_secret = os.environ.get('NAVER_CLIENT_SECRET', 'fXYMLK1N1X')

    def strip_tags(s): return re.sub('<[^>]+>', '', s or '')

    def naver_blog(query, display=20, sort='date'):
        url = ('https://openapi.naver.com/v1/search/blog.json?query='
               + urllib.parse.quote(query) + f'&display={display}&sort={sort}')
        req = urllib.request.Request(url)
        req.add_header('X-Naver-Client-Id',     client_id)
        req.add_header('X-Naver-Client-Secret', client_secret)
        with urllib.request.urlopen(req, timeout=10) as r:
            return json.loads(r.read().decode('utf-8'))

    def parse_date(s):
        try: return f"{s[:4]}-{s[4:6]}-{s[6:8]}"
        except: return ''

    def calc_score(total, latest, recent_30d, has_product):
        score = 0
        # 총 게시글 수 (최대 30점)
        if total >= 200: score += 30
        elif total >= 100: score += 25
        elif total >= 50:  score += 20
        elif total >= 20:  score += 15
        elif total >= 5:   score += 10
        elif total >= 1:   score += 5
        # 최근 글 날짜 (최대 40점)
        if latest:
            try:
                days = (dt2.now() - dt2.strptime(latest, '%Y-%m-%d')).days
                if   days <= 7:   score += 40
                elif days <= 14:  score += 35
                elif days <= 30:  score += 28
                elif days <= 60:  score += 20
                elif days <= 90:  score += 12
                elif days <= 180: score += 6
                elif days <= 365: score += 2
            except: pass
        # 30일 이내 게시글 수 (최대 20점)
        if   recent_30d >= 20: score += 20
        elif recent_30d >= 10: score += 16
        elif recent_30d >= 5:  score += 12
        elif recent_30d >= 3:  score += 8
        elif recent_30d >= 1:  score += 4
        # 제품 관련 포스팅 (10점)
        if has_product: score += 10
        return min(score, 100)

    def grade(score):
        if score >= 80: return 'A'
        if score >= 60: return 'B'
        if score >= 40: return 'C'
        if score >= 20: return 'D'
        return 'E'

    PRODUCT_KEYWORDS = ['엔픽스','줄즈','레카로','원더폴드','카오스','타프토이즈',
                        '유모차','카시트','보행기','웨건','하이체어','유아용품']

    results = []
    for seller in sellers[:50]:
        res = {'seller_name': seller, 'ok': False, 'error': '',
               'blog_total': 0, 'blog_latest': '', 'blog_recent_30d': 0,
               'blog_score': 0, 'blog_grade': 'E', 'blog_platform': '',
               'blog_has_product_post': 0, 'blog_recent_titles': '', 'blog_keywords': ''}
        try:
            clean = seller.replace('_', ' ').strip()
            d = naver_blog(clean, 20, 'date')
            total = d.get('total', 0)
            items = d.get('items', [])

            # 최신 날짜
            dates = [parse_date(i.get('postdate','')) for i in items if i.get('postdate')]
            latest = dates[0] if dates else ''

            # 30일 이내 글 수
            now = dt2.now()
            recent_30d = 0
            for i in items:
                pd = parse_date(i.get('postdate',''))
                if pd:
                    try:
                        if (now - dt2.strptime(pd, '%Y-%m-%d')).days <= 30:
                            recent_30d += 1
                    except: pass

            # 제품 관련 포스팅 여부
            all_text = ' '.join(strip_tags(i.get('title','')) + strip_tags(i.get('description',''))
                                for i in items)
            has_product = any(kw in all_text for kw in PRODUCT_KEYWORDS)

            # 블로그 플랫폼 파악
            platforms = []
            for i in items:
                link = i.get('link', '')
                if 'blog.naver' in link: platforms.append('네이버')
                elif 'tistory' in link:  platforms.append('티스토리')
                elif 'brunch' in link:   platforms.append('브런치')
                elif 'instagram' in link: platforms.append('인스타')
                elif 'youtube' in link:  platforms.append('유튜브')
            from collections import Counter
            platform_str = ', '.join(f"{k}({v})" for k,v in Counter(platforms).most_common(3))

            # 최근 제목 3개
            titles = [strip_tags(i.get('title',''))[:30] for i in items[:3]]

            # 키워드 추출 (제목에서)
            all_titles = ' '.join(strip_tags(i.get('title','')) for i in items[:10])
            found_kws = [kw for kw in PRODUCT_KEYWORDS if kw in all_titles]

            score = calc_score(total, latest, recent_30d, has_product)

            res.update({
                'ok': True,
                'blog_total': total,
                'blog_latest': latest,
                'blog_recent_30d': recent_30d,
                'blog_has_product_post': 1 if has_product else 0,
                'blog_platform': platform_str,
                'blog_recent_titles': ' | '.join(titles),
                'blog_keywords': ', '.join(found_kws),
                'blog_score': score,
                'blog_grade': grade(score),
            })
        except urllib.error.HTTPError as e:
            res['error'] = f'HTTP {e.code}'
        except Exception as e:
            res['error'] = str(e)[:60]
        results.append(res)

    return jsonify({'results': results, 'ok': True})


@app.route("/api/sns/save-search", methods=["POST"])
@login_required
def api_sns_save_search():
    """검색 결과 DB 저장"""
    results = request.json.get('results', [])
    conn = get_db()
    updated = 0
    now = datetime.now().strftime('%Y-%m-%d %H:%M')
    for r in results:
        if not r.get('ok'): continue
        conn.execute("""INSERT INTO sns_info
            (seller_name, blog_total_posts, blog_latest_date, blog_recent_30d,
             blog_has_product_post, blog_platform, blog_recent_titles, blog_keywords,
             blog_score, blog_grade, last_searched, updated_at)
            VALUES(?,?,?,?,?,?,?,?,?,?,?,?)
            ON CONFLICT(seller_name) DO UPDATE SET
            blog_total_posts=excluded.blog_total_posts,
            blog_latest_date=excluded.blog_latest_date,
            blog_recent_30d=excluded.blog_recent_30d,
            blog_has_product_post=excluded.blog_has_product_post,
            blog_platform=excluded.blog_platform,
            blog_recent_titles=excluded.blog_recent_titles,
            blog_keywords=excluded.blog_keywords,
            blog_score=excluded.blog_score,
            blog_grade=excluded.blog_grade,
            last_searched=excluded.last_searched,
            updated_at=excluded.updated_at""",
            (r['seller_name'], r['blog_total'], r['blog_latest'], r['blog_recent_30d'],
             r['blog_has_product_post'], r['blog_platform'], r['blog_recent_titles'],
             r['blog_keywords'], r['blog_score'], r['blog_grade'], now, now))
        updated += 1
    conn.commit(); conn.close()
    return jsonify({'ok': True, 'updated': updated})


@app.route("/api/sns/list")
@login_required
def api_sns_list():
    year = request.args.get("year", str(datetime.now().year))
    conn = get_db()
    sellers = [r[0] for r in conn.execute(
        f"SELECT DISTINCT real_seller FROM sales_data "
        f"WHERE real_seller!='' AND sale_date LIKE '{year}%' ORDER BY real_seller").fetchall()]
    sns_map = {r['seller_name']: dict(r) for r in conn.execute("SELECT * FROM sns_info").fetchall()}
    sales_map = {r[0]: r[1] for r in conn.execute(
        f"SELECT real_seller, SUM(total) FROM sales_data "
        f"WHERE sale_date LIKE '{year}%' AND real_seller!='' GROUP BY real_seller").fetchall()}
    conn.close()
    result = []
    for s in sellers:
        info = sns_map.get(s, {})
        result.append({
            'seller_name':         s,
            'blog_url':            info.get('blog_url',''),
            'blog_name':           info.get('blog_name',''),
            'blog_platform':       info.get('blog_platform',''),
            'blog_total_posts':    info.get('blog_total_posts',0),
            'blog_latest_date':    info.get('blog_latest_date',''),
            'blog_recent_30d':     info.get('blog_recent_30d',0),
            'blog_has_product_post': info.get('blog_has_product_post',0),
            'blog_recent_titles':  info.get('blog_recent_titles',''),
            'blog_keywords':       info.get('blog_keywords',''),
            'blog_score':          info.get('blog_score',0),
            'blog_grade':          info.get('blog_grade',''),
            'last_searched':       info.get('last_searched',''),
            'memo':                info.get('memo',''),
            'year_sales':          sales_map.get(s,0),
        })
    result.sort(key=lambda x: -x['year_sales'])
    return jsonify(result)


@app.route("/api/sns/save-memo", methods=["POST"])
@login_required
def api_sns_save_memo():
    """메모 + 블로그 URL/이름 수동 저장"""
    d    = request.json or {}
    name = d.get('seller_name','').strip()
    if not name: return jsonify({'ok':False}), 400
    now  = datetime.now().strftime('%Y-%m-%d %H:%M')
    conn = get_db()
    conn.execute("""INSERT INTO sns_info(seller_name,blog_url,blog_name,memo,updated_at)
        VALUES(?,?,?,?,?)
        ON CONFLICT(seller_name) DO UPDATE SET
        blog_url=CASE WHEN ?!='' THEN ? ELSE blog_url END,
        blog_name=CASE WHEN ?!='' THEN ? ELSE blog_name END,
        memo=excluded.memo, updated_at=excluded.updated_at""",
        (name, d.get('blog_url',''), d.get('blog_name',''), d.get('memo',''), now,
         d.get('blog_url',''), d.get('blog_url',''),
         d.get('blog_name',''), d.get('blog_name','')))
    conn.commit(); conn.close()
    return jsonify({'ok': True})


# Render/gunicorn 실행 시 자동 초기화
init_db()

if __name__ == "__main__":
    import webbrowser, threading
    threading.Timer(1.2, lambda: webbrowser.open("http://localhost:5001")).start()
    app.run(debug=False, port=5001)
