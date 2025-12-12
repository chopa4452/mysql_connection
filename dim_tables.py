import re
import pandas as pd
import mysql.connector

EXCEL_PATH = "2025년_08월_자동차_등록자료_통계.xlsx"

#  SSH 터널 기준: 127.0.0.1:3307 -> (EC2) -> RDS:3306
DB = {
    "host": "127.0.0.1",
    "port": 3307,
    "user": "admin",
    "password": "vmfhwprxm",
    "database": "SKN23",
}



# 공통 유틸

def clean(x) -> str:
    s = "" if x is None else str(x)
    s = s.replace("_x000D_", "").replace("\r", "").replace("\n", " ").strip()
    s = re.sub(r"\s+", " ", s)
    return s


def is_numeric_str(s: str) -> bool:
    return bool(re.fullmatch(r"[0-9]+", s))


def connect():
    return mysql.connector.connect(**DB)


def normalize_yn(x: str, default="N") -> str:
    x = clean(x).upper()
    return "Y" if x == "Y" else ("N" if x == "N" else default)



# 적재(Upsert)

def upsert_dim_region_sido(cur, sidos):
    sql = """
    INSERT INTO dim_region_sido (sido_name, use_yn)
    VALUES (%s, %s)
    ON DUPLICATE KEY UPDATE use_yn = VALUES(use_yn)
    """
    data = [(name, "Y") for name in sidos if name]
    cur.executemany(sql, data)


def fetch_sido_map(cur):
    cur.execute("SELECT sido_id, sido_name FROM dim_region_sido")
    return {name: sid for (sid, name) in cur.fetchall()}


def upsert_dim_region_sigungu(cur, sido_map, pairs):
    sql = """
    INSERT INTO dim_region_sigungu (sido_id, sigungu_name, use_yn)
    VALUES (%s, %s, %s)
    ON DUPLICATE KEY UPDATE use_yn = VALUES(use_yn)
    """
    data = []
    for sido_name, sigungu_name in pairs:
        sid = sido_map.get(sido_name)
        if sid is None:
            continue
        if not sigungu_name:
            continue
        data.append((sid, sigungu_name, "Y"))
    cur.executemany(sql, data)


def upsert_dim_age_group(cur, age_rows):
    sql = """
    INSERT INTO dim_age_group (age_group, sort_order)
    VALUES (%s, %s)
    ON DUPLICATE KEY UPDATE sort_order = VALUES(sort_order)
    """
    cur.executemany(sql, age_rows)


def upsert_dim_fuel(cur, fuel_rows):
    sql = """
    INSERT INTO dim_fuel (fuel_name, is_eco)
    VALUES (%s, %s)
    ON DUPLICATE KEY UPDATE is_eco = VALUES(is_eco)
    """
    cur.executemany(sql, fuel_rows)


# ERD 기준: flow_type 컬럼 없음
def upsert_dim_flow_subtype(cur, rows):
    sql = """
    INSERT INTO dim_flow_subtype (subtype_name, group_name, is_inheritance, is_gift)
    VALUES (%s, %s, %s, %s)
    ON DUPLICATE KEY UPDATE
    is_inheritance = VALUES(is_inheritance),
    is_gift        = VALUES(is_gift)
    """
    cur.executemany(sql, rows)



# 엑셀 추출(라벨/잡값 강하게 제거)

def extract_regions(excel_path):
    df = pd.read_excel(excel_path, sheet_name="02.통계표_시군구")
    sido_col, sigungu_col = "Unnamed: 0", "Unnamed: 1"

    df[sido_col] = df[sido_col].ffill()
    sub = df[[sido_col, sigungu_col]].dropna(subset=[sigungu_col]).copy()

    sub[sido_col] = sub[sido_col].map(clean)
    sub[sigungu_col] = sub[sigungu_col].map(clean)

    # 허용 시도 17개만 통과(라벨/헤더 완전 차단)
    allowed_sido = {
        "강원","경기","경남","경북","광주","대구","대전","부산","서울","세종","울산","인천","전남","전북","제주","충남","충북"
    }
    sub = sub[sub[sido_col].isin(allowed_sido)]
    sub = sub[sub[sigungu_col].notna() & (sub[sigungu_col].str.strip() != "")]

    sidos = sorted(sub[sido_col].unique().tolist())
    pairs = list({(r[sido_col], r[sigungu_col]) for _, r in sub.iterrows()})
    pairs.sort()
    return sidos, pairs


def extract_age_groups(excel_path):
    df = pd.read_excel(excel_path, sheet_name="04.성별_연령별")
    age_col = "Unnamed: 1"
    ages = df[age_col].dropna().astype(str).map(clean)

    # 허용 연령대 10개만 통과(계/연령/시도 등 라벨 차단)
    allowed_age = {
        "10대이하","20대","30대","40대","50대","60대","70대","80대","90대이상","법인 및 사업자"
    }
    filtered = [a for a in ages.unique() if a in allowed_age]

    order = ["10대이하","20대","30대","40대","50대","60대","70대","80대","90대이상","법인 및 사업자"]
    order_map = {name: i for i, name in enumerate(order, start=1)}

    rows = [(a, order_map[a]) for a in filtered]
    rows.sort(key=lambda x: x[1])
    return rows


def extract_fuels(excel_path):
    df = pd.read_excel(excel_path, sheet_name="10.연료별_등록현황")
    fuel_col = "Unnamed: 0"
    fuels = df[fuel_col].dropna().astype(str).map(clean)

    # 허용 연료 17개만 통과(총계/연료별 라벨 차단)
    allowed_fuel = {
        "CNG","LNG","경유","기타연료","등유","수소","수소전기","알코올","엘피지","전기","태양열",
        "하이브리드(CNG+전기)","하이브리드(LNG+전기)","하이브리드(LPG+전기)",
        "하이브리드(경유+전기)","하이브리드(휘발유+전기)","휘발유"
    }
    unique = [f for f in fuels.unique() if f in allowed_fuel]

    eco_keywords = ["전기", "수소", "하이브리드"]

    def eco_flag(name):
        return "Y" if any(k in name for k in eco_keywords) else "N"

    rows = [(f, eco_flag(f)) for f in sorted(set(unique))]
    return rows


def extract_flow_subtypes(excel_path):
    # 시도명/차종이 헤더에서 subtype로 섞여 들어오는 걸 차단
    sido_names = {
        "강원","경기","경남","경북","광주","대구","대전","부산","서울","세종","울산","인천","전남","전북","제주","충남","충북"
    }
    vehicle_types = {"승용", "승합", "화물", "특수"}

    stopwords_exact = {
        "조회년월:", "시도별", "차종별", "사유별", "계", "총계", "합계",
        "자동차", "통계표", "등록현황", "(당월계)", "(누계)", "<", ">", "종별", "용도별", "연료별", "시도", "시군구",
        "말소", "변경", "이전", "신규",
    }
    drop_like = {"시도를 달리하는", "직권말소"}

    def head_strings(sheet, nrows=6):
        df = pd.read_excel(excel_path, sheet_name=sheet, header=None, nrows=nrows)
        out = []
        for v in df.values.ravel():
            if pd.isna(v):
                continue
            s = clean(v)

            if not s:
                continue

            # 추가된 핵심 필터(시도/차종 차단)
            if s in sido_names:
                continue
            if s in vehicle_types:
                continue

            if s in stopwords_exact:
                continue
            if s in drop_like:
                continue
            if "조회년월" in s:
                continue
            if re.fullmatch(r"\d{4}\.\d{2}", s) or is_numeric_str(s):
                continue
            if len(s) <= 1:
                continue

            out.append(s)
        return out

    groups = {
        "신규": ["20.신규 등록현황(당월)", "21.신규 등록현황(누계)"],
        "변경": ["22.변경 등록현황(당월)", "23.변경 등록현황(누계)"],
        "이전": ["24.이전 등록현황(당월)", "25.이전 등록현황(누계)"],
        "말소": ["26.말소 등록현황(당월)", "27.말소 등록현황(누계)"],
    }

    rows = []
    seen = set()

    for gname, sheets in groups.items():
        candidates = []
        for sh in sheets:
            candidates.extend(head_strings(sh, nrows=6))

        for c in candidates:
            subtype = clean(c)
            if not subtype:
                continue

            key = (subtype, gname)
            if key in seen:
                continue

            is_inh = "Y" if "상속" in subtype else "N"
            is_gft = "Y" if "증여" in subtype else "N"

            rows.append((subtype, gname, is_inh, is_gft))
            seen.add(key)

    rows.sort(key=lambda x: (x[1], x[0]))
    return rows



# (선택) 초기화: 이미 들어간 DIM을 깨끗하게 지우고 다시 넣고 싶을 때

def reset_dim_tables(cur):
    cur.execute("SET FOREIGN_KEY_CHECKS=0")
    # FK 때문에 시군구 먼저
    cur.execute("TRUNCATE TABLE dim_region_sigungu")
    cur.execute("TRUNCATE TABLE dim_flow_subtype")
    cur.execute("TRUNCATE TABLE dim_fuel")
    cur.execute("TRUNCATE TABLE dim_age_group")
    cur.execute("TRUNCATE TABLE dim_region_sido")
    cur.execute("SET FOREIGN_KEY_CHECKS=1")


# --------------------------
# 적재 후 교차검증(팩트 출력)
# --------------------------
def post_checks(cur):
    # 1) 중복 체크(유니크 있으면 원칙적으로 0)
    checks = [
        ("dup_sido", """
            SELECT COUNT(*) FROM (
            SELECT sido_name FROM dim_region_sido GROUP BY sido_name HAVING COUNT(*) > 1
            ) t
        """),
        ("dup_age", """
            SELECT COUNT(*) FROM (
            SELECT age_group FROM dim_age_group GROUP BY age_group HAVING COUNT(*) > 1
            ) t
        """),
        ("dup_fuel", """
            SELECT COUNT(*) FROM (
            SELECT fuel_name FROM dim_fuel GROUP BY fuel_name HAVING COUNT(*) > 1
            ) t
        """),
        ("dup_flow_subtype", """
            SELECT COUNT(*) FROM (
            SELECT subtype_name, group_name FROM dim_flow_subtype
            GROUP BY subtype_name, group_name HAVING COUNT(*) > 1
            ) t
        """),

        # 2) NULL/공백 체크
        ("null_or_blank_sido", """
            SELECT COUNT(*) FROM dim_region_sido
            WHERE sido_name IS NULL OR TRIM(sido_name) = ''
        """),
        ("null_or_blank_sigungu", """
            SELECT COUNT(*) FROM dim_region_sigungu
            WHERE sigungu_name IS NULL OR TRIM(sigungu_name) = ''
        """),
        ("null_or_blank_age", """
            SELECT COUNT(*) FROM dim_age_group
            WHERE age_group IS NULL OR TRIM(age_group) = ''
        """),
        ("null_or_blank_fuel", """
            SELECT COUNT(*) FROM dim_fuel
            WHERE fuel_name IS NULL OR TRIM(fuel_name) = ''
        """),
        ("null_or_blank_flow_subtype", """
            SELECT COUNT(*) FROM dim_flow_subtype
            WHERE subtype_name IS NULL OR TRIM(subtype_name) = ''
            OR group_name IS NULL
        """),

        # 3) 라벨 섞임(대표 라벨만)
        ("label_in_sido", """
            SELECT COUNT(*) FROM dim_region_sido
            WHERE sido_name IN ('계','합계','총계','시도별','연료별','성별','연령/시도','월간증감','년간증감')
        """),
        ("label_in_age", """
            SELECT COUNT(*) FROM dim_age_group
            WHERE age_group IN ('계','합계','총계','연령/시도','성별')
        """),
        ("label_in_fuel", """
            SELECT COUNT(*) FROM dim_fuel
            WHERE fuel_name IN ('계','합계','총계','연료별')
        """),
    ]

    results = {}
    for name, q in checks:
        cur.execute(q)
        results[name] = cur.fetchone()[0]

    # 행수 카운트
    cur.execute("""
        SELECT 'dim_region_sido' AS tbl, COUNT(*) cnt FROM dim_region_sido
        UNION ALL SELECT 'dim_region_sigungu', COUNT(*) FROM dim_region_sigungu
        UNION ALL SELECT 'dim_age_group', COUNT(*) FROM dim_age_group
        UNION ALL SELECT 'dim_fuel', COUNT(*) FROM dim_fuel
        UNION ALL SELECT 'dim_flow_subtype', COUNT(*) FROM dim_flow_subtype
    """)
    counts = cur.fetchall()

    return counts, results


def main():
    # 엑셀 추출
    sidos, sigungu_pairs = extract_regions(EXCEL_PATH)
    age_rows = extract_age_groups(EXCEL_PATH)
    fuel_rows = extract_fuels(EXCEL_PATH)
    flow_rows = extract_flow_subtypes(EXCEL_PATH)

    conn = connect()
    try:
        conn.start_transaction()
        cur = conn.cursor()

        # 정말 깨끗하게 다시 넣고 싶으면 아래 1줄 주석 해제
        # reset_dim_tables(cur)

        upsert_dim_region_sido(cur, sidos)
        sido_map = fetch_sido_map(cur)
        upsert_dim_region_sigungu(cur, sido_map, sigungu_pairs)

        upsert_dim_age_group(cur, age_rows)
        upsert_dim_fuel(cur, fuel_rows)
        upsert_dim_flow_subtype(cur, flow_rows)

        # 커밋
        conn.commit()

        # 교차검증 출력(팩트)
        counts, checks = post_checks(cur)

        print("DONE  DIM 테이블 적재 완료")
        for tbl, cnt in counts:
            print(f"- {tbl}: {cnt}")

        print("\n[VALIDATION]")
        for k, v in checks.items():
            print(f"- {k}: {v}")

    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


if __name__ == "__main__":
    main()
