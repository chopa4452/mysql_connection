import re
import pandas as pd
import mysql.connector
from collections import defaultdict

# 파일만 바꿔서 월별로 실행
EXCEL_PATH = "2025년_11월_자동차_등록자료_통계.xlsx"

# SSH 터널 기준: 127.0.0.1:3307 -> (EC2) -> RDS:3306
DB = {
    "host": "127.0.0.1",
    "port": 3307,
    "user": "admin",
    "password": "vmfhwprxm",
    "database": "SKN23",
}

ALLOWED_SIDO = {
    "강원","경기","경남","경북","광주","대구","대전","부산","서울","세종","울산","인천","전남","전북","제주","충남","충북"
}

VALID_VK = {"승용","승합","화물","특수","합계"}  # fact_flow_count의 vehicle_kind


def clean(x) -> str:
    if x is None:
        return ""
    s = str(x)
    s = s.replace("_x000D_", "").replace("\r", "").replace("\n", " ").strip()
    s = re.sub(r"\s+", " ", s)
    if s.lower() in ("nan", "none"):
        return ""
    return s

def to_int(x) -> int:
    if x is None:
        return 0
    try:
        if isinstance(x, float) and pd.isna(x):
            return 0
    except:
        pass
    s = clean(x).replace(",", "")
    if s in ("", "-"):
        return 0
    try:
        return int(float(s))
    except:
        return 0

def connect():
    return mysql.connector.connect(**DB)

def chunked(lst, size=2000):
    for i in range(0, len(lst), size):
        yield lst[i:i+size]

def parse_year_month_from_filename(excel_path: str):
    # 예: 2025년_09월_자동차_등록자료_통계.xlsx
    m = re.search(r"(\d{4})년[_\-\s]*(\d{1,2})월", excel_path)
    if not m:
        return None
    return int(m.group(1)), int(m.group(2))

def parse_year_month_strict(df: pd.DataFrame, excel_path: str):
    """
    핵심: '조회년월'이 들어간 셀에서만 YYYY.MM 추출
    못 찾으면 파일명에서 추출
    """
    # 1) 시트 상단에서 "조회년월" 포함 셀 찾기
    max_r = min(20, df.shape[0])
    max_c = min(20, df.shape[1])
    for r in range(max_r):
        for c in range(max_c):
            v = clean(df.iat[r, c])
            if "조회년월" in v:
                m = re.search(r"(\d{4})\.(\d{1,2})", v)
                if m:
                    return int(m.group(1)), int(m.group(2))

    # 2) 파일명 fallback
    ym = parse_year_month_from_filename(excel_path)
    if ym:
        return ym

    raise ValueError("조회년월을 찾지 못했습니다. (시트 상단 '조회년월: YYYY.MM' 또는 파일명 패턴 확인)")


# DIM MAP
def fetch_sido_map(cur):
    cur.execute("SELECT sido_id, sido_name FROM dim_region_sido")
    return {name: sid for (sid, name) in cur.fetchall()}

def fetch_subtype_map(cur):
    # key: (flow_type(group_name), subtype_name) -> subtype_id
    cur.execute("SELECT subtype_id, subtype_name, group_name FROM dim_flow_subtype")
    return {(g, n): sid for (sid, n, g) in cur.fetchall()}

def ensure_flow_subtypes(cur, needed_pairs):
    """
    needed_pairs: set of (group_name=flow_type, subtype_name)
    dim_flow_subtype에 없으면 자동 insert
    """
    existing = fetch_subtype_map(cur)
    missing = [(g, n) for (g, n) in needed_pairs if (g, n) not in existing]
    if not missing:
        return

    sql = """
    INSERT INTO dim_flow_subtype (subtype_name, group_name, is_inheritance, is_gift)
    VALUES (%s, %s, %s, %s)
    ON DUPLICATE KEY UPDATE
    is_inheritance = VALUES(is_inheritance),
    is_gift        = VALUES(is_gift)
    """
    rows = []
    for g, n in missing:
        is_inh = "Y" if "상속" in n else "N"
        is_gft = "Y" if "증여" in n else "N"
        rows.append((n, g, is_inh, is_gft))
    cur.executemany(sql, rows)


# EXTRACTORS (20~27)
def extract_flow_new(excel_path: str, sheet_name: str, is_cumulative: str):
    df = pd.read_excel(excel_path, sheet_name=sheet_name, header=None)
    year, month = parse_year_month_strict(df, excel_path)

    row_vehicle = [clean(x) for x in df.iloc[2, :].tolist()]  # 승용/승합/...
    row_subtype = [clean(x) for x in df.iloc[3, :].tolist()]  # 신조차/수입차/부활차/계

    col_info = []
    vk = ""
    for c in range(2, df.shape[1]):
        if row_vehicle[c]:
            vk = row_vehicle[c]
        st = row_subtype[c]
        if not vk or not st:
            continue
        if vk not in VALID_VK:
            continue
        if st == "계":  # dim에서 '계'를 빼는 설계면 여기서도 제외
            continue
        col_info.append((c, vk, st))

    recs = []
    for r in range(4, df.shape[0]):
        sido = clean(df.iat[r, 0])
        if sido not in ALLOWED_SIDO:
            continue
        for c, vehicle_kind, subtype_name in col_info:
            recs.append({
                "year": year, "month": month,
                "sido_name": sido,
                "flow_type": "신규",
                "subtype_name": subtype_name,
                "vehicle_kind": vehicle_kind,
                "is_cumulative": is_cumulative,
                "flow_count": to_int(df.iat[r, c]),
            })
    return recs

def extract_flow_simple(excel_path: str, sheet_name: str, flow_type: str, is_cumulative: str):
    df = pd.read_excel(excel_path, sheet_name=sheet_name, header=None)
    year, month = parse_year_month_strict(df, excel_path)

    header2 = [clean(x) for x in df.iloc[2, :].tolist()]  # 사유별, 주소변경...
    header3 = [clean(x) for x in df.iloc[3, :].tolist()]  # (특정 컬럼만 상세설명 들어감)

    col_info = []
    for c in range(2, df.shape[1]):
        st = header3[c] if header3[c] else header2[c]
        st = clean(st)
        if not st or st == "계":
            continue
        col_info.append((c, st))

    recs = []
    for r in range(4, df.shape[0]):
        sido = clean(df.iat[r, 0])
        if sido not in ALLOWED_SIDO:
            continue
        for c, subtype_name in col_info:
            recs.append({
                "year": year, "month": month,
                "sido_name": sido,
                "flow_type": flow_type,
                "subtype_name": subtype_name,
                "vehicle_kind": None,  # 변경/이전은 차종축 없음
                "is_cumulative": is_cumulative,
                "flow_count": to_int(df.iat[r, c]),
            })
    return recs

def extract_flow_malso(excel_path: str, sheet_name: str, is_cumulative: str):
    df = pd.read_excel(excel_path, sheet_name=sheet_name, header=None)
    year, month = parse_year_month_strict(df, excel_path)

    # row3: subtype들이 깔려있음 (col2~)
    subtype_row = [clean(x) for x in df.iloc[3, :].tolist()]

    # subtype명 -> cols (같은 이름이 여러 번 나오면 합산)
    st_to_cols = defaultdict(list)
    for c in range(2, df.shape[1]):
        st = clean(subtype_row[c])
        if not st or st == "계":
            continue
        st_to_cols[st].append(c)

    # data: row4~
    data = df.iloc[4:, :].copy()
    data["sido"] = data[0].map(clean).replace("", pd.NA).ffill()
    data["vehicle_kind"] = data[1].map(clean)

    # vehicle_kind 정규화
    data["vehicle_kind"] = data["vehicle_kind"].replace({"계": "합계"})
    data = data[data["sido"].isin(ALLOWED_SIDO)]
    data = data[data["vehicle_kind"].isin(VALID_VK)]

    recs = []
    for _, row in data.iterrows():
        sido = clean(row["sido"])
        vk = clean(row["vehicle_kind"])

        for st, cols in st_to_cols.items():
            cnt = 0
            for c in cols:
                cnt += to_int(row.get(c))
            recs.append({
                "year": year, "month": month,
                "sido_name": sido,
                "flow_type": "말소",
                "subtype_name": st,
                "vehicle_kind": vk,
                "is_cumulative": is_cumulative,
                "flow_count": cnt,
            })
    return recs


# --------------------------
# LOAD to FACT
# --------------------------
def load_fact_flow_count(cur, excel_path: str):
    sido_map = fetch_sido_map(cur)

    recs = []
    recs += extract_flow_new(excel_path, "20.신규 등록현황(당월)", is_cumulative="N")
    recs += extract_flow_new(excel_path, "21.신규 등록현황(누계)", is_cumulative="Y")
    recs += extract_flow_simple(excel_path, "22.변경 등록현황(당월)", flow_type="변경", is_cumulative="N")
    recs += extract_flow_simple(excel_path, "23.변경 등록현황(누계)", flow_type="변경", is_cumulative="Y")
    recs += extract_flow_simple(excel_path, "24.이전 등록현황(당월)", flow_type="이전", is_cumulative="N")
    recs += extract_flow_simple(excel_path, "25.이전 등록현황(누계)", flow_type="이전", is_cumulative="Y")
    recs += extract_flow_malso(excel_path, "26.말소 등록현황(당월)", is_cumulative="N")
    recs += extract_flow_malso(excel_path, "27.말소 등록현황(누계)", is_cumulative="Y")

    # dim_flow_subtype 자동 보강
    needed_pairs = {(r["flow_type"], r["subtype_name"]) for r in recs}
    ensure_flow_subtypes(cur, needed_pairs)

    subtype_map = fetch_subtype_map(cur)

    sql = """
    INSERT INTO fact_flow_count
    (year, month, sido_id, flow_type, subtype_id, vehicle_kind, is_cumulative, flow_count)
    VALUES (%s,%s,%s,%s,%s,%s,%s,%s)
    ON DUPLICATE KEY UPDATE
    flow_count = VALUES(flow_count)
    """

    data = []
    skipped = 0
    for r in recs:
        sid = sido_map.get(r["sido_name"])
        stid = subtype_map.get((r["flow_type"], r["subtype_name"]))
        if sid is None or stid is None:
            skipped += 1
            continue

        data.append((
            r["year"], r["month"],
            sid,
            r["flow_type"],
            stid,
            r["vehicle_kind"],      # 변경/이전은 None
            r["is_cumulative"],
            r["flow_count"]
        ))

    for part in chunked(data, 2000):
        cur.executemany(sql, part)

    # 실행 로그용: 어떤 월이 들어갔는지
    ym = sorted({(r["year"], r["month"]) for r in recs})
    return len(data), skipped, ym


def main():
    conn = connect()
    try:
        conn.start_transaction()
        cur = conn.cursor()

        inserted, skipped, ym = load_fact_flow_count(cur, EXCEL_PATH)

        conn.commit()
        print(f" fact_flow_count 적재 완료: inserted/updated={inserted}, skipped={skipped}, ym={ym}")

    except Exception as e:
        conn.rollback()
        print(" 실패:", e)
        raise
    finally:
        conn.close()


if __name__ == "__main__":
    main()
