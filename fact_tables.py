import re
import pandas as pd
import mysql.connector
from collections import defaultdict

EXCEL_PATH = "2024년_10월_자동차_등록자료_통계.xlsx"

# SSH 터널 기준: 127.0.0.1:3307 -> (EC2) -> RDS:3306
DB = {
    "host": "127.0.0.1",
    "port": 3307,
    "user": "admin",
    "password": "vmfhwprxm",
    "database": "SKN23",
}

# --------------------------
# 공통 유틸
# --------------------------
def clean(x) -> str:
    s = "" if x is None else str(x)
    s = s.replace("_x000D_", "").replace("\r", "").replace("\n", " ").strip()
    s = re.sub(r"\s+", " ", s)
    return s

def to_int(x) -> int:
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return 0
    s = clean(x).replace(",", "")
    if s == "" or s == "-":
        return 0
    try:
        return int(float(s))
    except:
        return 0

def connect():
    return mysql.connector.connect(**DB)

ALLOWED_SIDO = {
    "강원","경기","경남","경북","광주","대구","대전","부산","서울","세종","울산","인천","전남","전북","제주","충남","충북"
}

def parse_year_month(df: pd.DataFrame) -> tuple[int, int]:
    # 상단에서 YYYY.MM 패턴 찾기
    for r in range(min(8, df.shape[0])):
        for c in range(min(10, df.shape[1])):
            v = clean(df.iat[r, c])
            m = re.search(r"(\d{4})\.(\d{2})", v)
            if m:
                return int(m.group(1)), int(m.group(2))
    raise ValueError("조회년월(YYYY.MM)을 찾지 못했습니다.")

def chunked(lst, size=2000):
    for i in range(0, len(lst), size):
        yield lst[i:i+size]


# DIM MAP 로딩

def fetch_maps(cur):
    cur.execute("SELECT sido_id, sido_name FROM dim_region_sido")
    sido_map = {name: sid for sid, name in cur.fetchall()}

    cur.execute("SELECT fuel_id, fuel_name FROM dim_fuel")
    fuel_map = {name: fid for fid, name in cur.fetchall()}

    cur.execute("SELECT age_group_id, age_group FROM dim_age_group")
    age_map = {name: aid for aid, name in cur.fetchall()}

    cur.execute("SELECT subtype_id, subtype_name, group_name FROM dim_flow_subtype")
    subtype_map = {(g, n): sid for sid, n, g in cur.fetchall()}

    return sido_map, fuel_map, age_map, subtype_map

def ensure_missing_flow_subtypes(cur, subtype_map, need_pairs):
    # need_pairs: set[(group_name, subtype_name)]
    missing = [(g, n) for (g, n) in need_pairs if (g, n) not in subtype_map]
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


# 1) fact_vehicle_stock (01/02/03)

def find_row_by_token(df, token: str) -> int:
    for r in range(df.shape[0]):
        left = " ".join([clean(df.iat[r, c]) for c in range(0, 5)])
        if token in left:
            return r
    raise ValueError(f"'{token}' 행을 못 찾음")

def extract_vehicle_stock_usage(sheet_name: str, usage_type: str, origin_type: str = "전체"):
    df = pd.read_excel(EXCEL_PATH, sheet_name=sheet_name, header=None)
    year, month = parse_year_month(df)

    # 시도 헤더: row2, col5~ (마지막 '합계' 제외)
    header_row = 2
    sido_cols = []
    for c in range(5, df.shape[1]):
        s = clean(df.iat[header_row, c])
        if s in ALLOWED_SIDO:
            sido_cols.append((c, s))

    # 필요한 행
    r_total = find_row_by_token(df, "총계")
    r_suv   = find_row_by_token(df, "승용차합계")
    r_van   = find_row_by_token(df, "승합차 합계")
    r_trk   = find_row_by_token(df, "화물자동차 합계")
    r_spc   = find_row_by_token(df, "특수자동차 합계")

    row_map = [
        ("합계", r_total),
        ("승용", r_suv),
        ("승합", r_van),
        ("화물", r_trk),
        ("특수", r_spc),
    ]

    recs = []
    for vehicle_kind, rr in row_map:
        for cc, sido in sido_cols:
            recs.append({
                "year": year, "month": month,
                "origin_type": origin_type,
                "sido_name": sido,
                "vehicle_kind": vehicle_kind,
                "usage_type": usage_type,
                "stock_count": to_int(df.iat[rr, cc]),
            })
    return recs

def extract_import_vehicle_stock_from_sheet03():
    df = pd.read_excel(EXCEL_PATH, sheet_name="03.수입차_시군구", header=None)
    year, month = parse_year_month(df)

    data = df.iloc[4:, 0:7].copy()
    data.columns = ["sido","sigungu","승용","승합","화물","특수","합계"]
    data["sido"] = data["sido"].map(clean).replace("", pd.NA).ffill()

    for col in ["승용","승합","화물","특수","합계"]:
        data[col] = data[col].apply(to_int)

    # 시/도 단위로만 합치기
    agg = data.groupby("sido")[["승용","승합","화물","특수","합계"]].sum().reset_index()

    recs = []
    for _, r in agg.iterrows():
        sido = clean(r["sido"])
        if sido not in ALLOWED_SIDO:
            continue
        for vk in ["승용","승합","화물","특수","합계"]:
            recs.append({
                "year": year, "month": month,
                "origin_type": "수입차",
                "sido_name": sido,
                "vehicle_kind": vk,
                "usage_type": "계",
                "stock_count": int(r[vk]),
            })
    return recs

def load_fact_vehicle_stock(cur, sido_map):
    recs = []
    recs += extract_vehicle_stock_usage("05.차종별_등록현황(전체)", usage_type="계", origin_type="전체")
    recs += extract_vehicle_stock_usage("06.차종별_등록현황(관용)", usage_type="관용", origin_type="전체")
    recs += extract_vehicle_stock_usage("07.차종별_등록현황(자가용)", usage_type="자가용", origin_type="전체")
    recs += extract_vehicle_stock_usage("08.차종별 등록현황(영업용)", usage_type="영업용", origin_type="전체")
    recs += extract_import_vehicle_stock_from_sheet03()

    sql = """
    INSERT INTO fact_vehicle_stock
      (year, month, origin_type, sido_id, vehicle_kind, usage_type, stock_count)
    VALUES (%s,%s,%s,%s,%s,%s,%s)
    ON DUPLICATE KEY UPDATE stock_count = VALUES(stock_count)
    """
    data = []
    for r in recs:
        sid = sido_map.get(r["sido_name"])
        if sid is None:
            continue
        data.append((r["year"], r["month"], r["origin_type"], sid, r["vehicle_kind"], r["usage_type"], r["stock_count"]))

    for part in chunked(data, 2000):
        cur.executemany(sql, part)


# 2) fact_owner_demo_stock (04)

def extract_owner_demo_stock():
    df = pd.read_excel(EXCEL_PATH, sheet_name="04.성별_연령별", header=None)
    year, month = parse_year_month(df)

    # row2: 성별/연령/총계/서울...
    header = [clean(x) for x in df.iloc[2, :].tolist()]
    sido_cols = [(i, header[i]) for i in range(3, len(header)) if header[i] in ALLOWED_SIDO]

    data = df.iloc[3:, :].copy()
    data["gender"] = data[0].map(clean).replace("", pd.NA).ffill()
    data["age_group"] = data[1].map(clean)

    # 합계/계 제외
    data = data[(data["gender"].isin(["남성","여성","기타"])) & (data["age_group"] != "계")]

    recs = []
    for _, r in data.iterrows():
        g = clean(r["gender"])
        a = clean(r["age_group"])
        for col, sido in sido_cols:
            recs.append({
                "year": year, "month": month,
                "sido_name": sido,
                "gender": g,
                "age_group": a,
                "stock_count": to_int(r[col])
            })
    return recs

def load_fact_owner_demo_stock(cur, sido_map, age_map):
    recs = extract_owner_demo_stock()

    sql = """
    INSERT INTO fact_owner_demo_stock
      (year, month, sido_id, gender, age_group_id, stock_count)
    VALUES (%s,%s,%s,%s,%s,%s)
    ON DUPLICATE KEY UPDATE stock_count = VALUES(stock_count)
    """
    data = []
    for r in recs:
        sid = sido_map.get(r["sido_name"])
        aid = age_map.get(r["age_group"])
        if sid is None or aid is None:
            continue
        data.append((r["year"], r["month"], sid, r["gender"], aid, r["stock_count"]))

    for part in chunked(data, 2000):
        cur.executemany(sql, part)


# 3) fact_fuel_stock (10)

def extract_fuel_stock():
    df = pd.read_excel(EXCEL_PATH, sheet_name="10.연료별_등록현황", header=None)
    year, month = parse_year_month(df)

    header = [clean(x) for x in df.iloc[2, :].tolist()]
    sido_cols = [(i, header[i]) for i in range(3, len(header)) if header[i] in ALLOWED_SIDO]

    data = df.iloc[4:, :].copy()
    data["fuel"] = data[0].map(clean).replace("", pd.NA).ffill()
    data["vehicle_kind"] = data[1].map(clean)
    data["business_type"] = data[2].map(clean)

    # 의미없는 줄 제거
    data = data[(data["vehicle_kind"] != "") & (data["business_type"] != "")]

    recs = []
    for _, r in data.iterrows():
        fuel = clean(r["fuel"])
        vk = clean(r["vehicle_kind"])
        bt = clean(r["business_type"])

        # sheet에서 '계'는 합계로 통일
        if vk == "계":
            vk = "합계"

        for col, sido in sido_cols:
            recs.append({
                "year": year, "month": month,
                "sido_name": sido,
                "fuel": fuel,
                "vehicle_kind": vk,
                "business_type": bt,
                "stock_count": to_int(r[col])
            })
    return recs

def load_fact_fuel_stock(cur, sido_map, fuel_map):
    recs = extract_fuel_stock()

    sql = """
    INSERT INTO fact_fuel_stock
      (year, month, sido_id, fuel_id, vehicle_kind, business_type, stock_count)
    VALUES (%s,%s,%s,%s,%s,%s,%s)
    ON DUPLICATE KEY UPDATE stock_count = VALUES(stock_count)
    """
    data = []
    for r in recs:
        sid = sido_map.get(r["sido_name"])
        fid = fuel_map.get(r["fuel"])
        if sid is None or fid is None:
            continue
        data.append((r["year"], r["month"], sid, fid, r["vehicle_kind"], r["business_type"], r["stock_count"]))

    for part in chunked(data, 2000):
        cur.executemany(sql, part)


# 4) fact_flow_count (20~27)
def extract_flow_new(sheet_name: str, is_cumulative: str):
    df = pd.read_excel(EXCEL_PATH, sheet_name=sheet_name, header=None)
    year, month = parse_year_month(df)

    row_vehicle = [clean(x) for x in df.iloc[2, :].tolist()]  # 승용/승합/...
    row_subtype = [clean(x) for x in df.iloc[3, :].tolist()]  # 신조차/수입차/부활차/계

    # vehicle kind ffill
    vk = ""
    col_info = []  # (col, vehicle_kind, subtype)
    for c in range(2, df.shape[1]):
        if row_vehicle[c] != "":
            vk = row_vehicle[c]
        st = row_subtype[c]
        if vk == "" or st == "":
            continue
        if st == "계":
            continue  # dim에서 계를 빼놨으니 fact에서도 제외
        col_info.append((c, vk, st))

    recs = []
    for r in range(4, df.shape[0]):
        sido = clean(df.iat[r, 0])
        if sido not in ALLOWED_SIDO:
            continue
        for c, vehicle_kind, subtype in col_info:
            recs.append({
                "year": year, "month": month,
                "sido_name": sido,
                "flow_type": "신규",
                "subtype_name": subtype,
                "vehicle_kind": vehicle_kind,   # 승용/승합/화물/특수/합계
                "is_cumulative": is_cumulative,
                "flow_count": to_int(df.iat[r, c])
            })
    return recs

def extract_flow_simple(sheet_name: str, flow_type: str, is_cumulative: str):
    df = pd.read_excel(EXCEL_PATH, sheet_name=sheet_name, header=None)
    year, month = parse_year_month(df)

    header = [clean(x) for x in df.iloc[2, :].tolist()]  # subtype들
    cols = []
    subtypes = []
    for c in range(2, len(header)):
        st = header[c]
        if st == "" or st == "계":
            continue
        cols.append(c)
        subtypes.append(st)

    recs = []
    for r in range(4, df.shape[0]):
        sido = clean(df.iat[r, 0])
        if sido not in ALLOWED_SIDO:
            continue
        for c, st in zip(cols, subtypes):
            recs.append({
                "year": year, "month": month,
                "sido_name": sido,
                "flow_type": flow_type,
                "subtype_name": st,
                "vehicle_kind": None,          # 변경/이전은 차종축 없음
                "is_cumulative": is_cumulative,
                "flow_count": to_int(df.iat[r, c])
            })
    return recs

def extract_flow_malso(sheet_name: str, is_cumulative: str):
    df = pd.read_excel(EXCEL_PATH, sheet_name=sheet_name, header=None)
    year, month = parse_year_month(df)

    # row3에 subtype 이름들이 쫙 있음 (col2~)
    subtype_row = [clean(x) for x in df.iloc[3, :].tolist()]

    subtype_cols = []
    for c in range(2, df.shape[1]):
        st = subtype_row[c]
        if st == "" or st == "계":
            continue
        subtype_cols.append((c, st))

    data = df.iloc[4:, :].copy()
    data["sido"] = data[0].map(clean).replace("", pd.NA).ffill()
    data["vehicle_kind"] = data[1].map(clean)

    valid_vk = {"승용","승합","화물","특수","계"}
    data = data[data["vehicle_kind"].isin(valid_vk)]
    data = data[data["sido"].isin(ALLOWED_SIDO)]

    # dim_flow_subtype가 (subtype_name, group_name)로 중복 제거되어 있으니,
    # 말소 내에서 같은 subtype_name이 두 번(자진/직권) 나올 경우 "합쳐서" 넣음
    agg = defaultdict(int)  # key -> sum
    for _, r in data.iterrows():
        sido = clean(r["sido"])
        vk = clean(r["vehicle_kind"])
        if vk == "계":
            vk = "합계"

        for c, st in subtype_cols:
            key = (year, month, sido, vk, st, is_cumulative)
            agg[key] += to_int(r[c])

    recs = []
    for (yy, mm, sido, vk, st, ic), cnt in agg.items():
        recs.append({
            "year": yy, "month": mm,
            "sido_name": sido,
            "flow_type": "말소",
            "subtype_name": st,
            "vehicle_kind": vk,
            "is_cumulative": ic,
            "flow_count": cnt
        })
    return recs

def load_fact_flow_count(cur, sido_map, subtype_map):
    recs = []
    recs += extract_flow_new("20.신규 등록현황(당월)", is_cumulative="N")
    recs += extract_flow_new("21.신규 등록현황(누계)", is_cumulative="Y")
    recs += extract_flow_simple("22.변경 등록현황(당월)", flow_type="변경", is_cumulative="N")
    recs += extract_flow_simple("23.변경 등록현황(누계)", flow_type="변경", is_cumulative="Y")
    recs += extract_flow_simple("24.이전 등록현황(당월)", flow_type="이전", is_cumulative="N")
    recs += extract_flow_simple("25.이전 등록현황(누계)", flow_type="이전", is_cumulative="Y")
    recs += extract_flow_malso("26.말소 등록현황(당월)", is_cumulative="N")
    recs += extract_flow_malso("27.말소 등록현황(누계)", is_cumulative="Y")

    # subtype_id 매핑이 없는 값이 있으면 dim_flow_subtype에 자동 보강(안정성)
    need_pairs = {(r["flow_type"], r["subtype_name"]) for r in recs}
    ensure_missing_flow_subtypes(cur, subtype_map, need_pairs)

    # 최신 맵 다시 로딩
    cur.execute("SELECT subtype_id, subtype_name, group_name FROM dim_flow_subtype")
    subtype_map = {(g, n): sid for sid, n, g in cur.fetchall()}

    sql = """
    INSERT INTO fact_flow_count
      (year, month, sido_id, flow_type, subtype_id, vehicle_kind, is_cumulative, flow_count)
    VALUES (%s,%s,%s,%s,%s,%s,%s,%s)
    ON DUPLICATE KEY UPDATE flow_count = VALUES(flow_count)
    """
    data = []
    for r in recs:
        sid = sido_map.get(r["sido_name"])
        stid = subtype_map.get((r["flow_type"], r["subtype_name"]))
        if sid is None or stid is None:
            continue
        data.append((r["year"], r["month"], sid, r["flow_type"], stid, r["vehicle_kind"], r["is_cumulative"], r["flow_count"]))

    for part in chunked(data, 2000):
        cur.executemany(sql, part)

# --------------------------
# 메인
# --------------------------
def main():
    conn = connect()
    try:
        conn.start_transaction()
        cur = conn.cursor()

        sido_map, fuel_map, age_map, subtype_map = fetch_maps(cur)

        # FACT 적재
        load_fact_vehicle_stock(cur, sido_map)
        load_fact_owner_demo_stock(cur, sido_map, age_map)
        load_fact_fuel_stock(cur, sido_map, fuel_map)
        load_fact_flow_count(cur, sido_map, subtype_map)

        conn.commit()
        print("DONE  FACT 테이블 적재 완료")

    except Exception as e:
        conn.rollback()
        print(" 실패:", e)
        raise
    finally:
        conn.close()

if __name__ == "__main__":
    main()
