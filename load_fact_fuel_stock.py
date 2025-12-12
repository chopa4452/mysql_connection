import re
import pandas as pd
import mysql.connector

EXCEL_PATH = "2024년_10월_자동차_등록자료_통계.xlsx"

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

SKIP_TOTAL = {"총계", "합계", "계", "Total", "TOTAL"}

# 유틸
def connect():
    return mysql.connector.connect(**DB)

def clean(x) -> str:
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return ""
    s = str(x)
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

def parse_year_month_from_filename(path: str) -> tuple[int, int]:
    # 예: 2025년_10월_자동차_등록자료_통계.xlsx
    m = re.search(r"(\d{4})\s*년[_\-\s]*(\d{1,2})\s*월", path)
    if not m:
        raise ValueError("파일명에서 'YYYY년_MM월' 패턴을 찾지 못했습니다.")
    return int(m.group(1)), int(m.group(2))

def chunked(lst, size=2000):
    for i in range(0, len(lst), size):
        yield lst[i:i+size]

# DIM 맵 로딩
def fetch_sido_map(cur):
    cur.execute("SELECT sido_id, sido_name FROM dim_region_sido")
    return {name: sid for (sid, name) in cur.fetchall()}

def fetch_fuel_map(cur):
    cur.execute("SELECT fuel_id, fuel_name FROM dim_fuel")
    return {name: fid for (fid, name) in cur.fetchall()}


# 표준화(옵션)
def normalize_vehicle_kind(vk: str) -> str:
    vk = clean(vk)
    # 엑셀에 '계'가 있으면 fact에는 '합계'로 통일
    if vk == "계":
        return "합계"
    return vk

def normalize_business_type(bt: str) -> str:
    bt = clean(bt)
    # 엑셀 표기 변형 대비(필요시 확장)
    if bt in {"계", "합계", "총계"}:
        return "계"
    return bt


# 추출: 10.연료별_등록현황
def extract_fuel_stock(excel_path: str):
    df = pd.read_excel(excel_path, sheet_name="10.연료별_등록현황", header=None)
    year, month = parse_year_month_from_filename(excel_path)

    # 시도 헤더는 보통 row=2, col=3~ (0=연료,1=차종,2=사업구분,3~=시도)
    header = [clean(x) for x in df.iloc[2, :].tolist()]
    sido_cols = [(i, header[i]) for i in range(3, len(header)) if header[i] in ALLOWED_SIDO]
    if not sido_cols:
        raise ValueError("[10.연료별_등록현황] 시도 헤더(col>=3)에서 시도명을 못 찾음")

    data = df.iloc[4:, :].copy()

    # 연료는 위에서 내려오는 구조라 ffill 필요
    data["fuel"] = data[0].map(clean).replace("", pd.NA).ffill()
    data["vehicle_kind"] = data[1].map(clean)
    data["business_type"] = data[2].map(clean)

    recs = []
    for _, r in data.iterrows():
        fuel = clean(r["fuel"])
        vk = normalize_vehicle_kind(r["vehicle_kind"])
        bt = normalize_business_type(r["business_type"])

        if not fuel or fuel in SKIP_TOTAL:
            continue
        if not vk or vk in SKIP_TOTAL:
            continue
        if not bt or bt in SKIP_TOTAL:
            continue

        # 차종은 4종+합계만 허용(원하면 더 풀어도 됨)
        if vk not in {"승용", "승합", "화물", "특수", "합계"}:
            continue

        for col_idx, sido_name in sido_cols:
            recs.append({
                "year": year,
                "month": month,
                "sido_name": sido_name,
                "fuel": fuel,
                "vehicle_kind": vk,
                "business_type": bt,
                "stock_count": to_int(r[col_idx])
            })

    return year, month, recs

# 적재
def load_fact_fuel_stock(cur, excel_path: str):
    sido_map = fetch_sido_map(cur)
    fuel_map = fetch_fuel_map(cur)

    year, month, recs = extract_fuel_stock(excel_path)
    print(f"[INFO] 대상월: {year}-{month:02d}, 추출 rows={len(recs)}")

    # 해당 월만 재적재(다른 달은 누적 유지)
    cur.execute("DELETE FROM fact_fuel_stock WHERE year=%s AND month=%s", (year, month))

    sql = """
    INSERT INTO fact_fuel_stock
      (year, month, sido_id, fuel_id, vehicle_kind, business_type, stock_count)
    VALUES (%s,%s,%s,%s,%s,%s,%s)
    ON DUPLICATE KEY UPDATE stock_count = VALUES(stock_count)
    """

    data = []
    missing_sido = set()
    missing_fuel = set()

    for r in recs:
        sid = sido_map.get(r["sido_name"])
        fid = fuel_map.get(r["fuel"])
        if sid is None:
            missing_sido.add(r["sido_name"])
            continue
        if fid is None:
            missing_fuel.add(r["fuel"])
            continue
        data.append((r["year"], r["month"], sid, fid, r["vehicle_kind"], r["business_type"], r["stock_count"]))

    if missing_sido:
        print("[WARN] dim_region_sido 매칭 실패:", sorted(missing_sido))
    if missing_fuel:
        print("[WARN] dim_fuel 매칭 실패:", sorted(missing_fuel))

    for part in chunked(data, 2000):
        cur.executemany(sql, part)

    print(f"[OK] fact_fuel_stock 적재 완료: {len(data)} rows (월={year}-{month:02d})")

def main():
    conn = connect()
    try:
        conn.start_transaction()
        cur = conn.cursor()

        load_fact_fuel_stock(cur, EXCEL_PATH)

        conn.commit()
        print("DONE")
    except Exception as e:
        conn.rollback()
        print("실패:", e)
        raise
    finally:
        conn.close()

if __name__ == "__main__":
    main()
