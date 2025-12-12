import re
import pandas as pd
import mysql.connector


# 설정
EXCEL_PATH = "2024년_10월_자동차_등록자료_통계.xlsx"

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

# 이 스크립트는 '시도'까지만 사용 (sigungu 없음)
# 사용 시트:
# - 05.차종별_등록현황(전체)  -> usage_type='계'
# - 06.차종별_등록현황(관용)  -> usage_type='관용'
# - 07.차종별_등록현황(자가용)-> usage_type='자가용'
# - 08.차종별 등록현황(영업용)-> usage_type='영업용'
# - 03.수입차_시군구 -> 시도 단위로 합쳐서 origin_type='수입차', usage_type='계'
SHEETS_USAGE = [
    ("05.차종별_등록현황(전체)", "계"),
    ("06.차종별_등록현황(관용)", "관용"),
    ("07.차종별_등록현황(자가용)", "자가용"),
    ("08.차종별 등록현황(영업용)", "영업용"),
]

IMPORT_SHEET = "03.수입차_시군구"


# 공통 유틸
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

def parse_year_month(df: pd.DataFrame) -> tuple[int, int]:
    """
    상단에서 YYYY.MM 또는 YYYY.M 또는 '2025년 10월' 형태를 찾아 year, month 추출
    """
    for r in range(min(10, df.shape[0])):
        for c in range(min(12, df.shape[1])):
            v = clean(df.iat[r, c])
            if not v:
                continue

            m = re.search(r"(\d{4})\.(\d{1,2})", v)  # 2025.9 / 2025.09 모두
            if m:
                return int(m.group(1)), int(m.group(2))

            m2 = re.search(r"(\d{4})\s*년\s*(\d{1,2})\s*월", v)
            if m2:
                return int(m2.group(1)), int(m2.group(2))

    raise ValueError("조회년월(예: 2025.10 또는 2025년 10월)을 찾지 못했습니다.")

def fetch_sido_map(cur) -> dict:
    cur.execute("SELECT sido_id, sido_name FROM dim_region_sido")
    rows = cur.fetchall()
    return {name: sid for (sid, name) in rows}

def chunked(lst, size=2000):
    for i in range(0, len(lst), size):
        yield lst[i:i+size]


# 05~08 시트에서 “시도별 차종 합계 행”만 추출
def get_sido_cols(df: pd.DataFrame, header_row=2, start_col=5):
    cols = []
    for c in range(start_col, df.shape[1]):
        s = clean(df.iat[header_row, c])
        if s in ALLOWED_SIDO:
            cols.append((c, s))
    return cols

def row_has_numbers(df: pd.DataFrame, r: int, sido_cols) -> bool:
    # 해당 행이 실제 데이터 행인지(헤더/라벨 행인지) 간단 판별
    # 시도 컬럼들 중 하나라도 숫자/숫자형이면 True
    for c, _ in sido_cols[:5]:  # 몇 개만 봐도 충분
        v = df.iat[r, c]
        sv = clean(v).replace(",", "")
        if sv.isdigit():
            return True
    return False

def find_row(df: pd.DataFrame, sido_cols, patterns, start_row=3, look_cols=5):
    """
    patterns: [regex, regex, ...]
    """
    regs = [re.compile(p) for p in patterns]
    for r in range(start_row, df.shape[0]):
        left = " ".join(clean(df.iat[r, c]) for c in range(min(look_cols, df.shape[1])))
        if not left:
            continue
        if any(rx.search(left) for rx in regs):
            if row_has_numbers(df, r, sido_cols):
                return r
    raise ValueError(f"행을 못 찾음: patterns={patterns}")

def extract_vehicle_stock_sheet(excel_path: str, sheet_name: str, usage_type: str):
    df = pd.read_excel(excel_path, sheet_name=sheet_name, header=None)
    year, month = parse_year_month(df)

    sido_cols = get_sido_cols(df, header_row=2, start_col=5)
    if not sido_cols:
        raise ValueError(f"[{sheet_name}] 시도 헤더 컬럼을 못 찾음 (row=2 col>=5에서 시도명이 있어야 함)")

    # 차종별 “합계 행” 찾기 (표기 흔들림 대비)
    # 합계(총계) / 승용 / 승합 / 화물 / 특수
    row_total = find_row(df, sido_cols, [r"총\s*계", r"합\s*계"], start_row=3)
    row_p = find_row(df, sido_cols, [r"승\s*용.*합\s*계", r"승용차.*합계"], start_row=3)
    row_v = find_row(df, sido_cols, [r"승\s*합.*합\s*계"], start_row=3)
    row_t = find_row(df, sido_cols, [r"화\s*물.*합\s*계", r"화물자동차.*합계"], start_row=3)
    row_s = find_row(df, sido_cols, [r"특\s*수.*합\s*계"], start_row=3)

    # fact_vehicle_stock.vehicle_kind 에 들어갈 값(네 ERD 기준)
    row_map = [
        ("합계", row_total),
        ("승용", row_p),
        ("승합", row_v),
        ("화물", row_t),
        ("특수", row_s),
    ]

    recs = []
    for vehicle_kind, rr in row_map:
        for cc, sido_name in sido_cols:
            recs.append({
                "year": year,
                "month": month,
                "origin_type": "전체",
                "sido_name": sido_name,
                "vehicle_kind": vehicle_kind,
                "usage_type": usage_type,
                "stock_count": to_int(df.iat[rr, cc]),
            })
    return year, month, recs


# 03.수입차_시군구 시트를 “시도 단위로 합산”하여 추출
def extract_import_vehicle_stock(excel_path: str):
    df = pd.read_excel(excel_path, sheet_name=IMPORT_SHEET, header=None)
    year, month = parse_year_month(df)

    data = df.iloc[4:, 0:7].copy()
    data.columns = ["sido","sigungu","승용","승합","화물","특수","합계"]

    data["sido"] = data["sido"].map(clean).replace("", pd.NA).ffill()
    data = data[data["sido"].isin(ALLOWED_SIDO)]

    for col in ["승용","승합","화물","특수","합계"]:
        data[col] = data[col].apply(to_int)

    # 시도 단위 집계
    agg = data.groupby("sido")[["승용","승합","화물","특수","합계"]].sum().reset_index()

    recs = []
    for _, r in agg.iterrows():
        sido = clean(r["sido"])
        for vk in ["승용","승합","화물","특수","합계"]:
            recs.append({
                "year": year,
                "month": month,
                "origin_type": "수입차",
                "sido_name": sido,
                "vehicle_kind": vk,
                "usage_type": "계",
                "stock_count": int(r[vk]),
            })
    return year, month, recs


# DB 적재 (월 단위 재적재 안전)
def load_fact_vehicle_stock(cur, excel_path: str):
    sido_map = fetch_sido_map(cur)

    all_recs = []
    ym = None

    # 05~08
    for sheet_name, usage_type in SHEETS_USAGE:
        y, m, recs = extract_vehicle_stock_sheet(excel_path, sheet_name, usage_type)
        all_recs.extend(recs)
        ym = (y, m)

    # 03 수입차(시도합산)
    y2, m2, recs2 = extract_import_vehicle_stock(excel_path)
    all_recs.extend(recs2)
    ym2 = (y2, m2)

    # year/month 불일치 방지
    if ym is None:
        raise ValueError("연/월을 추출하지 못했습니다.")
    if ym2 != ym:
        raise ValueError(f"시트 간 조회년월이 다릅니다: 05~08={ym}, 03={ym2}")

    year, month = ym
    print(f"[INFO] 대상월: {year}-{month:02d}")

    # 이 달만 재적재(다른 달 데이터는 보존)
    cur.execute("DELETE FROM fact_vehicle_stock WHERE year=%s AND month=%s", (year, month))

    sql = """
    INSERT INTO fact_vehicle_stock
    (year, month, origin_type, sido_id, vehicle_kind, usage_type, stock_count)
    VALUES (%s,%s,%s,%s,%s,%s,%s)
    ON DUPLICATE KEY UPDATE stock_count = VALUES(stock_count)
    """

    data = []
    for r in all_recs:
        sid = sido_map.get(r["sido_name"])
        if sid is None:
            # dim_region_sido에 없는 시도면 스킵
            continue
        data.append((
            r["year"], r["month"], r["origin_type"], sid,
            r["vehicle_kind"], r["usage_type"], r["stock_count"]
        ))

    # 대량 insert
    for part in chunked(data, 2000):
        cur.executemany(sql, part)

    print(f"[OK] fact_vehicle_stock 적재 완료: {len(data)} rows (월={year}-{month:02d})")

def main():
    conn = connect()
    try:
        conn.start_transaction()
        cur = conn.cursor()

        load_fact_vehicle_stock(cur, EXCEL_PATH)

        conn.commit()
        print("DONE ")
    except Exception as e:
        conn.rollback()
        print(" 실패:", e)
        raise
    finally:
        conn.close()

if __name__ == "__main__":
    main()
