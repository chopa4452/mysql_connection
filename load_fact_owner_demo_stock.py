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

ALLOWED_GENDER = {"남성", "여성", "기타"}


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

def parse_year_month(df: pd.DataFrame) -> tuple[int, int]:
    for r in range(min(10, df.shape[0])):
        for c in range(min(12, df.shape[1])):
            v = clean(df.iat[r, c])
            if not v:
                continue
            m = re.search(r"(\d{4})\.(\d{1,2})", v)  # 2025.9 / 2025.09
            if m:
                return int(m.group(1)), int(m.group(2))
            m2 = re.search(r"(\d{4})\s*년\s*(\d{1,2})\s*월", v)
            if m2:
                return int(m2.group(1)), int(m2.group(2))
    raise ValueError("조회년월(YYYY.MM)을 찾지 못했습니다.")

def chunked(lst, size=2000):
    for i in range(0, len(lst), size):
        yield lst[i:i+size]


# 표준화(매칭 성공률 올리기)
def normalize_gender(s: str) -> str:
    s = clean(s)
    if s in ALLOWED_GENDER:
        return s
    # 혹시 영문/기타 표기 들어오면 여기서 매핑 가능
    return s

def normalize_age(s: str) -> str:
    """
    엑셀 달마다 미세하게 다른 표기를 dim_age_group과 최대한 맞추기
    목표 표준:
      10대이하, 20대, 30대, 40대, 50대, 60대, 70대, 80대, 90대이상, 법인 및 사업자
    """
    s = clean(s)
    if not s:
        return ""

    # '계' 제거용
    if s in {"계", "합계", "총계"}:
        return s

    # 공백 제거(예: '90대 이상' -> '90대이상', '10대 이하' -> '10대이하')
    s = s.replace(" ", "")

    # 흔한 변형 정리
    s = s.replace("10대이하", "10대이하")
    s = s.replace("10대이하", "10대이하")

    # '이상' 표기 통일
    s = s.replace("90대이상", "90대이상")
    s = s.replace("90대이상", "90대이상")

    # '90대 이상' 같은 걸 위에서 공백 제거로 이미 '90대이상' 됨
    # '10대 이하'도 '10대이하'

    # '법인및사업자' 같은 경우 대비
    if "법인" in s and "사업자" in s:
        return "법인 및 사업자"

    # '20대'~'80대' 정규화
    m = re.fullmatch(r"(\d{2})대", s)
    if m:
        return f"{m.group(1)}대"

    # '10대이하' / '90대이상'은 그대로
    if s in {"10대이하", "90대이상"}:
        return s

    return s


# DIM 맵
def fetch_sido_map(cur):
    cur.execute("SELECT sido_id, sido_name FROM dim_region_sido")
    return {name: sid for (sid, name) in cur.fetchall()}

def fetch_age_map(cur):
    cur.execute("SELECT age_group_id, age_group FROM dim_age_group")
    return {name: aid for (aid, name) in cur.fetchall()}


# 추출(04.성별_연령별)
def extract_owner_demo(excel_path: str):
    df = pd.read_excel(excel_path, sheet_name="04.성별_연령별", header=None)
    year, month = parse_year_month(df)

    # 시도 헤더: row 2, col 3~ 에 서울/부산... (0=성별,1=연령,2=합계,3~=시도)
    header = [clean(x) for x in df.iloc[2, :].tolist()]
    sido_cols = [(i, header[i]) for i in range(3, len(header)) if header[i] in ALLOWED_SIDO]
    if not sido_cols:
        raise ValueError("[04.성별_연령별] 시도 헤더(col>=3)에서 시도명을 못 찾음")

    data = df.iloc[3:, :].copy()

    # 핵심: 성별은 ffill 해야 함 (빈칸 많음)
    data["gender"] = data[0].map(clean).replace("", pd.NA).ffill()
    data["age_group"] = data[1].map(clean)

    recs = []
    for _, r in data.iterrows():
        g = normalize_gender(r["gender"])
        a_raw = r["age_group"]
        a = normalize_age(a_raw)

        if g not in ALLOWED_GENDER:
            continue
        if not a or a in {"계", "합계", "총계"}:
            continue

        for col_idx, sido_name in sido_cols:
            recs.append({
                "year": year,
                "month": month,
                "sido_name": sido_name,
                "gender": g,
                "age_group": a,
                "stock_count": to_int(r[col_idx]),
            })

    return year, month, recs


# 적재 (월별 누적 + 해당 월 재실행 안전)
def load_fact_owner_demo_stock(cur, excel_path: str):
    sido_map = fetch_sido_map(cur)
    age_map = fetch_age_map(cur)

    year, month, recs = extract_owner_demo(excel_path)
    print(f"[INFO] 대상월: {year}-{month:02d}, 추출 rows={len(recs)}")

    # 이 달만 재적재 (다른 달은 보존)
    cur.execute("DELETE FROM fact_owner_demo_stock WHERE year=%s AND month=%s", (year, month))

    sql = """
    INSERT INTO fact_owner_demo_stock
    (year, month, sido_id, gender, age_group_id, stock_count)
    VALUES (%s,%s,%s,%s,%s,%s)
    ON DUPLICATE KEY UPDATE stock_count = VALUES(stock_count)
    """

    data = []
    missing_age = set()
    missing_sido = set()

    for r in recs:
        sid = sido_map.get(r["sido_name"])
        aid = age_map.get(r["age_group"])
        if sid is None:
            missing_sido.add(r["sido_name"])
            continue
        if aid is None:
            missing_age.add(r["age_group"])
            continue
        data.append((r["year"], r["month"], sid, r["gender"], aid, r["stock_count"]))

    # 디버깅용: 매칭 실패 로그
    if missing_sido:
        print("[WARN] dim_region_sido 매칭 실패:", sorted(missing_sido))
    if missing_age:
        print("[WARN] dim_age_group 매칭 실패:", sorted(missing_age))

    for part in chunked(data, 2000):
        cur.executemany(sql, part)

    print(f"[OK] fact_owner_demo_stock 적재 완료: {len(data)} rows (월={year}-{month:02d})")

def main():
    conn = connect()
    try:
        conn.start_transaction()
        cur = conn.cursor()

        load_fact_owner_demo_stock(cur, EXCEL_PATH)

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
