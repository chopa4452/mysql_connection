import re
import pandas as pd
import mysql.connector

EXCEL_PATH = "2025년_10월_자동차_등록자료_통계.xlsx"

# 너가 쓰던 SSH 터널 기준 (127.0.0.1:3307)
DB = {
    "host": "127.0.0.1",
    "port": 3307,
    "user": "admin",
    "password": "vmfhwprxm",
    "database": "SKN23",
}

# 잘못 들어간 값 덮어쓰기만으로도 해결되긴 하지만,
# 과거 잘못 적재된 찌꺼기(케이스별)까지 없애려면 True 추천
DELETE_TARGET_MONTH_BEFORE_LOAD = True

ALLOWED_SIDO = {
    "강원","경기","경남","경북","광주","대구","대전","부산","서울","세종","울산","인천","전남","전북","제주","충남","충북"
}
ALLOWED_GENDER = {"남성", "여성", "기타"}

def connect():
    return mysql.connector.connect(**DB)

def clean(x) -> str:
    if x is None or pd.isna(x):
        return ""
    s = str(x)
    s = s.replace("_x000D_", "").replace("\r", "").replace("\n", " ").strip()
    s = re.sub(r"\s+", " ", s)
    return s

def parse_yyyymm_from_sheet(df: pd.DataFrame) -> tuple[int, int]:
    # 상단 8줄 정도에서 "조회년월: 2025.10" 같은 패턴 탐색
    for r in range(min(8, df.shape[0])):
        for c in range(min(12, df.shape[1])):
            v = clean(df.iat[r, c])
            m = re.search(r"(\d{4})\.(\d{1,2})", v)
            if m:
                return int(m.group(1)), int(m.group(2))
    # 못 찾으면 텍스트로 다시 탐색
    for r in range(min(10, df.shape[0])):
        row_txt = " ".join([clean(x) for x in df.iloc[r, :].tolist() if clean(x)])
        m = re.search(r"(\d{4})\.(\d{1,2})", row_txt)
        if m:
            return int(m.group(1)), int(m.group(2))
    raise ValueError("조회년월(YYYY.MM)을 시트 상단에서 찾지 못했어.")

def normalize_age(raw) -> str:
    s = clean(raw)
    s = s.replace("\u00A0", "").replace(" ", "")  # NBSP/공백 제거

    if not s:
        return ""

    # 법인/사업자
    if "법인" in s or "사업자" in s:
        return "법인 및 사업자"

    # 10대 이하
    if s.startswith("10대"):
        return "10대이하"

    # 20~80대, 90대 이상
    m = re.search(r"(\d{2})", s)
    if m:
        n = int(m.group(1))
        if n in (20,30,40,50,60,70,80):
            return f"{n}대"
        if n >= 90:
            return "90대이상"

    return s

def find_header_row_for_sido(df: pd.DataFrame) -> int:
    """
    04시트에서 '시도 헤더가 있는 줄' 자동 탐색:
    한 줄에 ALLOWED_SIDO가 일정 개수(>=10) 이상 나타나는 줄을 헤더로 판단
    """
    best_row = -1
    best_hit = 0
    for r in range(min(12, df.shape[0])):  # 상단만 훑으면 충분
        hits = 0
        for c in range(min(40, df.shape[1])):
            v = clean(df.iat[r, c])
            if v in ALLOWED_SIDO:
                hits += 1
        if hits > best_hit:
            best_hit = hits
            best_row = r

    if best_row == -1 or best_hit < 10:
        # fallback: 기존에 많이 쓰는 헤더 row=2
        return 2
    return best_row

def fetch_maps(cur):
    cur.execute("SELECT sido_id, sido_name FROM dim_region_sido")
    sido_map = {clean(name): sid for (sid, name) in cur.fetchall()}

    cur.execute("SELECT age_group_id, age_group FROM dim_age_group")
    raw_age_map = {clean(name): aid for (aid, name) in cur.fetchall()}

    # dim 값도 normalize해서 매칭 강화
    age_map = {}
    for name, aid in raw_age_map.items():
        age_map[normalize_age(name)] = aid

    return sido_map, age_map

def to_int(x) -> int:
    if x is None or pd.isna(x):
        return 0
    s = clean(x).replace(",", "")
    if s == "" or s == "-":
        return 0
    try:
        return int(float(s))
    except:
        return 0

def load_fact_owner_demo_stock():
    df = pd.read_excel(EXCEL_PATH, sheet_name="04.성별_연령별", header=None)
    year, month = parse_yyyymm_from_sheet(df)

    header_row = find_header_row_for_sido(df)

    # 보통 0=성별, 1=연령, 2=합계, 3~ 시도들
    # 헤더에서 시도 컬럼 인덱스 추출
    sido_cols = []
    for c in range(df.shape[1]):
        h = clean(df.iat[header_row, c])
        if h in ALLOWED_SIDO:
            sido_cols.append((c, h))

    if not sido_cols:
        raise ValueError("시도 헤더 컬럼을 찾지 못했어. 04시트 헤더 구조를 확인해야 함.")

    data_start = header_row + 1  # 헤더 다음 줄부터 데이터
    data = df.iloc[data_start:, :].copy()

    # 성별: raw에서 ffill 먼저 (NaN이 문자열로 바뀌기 전에!)
    gender_series = data.iloc[:, 0].ffill().apply(clean)
    age_series = data.iloc[:, 1].apply(normalize_age)

    # (성별/연령) + 시도별 값으로 레코드 생성
    recs = []
    unknown_ages = set()

    for i in range(len(data)):
        gender = clean(gender_series.iat[i])
        age = normalize_age(age_series.iat[i])

        # 불필요 행 제거
        if gender not in ALLOWED_GENDER:
            continue
        if not age:
            continue
        if age in {"총계","합계","계"}:
            continue

        # 기타는 보통 '법인 및 사업자'만 의미있음 (나머지면 제외)
        if gender == "기타" and age != "법인 및 사업자":
            continue

        for c, sido_name in sido_cols:
            v = data.iat[i, c]
            if pd.isna(v):
                continue
            recs.append((year, month, sido_name, gender, age, to_int(v)))

    return year, month, recs, sorted(unknown_ages)

def main():
    year, month, recs, _ = load_fact_owner_demo_stock()

    conn = connect()
    try:
        conn.start_transaction()
        cur = conn.cursor()

        sido_map, age_map = fetch_maps(cur)

        if DELETE_TARGET_MONTH_BEFORE_LOAD:
            cur.execute(
                "DELETE FROM fact_owner_demo_stock WHERE year=%s AND month=%s",
                (year, month)
            )

        sql = """
        INSERT INTO fact_owner_demo_stock
          (year, month, sido_id, gender, age_group_id, stock_count)
        VALUES (%s,%s,%s,%s,%s,%s)
        ON DUPLICATE KEY UPDATE stock_count = VALUES(stock_count)
        """

        data = []
        missing_age = set()
        missing_sido = set()

        for (y, m, sido_name, gender, age, cnt) in recs:
            sid = sido_map.get(sido_name)
            if sid is None:
                missing_sido.add(sido_name)
                continue

            aid = age_map.get(age)
            if aid is None:
                missing_age.add(age)
                continue

            data.append((y, m, sid, gender, aid, cnt))

        if missing_sido:
            print("⚠️ dim_region_sido에 없는 시도명(매핑 실패):", sorted(missing_sido))
        if missing_age:
            print("⚠️ dim_age_group에 없는 연령대(정규화 후):", sorted(missing_age))

        # 배치 insert
        cur.executemany(sql, data)
        conn.commit()

        print(f" owner_demo 재적재 완료: {len(data)} rows inserted/updated  (target={year}-{month:02d})")

    except Exception as e:
        conn.rollback()
        print(" 실패:", e)
        raise
    finally:
        conn.close()

if __name__ == "__main__":
    main()
