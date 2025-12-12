import re
import pandas as pd
import mysql.connector

EXCEL_PATH = "2025년_10월_자동차_등록자료_통계.xlsx"
DB = dict(host="localhost", user="root", password="root", database="python_test")

SKIP_TOTAL = {"총계", "합계", "계", "Total", "TOTAL"}


# 공통 유틸

def get_conn():
    return mysql.connector.connect(**DB)

def clean(x) -> str:
    if pd.isna(x):   # NaN/None 전부 빈문자 처리
        return ""
    s = str(x)
    s = s.replace("_x000D_", "").replace("\r", "").replace("\n", " ").strip()
    s = re.sub(r"\s+", " ", s)
    return s

def parse_yyyymm_from_sheet(df):
    # "조회년월: 2025.10" 형태에서 year=2025, month=10 추출
    text = ""
    for i in range(min(6, len(df))):
        row = " ".join([str(x) for x in df.iloc[i].tolist() if pd.notna(x)])
        if "조회년월" in row:
            text = row
            break
    m = re.search(r"(\d{4})\.(\d{1,2})", text)
    if not m:
        raise ValueError("조회년월을 못 찾았어. (시트 상단 '조회년월: 2025.10' 형태 확인)")
    return int(m.group(1)), int(m.group(2))


# DIM 적재

def upsert_dim_region_sido(cur, sido_names):
    cleaned = [clean_text(x) for x in sido_names]
    cleaned = [x for x in cleaned if x and x not in SKIP_TOTAL]  # ✅ None 제거 + 총계 제거

    for name in sorted(set(cleaned)):
        cur.execute(
            "INSERT IGNORE INTO dim_region_sido(sido_name, use_yn) VALUES (%s,'Y')",
            (name,)
        )

def upsert_dim_region_sigungu(cur, pairs, sido_id_map):
    cleaned_pairs = []
    for sido_name, sigungu_name in pairs:
        sido_name = clean_text(sido_name)
        sigungu_name = clean_text(sigungu_name)
        if not sido_name or not sigungu_name:
            continue
        if sigungu_name in SKIP_TOTAL:
            continue
        cleaned_pairs.append((sido_name, sigungu_name))

    for sido_name, sigungu_name in sorted(set(cleaned_pairs)):
        sido_id = sido_id_map.get(sido_name)
        if not sido_id:
            continue
        cur.execute(
            "INSERT IGNORE INTO dim_region_sigungu(sido_id, sigungu_name, use_yn) VALUES (%s,%s,'Y')",
            (sido_id, sigungu_name)
        )


def upsert_dim_age_group(cur, age_groups):
    def order_key(a):
        if a == "10대이하": return 10
        if a == "90대이상": return 90
        if a == "계": return 99
        m = re.search(r"(\d+)대", a)
        return int(m.group(1)) if m else 98

    for a in sorted(set([clean_text(x) for x in age_groups])):
        if not a:
            continue
        cur.execute(
            "INSERT IGNORE INTO dim_age_group(age_group, sort_order) VALUES (%s,%s)",
            (a, order_key(a))
        )

def upsert_dim_fuel(cur, fuel_names):
    eco_set = {"전기", "수소", "하이브리드", "하이브리드(HEV)", "플러그인하이브리드", "PHEV"}
    for f in sorted(set([clean_text(x) for x in fuel_names])):
        if not f:
            continue
        is_eco = 'Y' if f in eco_set else 'N'
        cur.execute(
            "INSERT IGNORE INTO dim_fuel(fuel_name, is_eco) VALUES (%s,%s)",
            (f, is_eco)
        )


# FACT 적재

def load_fact_vehicle_stock_01(cur, year, month, df, sido_id_map):
    # 01.통계표 : (시도별) + (차종) + (용도)
    n_cols = df.shape[1]

    header_vehicle = df.iloc[2, :n_cols].tolist()
    header_usage = df.iloc[3, :n_cols].tolist()
    max_col = min(n_cols, len(header_vehicle), len(header_usage))

    data = df.iloc[5:, :n_cols].copy()
    data.columns = list(range(n_cols))
    data["sido"] = data.iloc[:, 0].astype(str).str.strip()
    data = data[~data["sido"].isin(["nan", "None", "총계", "합계", "계"])]

    rows = []
    for col_idx in range(2, max_col):
        vk = clean_text(header_vehicle[col_idx])
        ut = clean_text(header_usage[col_idx])
        if not vk or not ut:
            continue
        if vk not in ["승용", "승합", "화물", "특수"]:
            continue

        for _, r in data.iterrows():
            sido_name = clean_text(r["sido"])
            if not sido_name:
                continue
            sido_id = sido_id_map.get(sido_name)
            if not sido_id:
                continue

            val = r.iloc[col_idx]
            if pd.isna(val):
                continue

            rows.append((year, month, "전체", sido_id, None, vk, ut, int(val)))

    if rows:
        cur.executemany(
            """
            INSERT INTO fact_vehicle_stock
            (year, month, origin_type, sido_id, sigungu_id, vehicle_kind, usage_type, stock_count)
            VALUES (%s,%s,%s,%s,%s,%s,%s,%s)
            """,
            rows
        )

def load_fact_vehicle_stock_02(cur, year, month, df, sido_id_map, sigungu_id_map):
    # 02.통계표_시군구 : (시도/시군구) + (차종) + (용도)
    n_cols = df.shape[1]

    header_vehicle = df.iloc[2, :n_cols].tolist()
    header_usage = df.iloc[3, :n_cols].tolist()
    max_col = min(n_cols, len(header_vehicle), len(header_usage))

    data = df.iloc[4:, :n_cols].copy()
    data.columns = list(range(n_cols))

    data["sido"] = data.iloc[:, 0].ffill().astype(str).str.strip()
    data["sigungu"] = data.iloc[:, 1].astype(str).str.strip()

    rows = []
    for col_idx in range(2, max_col):
        vk = clean_text(header_vehicle[col_idx])
        ut = clean_text(header_usage[col_idx])
        if not vk or not ut:
            continue
        if vk not in ["승용", "승합", "화물", "특수"]:
            continue

        for _, r in data.iterrows():
            sido_name = clean_text(r["sido"])
            sigungu_name = clean_text(r["sigungu"])
            if not sido_name or not sigungu_name:
                continue
            if sigungu_name in SKIP_TOTAL:
                continue

            sido_id = sido_id_map.get(sido_name)
            sigungu_id = sigungu_id_map.get((sido_name, sigungu_name))
            if not sido_id or not sigungu_id:
                continue

            val = r.iloc[col_idx]
            if pd.isna(val):
                continue

            rows.append((year, month, "전체", sido_id, sigungu_id, vk, ut, int(val)))

    if rows:
        cur.executemany(
            """
            INSERT INTO fact_vehicle_stock
            (year, month, origin_type, sido_id, sigungu_id, vehicle_kind, usage_type, stock_count)
            VALUES (%s,%s,%s,%s,%s,%s,%s,%s)
            """,
            rows
        )

def load_fact_vehicle_stock_03_import(cur, year, month, df, sido_id_map, sigungu_id_map):
    # 03.수입차_시군구 : (시도/시군구) + (승용/승합/화물/특수/합계)
    n_cols = df.shape[1]
    header = df.iloc[2, :n_cols].tolist()

    data = df.iloc[4:, :n_cols].copy()
    data.columns = list(range(n_cols))
    data["sido"] = data.iloc[:, 0].ffill().astype(str).str.strip()
    data["sigungu"] = data.iloc[:, 1].astype(str).str.strip()

    max_col = min(7, n_cols)  # 원래 2~6까지 쓰는데 컬럼이 더 적을 수 있으니 방어

    rows = []
    for col_idx in range(2, max_col):
        vk = clean_text(header[col_idx])
        if vk not in ["승용", "승합", "화물", "특수", "합계"]:
            continue

        for _, r in data.iterrows():
            sido_name = clean_text(r["sido"])
            sigungu_name = clean_text(r["sigungu"])
            if not sido_name or not sigungu_name:
                continue
            if sigungu_name in SKIP_TOTAL:
                continue

            sido_id = sido_id_map.get(sido_name)
            sigungu_id = sigungu_id_map.get((sido_name, sigungu_name))
            if not sido_id or not sigungu_id:
                continue

            val = r.iloc[col_idx]
            if pd.isna(val):
                continue

            rows.append((year, month, "수입차", sido_id, sigungu_id, vk, "계", int(val)))

    if rows:
        cur.executemany(
            """
            INSERT INTO fact_vehicle_stock
            (year, month, origin_type, sido_id, sigungu_id, vehicle_kind, usage_type, stock_count)
            VALUES (%s,%s,%s,%s,%s,%s,%s,%s)
            """,
            rows
        )

def load_fact_owner_demo_stock_04(cur, year, month, df, sido_id_map, age_id_map):
    # 04.성별_연령별 : 행(성별/연령) + 열(시도)
    n_cols = df.shape[1]

    header_row = 2    # 시도 헤더가 있는 줄
    data_start = 3    # 실제 데이터 시작 줄(헤더 아래)

    data = df.iloc[data_start:, :n_cols].copy()
    data.columns = list(range(n_cols))

    # 성별/연령 열
    data["gender"] = data.iloc[:, 0].ffill().astype(str).str.strip()
    data["age_group"] = data.iloc[:, 1].astype(str).str.strip()

    # 시도 컬럼: 3번째 컬럼부터(보통 0=성별, 1=연령, 2=합계, 3~ = 시도)
    sido_cols = []
    for idx in range(3, n_cols):
        name = clean_text(df.iloc[header_row, idx])
        if not name or name in SKIP_TOTAL:
            continue
        sido_cols.append((idx, name))

    rows = []
    for _, r in data.iterrows():
        gender = clean_text(r["gender"])
        age = clean_text(r["age_group"])
        if not gender or not age:
            continue

        age_id = age_id_map.get(age)
        if not age_id:
            continue

        for col_idx, sido_name in sido_cols:
            sido_id = sido_id_map.get(sido_name)
            if not sido_id:
                continue

            val = r.iloc[col_idx]
            if pd.isna(val):
                continue

            rows.append((year, month, sido_id, gender, age_id, int(val)))

    if rows:
        cur.executemany(
            """
            INSERT INTO fact_owner_demo_stock
            (year, month, sido_id, gender, age_group_id, stock_count)
            VALUES (%s,%s,%s,%s,%s,%s)
            """,
            rows
        )

def load_fact_fuel_stock_10(cur, year, month, df, sido_id_map, fuel_id_map):
    # 10.연료별_등록현황 : 행(연료/차종/사업구분) + 열(시도)
    n_cols = df.shape[1]

    header_row = 2
    data_start = 4

    data = df.iloc[data_start:, :n_cols].copy()
    data.columns = list(range(n_cols))

    data["fuel"] = data.iloc[:, 0].ffill()
    data["vehicle_kind"] = data.iloc[:, 1]
    data["business_type"] = data.iloc[:, 2]

    data["fuel"] = data["fuel"].apply(clean_text)
    data["vehicle_kind"] = data["vehicle_kind"].apply(clean_text)
    data["business_type"] = data["business_type"].apply(clean_text)

    sido_cols = []
    for idx in range(3, n_cols):
        name = clean_text(df.iloc[header_row, idx])
        if not name or name in SKIP_TOTAL:
            continue
        sido_cols.append((idx, name))

    rows = []
    for _, r in data.iterrows():
        fuel = r["fuel"]
        vk = r["vehicle_kind"]
        bt = r["business_type"]

        if not (fuel and vk and bt):
            continue
        if vk not in ["승용", "승합", "화물", "특수"]:
            continue

        fuel_id = fuel_id_map.get(fuel)
        if not fuel_id:
            continue

        for col_idx, sido_name in sido_cols:
            sido_id = sido_id_map.get(sido_name)
            if not sido_id:
                continue

            val = r.iloc[col_idx]
            if pd.isna(val):
                continue

            rows.append((year, month, sido_id, fuel_id, vk, bt, int(val)))

    if rows:
        cur.executemany(
            """
            INSERT INTO fact_fuel_stock
            (year, month, sido_id, fuel_id, vehicle_kind, business_type, stock_count)
            VALUES (%s,%s,%s,%s,%s,%s,%s)
            """,
            rows
        )


# 메인 실행

def main():
    # 시트 읽기
    raw_01 = pd.read_excel(EXCEL_PATH, sheet_name="01.통계표", header=None)
    raw_02 = pd.read_excel(EXCEL_PATH, sheet_name="02.통계표_시군구", header=None)
    raw_03 = pd.read_excel(EXCEL_PATH, sheet_name="03.수입차_시군구", header=None)
    raw_04 = pd.read_excel(EXCEL_PATH, sheet_name="04.성별_연령별", header=None)
    raw_10 = pd.read_excel(EXCEL_PATH, sheet_name="10.연료별_등록현황", header=None)

    year, month = parse_yyyymm_from_sheet(raw_01)
    print(f"[INFO] 조회년월: {year}-{month:02d}")

    conn = get_conn()
    conn.autocommit = False
    cur = conn.cursor()

    try:
        # (테스트) FACT 비우기
        for t in ["fact_fuel_stock", "fact_owner_demo_stock", "fact_vehicle_stock"]:
            cur.execute(f"TRUNCATE TABLE {t}")


        # DIM 채우기

        sido_names = set()

        # 01/02/03 시도명(첫 컬럼)
        sido_names |= set([clean_text(x) for x in raw_01.iloc[5:, 0].dropna().tolist()])
        sido_names |= set([clean_text(x) for x in raw_02.iloc[4:, 0].dropna().tolist()])
        sido_names |= set([clean_text(x) for x in raw_03.iloc[4:, 0].dropna().tolist()])

        # 04/10 헤더의 시도명(2행)
        sido_names |= set([clean_text(x) for x in raw_04.iloc[2, :].tolist()])
        sido_names |= set([clean_text(x) for x in raw_10.iloc[2, :].tolist()])

        upsert_dim_region_sido(cur, sido_names)
        conn.commit()

        cur.execute("SELECT sido_name, sido_id FROM dim_region_sido")
        sido_id_map = {name: id_ for (name, id_) in cur.fetchall()}

        # 시군구(02/03: 시도 ffill + 시군구)
        pairs = []
        tmp02 = raw_02.iloc[4:, [0, 1]].copy()
        tmp02.iloc[:, 0] = tmp02.iloc[:, 0].ffill()
        for a, b in tmp02.dropna().values.tolist():
            pairs.append((a, b))

        tmp03 = raw_03.iloc[4:, [0, 1]].copy()
        tmp03.iloc[:, 0] = tmp03.iloc[:, 0].ffill()
        for a, b in tmp03.dropna().values.tolist():
            pairs.append((a, b))

        upsert_dim_region_sigungu(cur, pairs, sido_id_map)
        conn.commit()

        cur.execute("""
            SELECT s.sido_name, g.sigungu_name, g.sigungu_id
            FROM dim_region_sigungu g
            JOIN dim_region_sido s ON s.sido_id = g.sido_id
        """)
        sigungu_id_map = {(sido, sigungu): gid for (sido, sigungu, gid) in cur.fetchall()}

        # 연령대
        age_groups = set([clean_text(x) for x in raw_04.iloc[3:, 1].dropna().tolist()])
        upsert_dim_age_group(cur, age_groups)
        conn.commit()

        cur.execute("SELECT age_group, age_group_id FROM dim_age_group")
        age_id_map = {a: i for (a, i) in cur.fetchall()}

        # 연료
        fuels = set([clean_text(x) for x in raw_10.iloc[4:, 0].dropna().tolist()])
        upsert_dim_fuel(cur, fuels)
        conn.commit()

        cur.execute("SELECT fuel_name, fuel_id FROM dim_fuel")
        fuel_id_map = {f: i for (f, i) in cur.fetchall()}


        # FACT 적재

        load_fact_vehicle_stock_01(cur, year, month, raw_01, sido_id_map)
        load_fact_vehicle_stock_02(cur, year, month, raw_02, sido_id_map, sigungu_id_map)
        load_fact_vehicle_stock_03_import(cur, year, month, raw_03, sido_id_map, sigungu_id_map)
        load_fact_owner_demo_stock_04(cur, year, month, raw_04, sido_id_map, age_id_map)
        load_fact_fuel_stock_10(cur, year, month, raw_10, sido_id_map, fuel_id_map)

        conn.commit()
        print("[OK] 적재 완료!")

    except Exception as e:
        conn.rollback()
        print("[ERROR] 롤백했어. 에러:", e)
        raise
    finally:
        cur.close()
        conn.close()

if __name__ == "__main__":
    main()
