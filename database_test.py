import streamlit as st
import pandas as pd
import mysql.connector

st.set_page_config(page_title="DB 적재 검증 차트", layout="wide")
st.title("✅ DB 적재 검증(비교용) 차트")

DB = {
    "host": "127.0.0.1",
    "port": 3307,
    "user": "admin",
    "password": "vmfhwprxm",
    "database": "SKN23",
    "connection_timeout": 5,
}

def get_conn():
    return mysql.connector.connect(**DB)

def query_df(sql: str, params: tuple = ()):
    conn = get_conn()
    try:
        cur = conn.cursor(dictionary=True)
        cur.execute(sql, params)
        rows = cur.fetchall()
        return pd.DataFrame(rows)
    finally:
        conn.close()

# -------------------------
# 0) 연결/기본 점검 (화면에 표시)
# -------------------------
try:
    ping = query_df("SELECT 1 AS ok")
    st.success("DB 연결 성공 ✅")
except Exception as e:
    st.error(f"DB 연결 실패 ❌: {e}")
    st.stop()

# -------------------------
# 1) (year, month) 실제 존재값 가져오기
# -------------------------
@st.cache_data(ttl=60)
def get_available_ym():
    df = query_df("""
        SELECT 'fact_vehicle_stock' AS tbl, year, month, COUNT(*) AS cnt
        FROM fact_vehicle_stock GROUP BY year, month
        UNION ALL
        SELECT 'fact_owner_demo_stock', year, month, COUNT(*)
        FROM fact_owner_demo_stock GROUP BY year, month
        UNION ALL
        SELECT 'fact_fuel_stock', year, month, COUNT(*)
        FROM fact_fuel_stock GROUP BY year, month
        UNION ALL
        SELECT 'fact_flow_count', year, month, COUNT(*)
        FROM fact_flow_count GROUP BY year, month
        ORDER BY year, month, tbl
    """)
    return df

ym_df = get_available_ym()
if ym_df.empty:
    st.warning("FACT 테이블에 데이터가 하나도 없습니다.")
    st.stop()

st.subheader("1) FACT 테이블별 월별 행 수(존재 여부 확인)")
st.dataframe(ym_df, use_container_width=True)

pivot_cnt = ym_df.pivot_table(index=["year","month"], columns="tbl", values="cnt", aggfunc="sum").fillna(0).reset_index()
pivot_cnt["ym"] = pivot_cnt["year"].astype(str) + "-" + pivot_cnt["month"].astype(int).astype(str).str.zfill(2)
pivot_cnt = pivot_cnt.set_index("ym")

st.bar_chart(pivot_cnt.drop(columns=["year","month"], errors="ignore"))

# (year, month) 선택은 "실제로 있는 값"에서만 고르게
ym_list = pivot_cnt.index.tolist()
default_ym = ym_list[-1]

st.subheader("2) 검증할 기준 월 선택(실제 존재하는 월만 표시)")
picked_ym = st.selectbox("year-month", ym_list, index=len(ym_list)-1)
picked_year = int(picked_ym.split("-")[0])
picked_month = int(picked_ym.split("-")[1])

# -------------------------
# 2) 성별/연령대/연료/용도 - “실데이터”로 차트
# -------------------------
st.divider()
c1, c2 = st.columns(2)

with c1:
    st.subheader("성별 보유대수 합계(전국)")

    df = query_df("""
        SELECT gender, SUM(stock_count) AS cnt
        FROM fact_owner_demo_stock
        WHERE year=%s AND month=%s
        GROUP BY gender
        ORDER BY cnt DESC
    """, (picked_year, picked_month))

    if df.empty:
        st.warning("선택한 월에 fact_owner_demo_stock 데이터 없음")
    else:
        st.bar_chart(df.set_index("gender")["cnt"])
        st.dataframe(df, use_container_width=True)

with c2:
    st.subheader("연령대 보유대수 합계(전국)")

    df = query_df("""
        SELECT g.age_group, SUM(f.stock_count) AS cnt
        FROM fact_owner_demo_stock f
        JOIN dim_age_group g ON g.age_group_id=f.age_group_id
        WHERE f.year=%s AND f.month=%s
        GROUP BY g.age_group, g.sort_order
        ORDER BY g.sort_order
    """, (picked_year, picked_month))

    if df.empty:
        st.warning("선택한 월에 연령대 데이터 없음")
    else:
        st.bar_chart(df.set_index("age_group")["cnt"])
        st.dataframe(df, use_container_width=True)

st.divider()

c3, c4 = st.columns(2)

with c3:
    st.subheader("연료 TOP 10(전국)")

    df = query_df("""
        SELECT fu.fuel_name, SUM(f.stock_count) AS cnt
        FROM fact_fuel_stock f
        JOIN dim_fuel fu ON fu.fuel_id=f.fuel_id
        WHERE f.year=%s AND f.month=%s
        GROUP BY fu.fuel_name
        ORDER BY cnt DESC
        LIMIT 10
    """, (picked_year, picked_month))

    if df.empty:
        st.warning("선택한 월에 fact_fuel_stock 데이터 없음")
    else:
        st.bar_chart(df.set_index("fuel_name")["cnt"])
        st.dataframe(df, use_container_width=True)

with c4:
    st.subheader("용도(관용/자가용/영업용) 합계(전국)")

    # ✅ 여기서 가장 많이 틀리는 부분: usage_type 값이 DB에 뭐로 들어갔는지 모르면 필터링하면 0됨
    # 그래서 우선 "해당 월에 존재하는 usage_type"을 그대로 집계 (필터 X)
    df = query_df("""
        SELECT usage_type, SUM(stock_count) AS cnt
        FROM fact_vehicle_stock
        WHERE year=%s AND month=%s
        AND vehicle_kind='합계'
        GROUP BY usage_type
        ORDER BY cnt DESC
    """, (picked_year, picked_month))

    if df.empty:
        st.warning("선택한 월에 fact_vehicle_stock 데이터 없음")
    else:
        st.bar_chart(df.set_index("usage_type")["cnt"])
        st.dataframe(df, use_container_width=True)

st.divider()

# -------------------------
# 3) “값이 이상하게 반복되는지” 빠른 체크 차트
# -------------------------
st.subheader("3) 월별 합계 비교(값이 월마다 달라지는지 빠른 체크)")

m1, m2 = st.columns(2)

with m1:
    st.caption("fact_fuel_stock: 월별 총합(stock_count 합)")
    df = query_df("""
        SELECT CONCAT(year,'-',LPAD(month,2,'0')) AS ym,
        SUM(stock_count) AS total
        FROM fact_fuel_stock
        GROUP BY year, month
        ORDER BY year, month
    """)
    if not df.empty:
        st.line_chart(df.set_index("ym")["total"])
        st.dataframe(df, use_container_width=True)

with m2:
    st.caption("fact_owner_demo_stock: 월별 총합(stock_count 합)")
    df = query_df("""
        SELECT CONCAT(year,'-',LPAD(month,2,'0')) AS ym,
        SUM(stock_count) AS total
        FROM fact_owner_demo_stock
        GROUP BY year, month
        ORDER BY year, month
    """)
    if not df.empty:
        st.line_chart(df.set_index("ym")["total"])
        st.dataframe(df, use_container_width=True)

st.caption("✅ 위 라인차트에서 월별 total이 완전히 동일하게 반복되면, ETL에서 year/month 파싱 또는 파일 경로/시트 읽기 로직이 잘못된 가능성이 큽니다.")
