import os
import math
from datetime import datetime, timedelta
import pandas as pd
import streamlit as st

# ──────────────────────────────────────────────
# 0. 페이지 기본 설정
# ──────────────────────────────────────────────
st.set_page_config(page_title="원화 담당자 추론 시스템", page_icon="🔍", layout="wide")

st.markdown("""
    <style>
    section[data-testid="stSidebar"] { width: 250px !important; }
    </style>
""", unsafe_allow_html=True)

# ──────────────────────────────────────────────
# 1. 데이터 로딩
# ──────────────────────────────────────────────
DATA_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "DB_PIC.xlsx")

REQUIRED_COLS_MAP = {
    "구분": ["구분", "유형", "Type"],
    "계좌번호": ["계좌번호", "계좌", "Account"],
    "내역": ["내역", "적요", "비고", "거래내역", "Description"],
    "금액": ["입금액", "금액", "거래금액", "Amount"],
    "담당자": ["담당자", "처리자", "PIC", "Person"],
}

@st.cache_data(show_spinner="☁️ 데이터 로딩 중...")
def load_data():
    try:
        df = pd.read_excel(DATA_FILE, engine="openpyxl")
    except:
        df = pd.read_excel(DATA_FILE)
    
    col_map = {}
    for logical_name, candidates in REQUIRED_COLS_MAP.items():
        found = next((c for c in candidates if c in df.columns), None)
        if found is None:
            st.error(f"❌ '{logical_name}' 컬럼을 찾을 수 없습니다."); st.stop()
        col_map[logical_name] = found

    date_col = next((c for c in ["거래일자", "날짜", "Date"] if c in df.columns), None)
    col_map["날짜"] = date_col
    if date_col:
        df[date_col] = pd.to_datetime(df[date_col], errors="coerce")

    for key in ["구분", "계좌번호", "내역", "담당자"]:
        df[col_map[key]] = df[col_map[key]].astype(str).fillna("").str.strip()
    df[col_map["금액"]] = pd.to_numeric(df[col_map["금액"]], errors="coerce").fillna(0).astype(int)
    
    return df, col_map

# ──────────────────────────────────────────────
# 2. 공통 테이블 설정 (UI 개선의 핵심)
# ──────────────────────────────────────────────
# 💡 데이터프레임의 실제 열 이름과 일치하도록 정의함
COLUMN_CONFIG = {
    "구분": st.column_config.TextColumn("구분", width="small"),
    "계좌번호": st.column_config.TextColumn("계좌번호", width="medium"),
    "거래일자": st.column_config.TextColumn("거래일자", width="small"),
    "금액": st.column_config.NumberColumn("금액", format="%d", width="small"), # 콤마 자동 생성
    "과거내역": st.column_config.TextColumn("과거내역", width="large"),
    "담당자": st.column_config.TextColumn("담당자", width="small"),
    "점수": st.column_config.ProgressColumn("확신도", format="%.0f%%", min_value=0, max_value=100, width="small"), # 게이지 바 UI
    "매칭근거": st.column_config.TextColumn("매칭근거", width="large"),
}

# ──────────────────────────────────────────────
# 3. 추론 및 통계 알고리즘
# ──────────────────────────────────────────────
def infer_person(df, col_map, input_category, input_account, input_desc):
    today = datetime.now()
    c_cat, c_acc, c_desc, c_person, c_date, c_amount = (
        col_map["구분"], col_map["계좌번호"], col_map["내역"], 
        col_map["담당자"], col_map.get("날짜"), col_map["금액"]
    )
    
    input_tokens = set(input_desc.strip().lower().split())
    if not input_tokens: return None

    records = []
    for _, row in df.iterrows():
        score, reasons = 0.0, []
        db_desc = str(row[c_desc]).lower()
        
        matched_cnt = sum(1 for t in input_tokens if t in db_desc)
        kw_sim = matched_cnt / len(input_tokens) if len(input_tokens) > 0 else 0
        
        if input_desc.strip() == db_desc.strip():
            kw_sim = 1.0; reasons.append("내역 완전 일치🎯")
        
        if kw_sim > 0.3:
            score += (kw_sim * 65)
            if "내역 완전 일치" not in str(reasons): reasons.append(f"내역 유사({kw_sim:.0%})")
        else: continue

        if input_account.strip() and str(row[c_acc]).strip() == input_account.strip():
            score += 15; reasons.append("계좌 일치")
        if input_category and str(row[c_cat]).strip() == input_category:
            score += 5; reasons.append("구분 일치")

        raw_date = row[c_date]
        formatted_date = "-"
        if c_date and pd.notna(raw_date):
            delta_days = max(0, (today - raw_date).days)
            rec_score = math.exp(-delta_days / 365)
            score += (rec_score * 15)
            formatted_date = raw_date.strftime('%Y-%m-%d')

        records.append({
            "구분": str(row[c_cat]), "계좌번호": str(row[c_acc]), "거래일자": formatted_date,
            "금액": row[c_amount], "과거내역": db_desc, "담당자": str(row[c_person]),
            "점수": min(100.0, round(score, 1)), "매칭근거": ", ".join(reasons)
        })

    if not records: return None
    return pd.DataFrame(records).sort_values(["점수", "거래일자"], ascending=False)

def generate_report(df, col_map):
    c_cat, c_desc, c_person, c_date = col_map["구분"], col_map["내역"], col_map["담당자"], col_map.get("날짜")
    
    one_year_ago = datetime.now() - timedelta(days=365)
    if c_date:
        report_df = df[df[c_date] >= one_year_ago].copy()
    else:
        report_df = df.copy()

    report_df["내역유형"] = report_df[c_desc].apply(lambda x: " ".join(str(x).split()[:2]))
    summary = report_df.groupby([c_cat, "내역유형", c_person]).size().reset_index(name="처리건수")
    
    idx = summary.groupby([c_cat, "내역유형"])["처리건수"].idxmax()
    final_report = summary.loc[idx].sort_values([c_cat, "처리건수"], ascending=[True, False])
    
    final_report.columns = ["구분(공장)", "주요내역패턴", "주요담당자", "최근1년처리건수"]
    return final_report

# ──────────────────────────────────────────────
# 4. 메인 UI
# ──────────────────────────────────────────────
def main():
    st.title("🔍 원화 담당자 업무 지원 시스템")
    df, col_map = load_data()

    with st.sidebar.form("input_form"):
        st.header("📝 검색/추론 입력")
        unique_cats = sorted([c for c in df[col_map["구분"]].unique() if str(c).strip()])
        input_category = st.selectbox("구분(공장)", options=["선택 안 함"] + unique_cats)
        cat_value = "" if input_category == "선택 안 함" else input_category
        input_account = st.text_input("계좌번호")
        input_desc = st.text_input("내역 (추론 시 필수)")
        input_pic = st.text_input("담당자 (조회 시)")
        submitted = st.form_submit_button("🔎 실행", use_container_width=True)

    tab1, tab2 = st.tabs(["🎯 실시간 추론 및 조회", "📊 주요 담당자 리포트"])

    with tab1:
        if submitted:
            if input_pic.strip():
                st.subheader(f"👤 [{input_pic.strip()}] 담당자 내역 조회")
                f_df = df[df[col_map["담당자"]].str.contains(input_pic.strip(), na=False)].copy()
                
                if cat_value: f_df = f_df[f_df[col_map["구분"]] == cat_value]
                if input_account.strip(): f_df = f_df[f_df[col_map["계좌번호"]].str.contains(input_account.strip())]
                if input_desc.strip(): f_df = f_df[f_df[col_map["내역"]].str.contains(input_desc.strip(), na=False, case=False)]
                
                if f_df.empty:
                    st.warning("검색 결과가 없습니다.")
                else:
                    # 💡 UI 고정을 위해 컬럼명 통일 (rename)
                    f_df = f_df.rename(columns={
                        col_map["구분"]: "구분", col_map["계좌번호"]: "계좌번호", 
                        col_map["내역"]: "과거내역", col_map["담당자"]: "담당자", col_map["금액"]: "금액"
                    })
                    f_df["거래일자"] = f_df[col_map["날짜"]].dt.strftime('%Y-%m-%d') if col_map["날짜"] else "-"
                    
                    view_cols = ["구분", "계좌번호", "거래일자", "금액", "과거내역", "담당자"]
                    st.dataframe(f_df[view_cols], use_container_width=True, hide_index=True, column_config=COLUMN_CONFIG)

            elif input_desc.strip():
                res_df = infer_person(df, col_map, cat_value, input_account, input_desc)
                if res_df is not None:
                    top = res_df.iloc[0]
                    st.subheader("🏆 추론 결과")
                    m1, m2, m3 = st.columns(3)
                    m1.metric("👤 추천 담당자", top["담당자"])
                    m2.metric("🎯 확신도 점수", f"{top['점수']}%")
                    m3.metric("📊 신뢰도", "높음" if top["점수"] >= 70 else "보통" if top["점수"] >= 40 else "낮음")
                    
                    st.divider()
                    st.subheader("📋 참고할 유사 사례 (Top 10)")
                    temp_df = res_df.head(10).copy()
                    
                    # 💡 컬럼 순서 및 구성 고정
                    infer_cols = ["구분", "계좌번호", "거래일자", "금액", "과거내역", "담당자", "점수", "매칭근거"]
                    st.dataframe(temp_df[infer_cols], use_container_width=True, hide_index=True, column_config=COLUMN_CONFIG)
                else:
                    st.error("유사한 내역을 찾을 수 없습니다.")
            else:
                st.warning("내역이나 담당자 중 하나는 입력해야 합니다.")
        else:
            st.info("왼쪽 사이드바에서 검색 조건을 입력하세요.")

    with tab2:
        st.subheader("📈 구분(공장)별 주요 업무 분장 현황")
        st.write("과거 1년 동안의 거래를 분석하여 요약한 리포트입니다.")
        
        report_data = generate_report(df, col_map)
        factories = ["전체"] + list(report_data["구분(공장)"].unique())
        selected_factory = st.selectbox("구분(공장) 필터", factories)
        
        if selected_factory == "전체":
            filtered_report = report_data
        else:
            filtered_report = report_data[report_data["구분(공장)"] == selected_factory]
        
        # 리포트용 별도 설정
        st.dataframe(filtered_report, use_container_width=True, hide_index=True, column_config={
            "최근1년처리건수": st.column_config.NumberColumn(format="%d")
        })
        
        csv = filtered_report.to_csv(index=False).encode('utf-8-sig')
        st.download_button(
            label="📥 리포트 엑셀(CSV) 다운로드",
            data=csv,
            file_name=f"담당자_리포트_{datetime.now().strftime('%Y%m%d')}.csv",
            mime="text/csv",
        )

if __name__ == "__main__":
    main()
