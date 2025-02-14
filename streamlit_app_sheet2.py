import streamlit as st
import pandas as pd
from datetime import datetime, timedelta

# 📌 현재 날짜 기준 전월 및 당월 계산
today = datetime.today()
current_month = today.strftime("%Y-%m")
previous_month = (today.replace(day=1) - timedelta(days=1)).strftime("%Y-%m")

# 📌 분석 대상 컬럼 설정
date_columns = ["입사일", "퇴사일"]
employee_types = ["정규직", "계약직", "파견직", "임원"]  # 가나다순 정렬

# 📌 웹 애플리케이션 UI
st.title("📊 섹스보지")
st.write("엑셀 파일을 업로드하면 자동으로 모든 시트를 분석합니다.")

# 📌 파일 업로드
uploaded_file = st.file_uploader("엑셀 파일을 선택하세요", type=["xlsx"])

if uploaded_file:
    try:
        # 📌 엑셀 파일의 모든 시트 불러오기
        sheets = pd.read_excel(uploaded_file, sheet_name=None, engine="openpyxl")

        for sheet_name, df in sheets.items():  # 각 시트를 반복하며 처리
            st.subheader(f"📄 시트 이름: {sheet_name}")

            # 📌 "도이치오토월드" 시트에서는 "장준호" 제외
            if sheet_name == "도이치오토월드" and "성명" in df.columns:
                df = df.loc[df["성명"] != "장준호"]

            # 📌 "DT네트웍스" 시트에서는 "권혁민" 제외
            if sheet_name == "DT네트웍스" and "성명" in df.columns:
                df = df.loc[df["성명"] != "권혁민"]

            # 📌 컬럼명 공백 제거
            df.columns = df.columns.str.strip()

            # 📌 날짜 변환 (문자열 → datetime 변환 후 .dt 사용)
            for col in date_columns:
                df[col] = pd.to_datetime(df[col], errors="coerce")  # NaN 값 처리

            # 📌 날짜를 'YYYY-MM' 형식으로 변환
            df["입사일"] = df["입사일"].dt.strftime("%Y-%m")
            df["퇴사일"] = df["퇴사일"].dt.strftime("%Y-%m")

            # 📌 1. "입사일"이 전월인 인원 수 & "퇴사일"이 전월인 인원 수
            new_hires_prev_month = df[df["입사일"] == previous_month].shape[0]
            resigned_prev_month = df[df["퇴사일"] == previous_month].shape[0]

            # 📌 2. "퇴사일"이 비어있거나 당월인 인원 수
            active_or_resigned_this_month = df[df["퇴사일"].isna() | (df["퇴사일"] == current_month)].shape[0]

            # 📌 3. "입사일"이 전월이며, "사원구분명"별 인원 수
            new_hires_by_type = df[df["입사일"] == previous_month]["사원구분명"].value_counts()

            # 📌 4. "퇴사일"이 비어있거나 당월이며, "사원구분명"별 인원 수
            active_or_resigned_this_month_by_type = df[df["퇴사일"].isna() | (df["퇴사일"] == current_month)]["사원구분명"].value_counts()

            # 📌 5. "퇴사일"이 전월이며, "사원구분명"별 인원 수
            resigned_by_type_prev_month = df[df["퇴사일"] == previous_month]["사원구분명"].value_counts()

            # 📌 결과 출력
            # 📊 "사원구분명"별 분석 결과 출력
            st.write("📌 1. **인원 수:**")
            for emp_type in employee_types:
                st.write(f"  - {emp_type}: {active_or_resigned_this_month_by_type.get(emp_type, 0)}명")

            st.write("📌 2. **전월 입사자 수:**")
            for emp_type in employee_types:
                st.write(f"  - {emp_type}: {new_hires_by_type.get(emp_type, 0)}명")

            st.write("📌 3. **전월 퇴사자 수:**")
            for emp_type in employee_types:
                st.write(f"  - {emp_type}: {resigned_by_type_prev_month.get(emp_type, 0)}명")

            st.markdown("---")  # 구분선 추가

    except Exception as e:
        st.error(f"❌ 파일 처리 중 오류 발생: {str(e)}")

