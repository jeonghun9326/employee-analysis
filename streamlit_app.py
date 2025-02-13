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
st.title("📊 인원 분석 시스템")
st.write("엑셀 파일을 업로드하면 자동으로 데이터를 분석합니다.")

# 📌 파일 업로드
uploaded_file = st.file_uploader("엑셀 파일을 선택하세요", type=["xlsx"])

if uploaded_file:
    try:
        # 📌 엑셀 파일 로드
        df = pd.read_excel(uploaded_file, parse_dates=date_columns, engine="openpyxl")

        # 📌 데이터 정리
        df.columns = df.columns.str.strip()  # 컬럼명 공백 제거
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
        st.subheader("📊 분석 결과")
        st.write(f"1. **전월({previous_month}) 입사자 수**: {new_hires_prev_month}")
        st.write(f"2. **전월({previous_month}) 퇴사자 수**: {resigned_prev_month}")
        st.write(f"3. **퇴사일이 비어있거나 당월({current_month})인 인원 수**: {active_or_resigned_this_month - 1}")

        # 📊 "사원구분명"별 분석 결과 출력
        st.subheader("📌 사원구분별 분석")

        st.write("📌 4. **인원 수:**")
        for emp_type in employee_types:
            st.write(f"  - {emp_type}: {active_or_resigned_this_month_by_type.get(emp_type, 0)}명")

        st.write("📌 5. **전월 입사자 수:**")
        for emp_type in employee_types:
            st.write(f"  - {emp_type}: {new_hires_by_type.get(emp_type, 0)}명")

        st.write("📌 6. **전월 퇴사자 수:**")
        for emp_type in employee_types:
            st.write(f"  - {emp_type}: {resigned_by_type_prev_month.get(emp_type, 0)}명")

    except Exception as e:
        st.error(f"❌ 파일 처리 중 오류 발생: {str(e)}")
