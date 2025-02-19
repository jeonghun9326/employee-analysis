import streamlit as st
from streamlit_app_HR import run_excel_analysis
from streamlit_app_insurance import run_insurance_analysis

# ✅ 기능 선택 UI
def select_feature():
    """ Streamlit UI에서 사용자가 사용할 기능을 선택하는 함수 """
    return st.sidebar.selectbox(
        "📌 사용할 기능을 선택하세요",
        ["엑셀 병합 및 인원 분석", "4대보험료 검증 시스템", "급여 업무 시스템", "채용 분석 시스템"]
    )

# ✅ 기능 실행
def main():
    st.title("📊 다중 엑셀 분석 시스템")

    feature_option = select_feature()

    if feature_option == "엑셀 병합 및 인원 분석":
        run_excel_analysis()  # ✅ 엑셀 병합 및 인원 분석 실행
    if feature_option == "4대보험료 검증 시스템":
        run_insurance_analysis()  # ✅ 엑셀 병합 및 인원 분석 실행

if __name__ == "__main__":
    main()
