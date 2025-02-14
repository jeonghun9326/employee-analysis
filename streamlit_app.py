import os
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from datetime import datetime, timedelta
import tempfile

# 📌 현재 날짜 기준 전월 및 당월 계산
today = datetime.today()
current_month = today.strftime("%Y-%m")
previous_month = (today.replace(day=1) - timedelta(days=1)).strftime("%Y-%m")
previous_month_last_day = (today.replace(day=1) - timedelta(days=1)).strftime("%Y-%m-%d")

# 📌 분석 대상 컬럼 설정
date_columns = ["입사일", "퇴사일"]
employee_types = ["정규직", "계약직", "파견직", "임원"]  # 가나다순 정렬

# 📌 Streamlit UI
st.title("📊 다중 엑셀 병합 및 인원 분석")
st.write("엑셀 파일을 업로드하면 자동으로 병합 후 분석을 수행합니다.")

# 📌 다중 엑셀 파일 업로드
uploaded_files = st.file_uploader("📂 엑셀 파일을 선택하세요", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    try:
        # 📌 임시 폴더 생성
        temp_dir = tempfile.mkdtemp()
        merged_excel_path = os.path.join(temp_dir, "merged_excel.xlsx")

        # 📌 업로드된 파일 저장
        file_paths = []
        for uploaded_file in uploaded_files:
            file_path = os.path.join(temp_dir, uploaded_file.name)
            with open(file_path, "wb") as f:
                f.write(uploaded_file.read())
            file_paths.append(file_path)

        # 📌 엑셀 병합 함수 실행
        def merge_excel_files(files, output_file):
            with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
                for file in files:
                    try:
                        wb = load_workbook(file, data_only=True)
                        sheet_names = wb.sheetnames  

                        if not sheet_names:
                            st.warning(f"⚠️ 파일 `{os.path.basename(file)}` 에 사용 가능한 시트가 없어 건너뜁니다.")
                            continue

                        for sheet_name in sheet_names:
                            ws = wb[sheet_name]
                            data = [[cell.value for cell in row] for row in ws.iter_rows()]
                            
                            if not data or all(all(cell is None for cell in row) for row in data):
                                st.warning(f"⚠️ 파일 `{os.path.basename(file)}` 의 시트 `{sheet_name}` 가 비어 있어 건너뜁니다.")
                                continue

                            header_row_index = None
                            for idx, row in enumerate(data):
                                if row[0] == "No":
                                    header_row_index = idx
                                    break

                            if header_row_index is not None:
                                headers = data[header_row_index]
                                df = pd.DataFrame(data[header_row_index + 1:], columns=headers)
                            else:
                                df = pd.DataFrame(data[1:], columns=data[0])

                            sheet_name_trimmed = os.path.splitext(os.path.basename(file))[0][:31]
                            df.to_excel(writer, sheet_name=sheet_name_trimmed, index=False)

                    except Exception as e:
                        st.error(f"🚨 파일 `{os.path.basename(file)}` 처리 중 오류 발생: {e}")

        merge_excel_files(file_paths, merged_excel_path)
        st.success("✅ 엑셀 파일 병합 완료!")

        # 📌 병합된 엑셀 파일 분석 및 추가 시트 생성
        with pd.ExcelWriter(merged_excel_path, engine="openpyxl", mode="a") as writer:
            sheets = pd.read_excel(merged_excel_path, sheet_name=None, engine="openpyxl")

            for sheet_name, df in sheets.items():
                st.subheader(f"📄 시트 이름: {sheet_name}")

                # 📌 컬럼명 정리
                if "Starting Date" in df.columns:
                    df.rename(columns={"Starting Date": "입사일"}, inplace=True)
                df.columns = df.columns.str.strip()

                # 📌 날짜 변환
                if "입사일" in df.columns:
                    df["입사일"] = pd.to_datetime(df["입사일"], errors="coerce").dt.strftime("%Y-%m-%d")
                if "퇴사일" not in df.columns:
                    df["퇴사일"] = None
                if "Remark" in df.columns:
                    df.loc[df["Remark"].astype(str).str.startswith("Resigned and last working"), "퇴사일"] = previous_month_last_day

                # 📌 "사원구분명" 컬럼 자동 생성
                if "사원구분명" not in df.columns:
                    df["사원구분명"] = None
                if "Contract Type" in df.columns:
                    df.loc[df["Contract Type"].astype(str).str.contains("FDC", na=False), "사원구분명"] = "계약직"
                    df.loc[df["Contract Type"].astype(str).str.contains("UDC", na=False), "사원구분명"] = "정규직"

                # 📌 날짜 변환 (YYYY-MM)
                for col in date_columns:
                    df[col] = pd.to_datetime(df[col], errors="coerce").dt.strftime("%Y-%m")

                # 📌 퇴사자 데이터 저장
                if "퇴사일" in df.columns and "부서명" in df.columns and "성명" in df.columns and "직급명" in df.columns:
                    resigned_details = df[df["퇴사일"] == previous_month][["부서명", "성명", "직급명"]]
                    if not resigned_details.empty:
                        st.write(f"📌 전월({previous_month}) 퇴사자 상세 내역:")
                        st.dataframe(resigned_details)
                        resigned_details.to_excel(writer, sheet_name=f"{sheet_name}_퇴사자", index=False)

                # 📌 입사자 데이터 저장
                if "입사일" in df.columns and "부서명" in df.columns and "성명" in df.columns and "직급명" in df.columns:
                    new_hires_details = df[df["입사일"] == previous_month][["부서명", "성명", "직급명"]]
                    if not new_hires_details.empty:
                        st.write(f"📌 전월({previous_month}) 입사자 상세 내역:")
                        st.dataframe(new_hires_details)
                        new_hires_details.to_excel(writer, sheet_name=f"{sheet_name}_입사자", index=False)

        # 📌 병합된 엑셀 다운로드 제공
        st.download_button(label="📥 병합된 엑셀 다운로드", data=open(merged_excel_path, "rb").read(), file_name="merged_excel.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        st.error(f"❌ 오류 발생: {e}")

