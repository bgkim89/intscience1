import streamlit as st
import pandas as pd
import re
from io import BytesIO
from docx import Document

def extract_info_from_docx(file):
    doc = Document(file)
    results = []

    # Word는 페이지 구분 정보가 명확하지 않아서 페이지 단위로 나누는 기능은 제한됨.
    # 여기서는 각 표가 한 페이지에 하나씩 있다고 가정하고 각 표 앞 단락에서 5자리 숫자 추출

    paragraphs = doc.paragraphs
    tables = doc.tables
    para_texts = [p.text.strip() for p in paragraphs if p.text.strip() != ""]

    para_idx = 0
    for table in tables:
        # 표 앞 단락에서 5자리 숫자 찾기
        number_found = ""
        while para_idx < len(para_texts):
            match = re.search(r"\b\d{5}\b", para_texts[para_idx])
            para_idx += 1
            if match:
                number_found = match.group()
                break

        try:
            cell_b = table.cell(0, 0).text.strip()
        except:
            cell_b = ""

        try:
            cell_c = table.cell(1, 0).text.strip()
        except:
            cell_c = ""

        results.append([number_found, cell_b, cell_c])

    return results

def to_csv(data):
    df = pd.DataFrame(data, columns=["A열(숫자)", "B열(1행1열)", "C열(2행1열)"])
    return df.to_csv(index=False).encode("utf-8-sig")

st.title("Word ➜ CSV 변환기")

uploaded_file = st.file_uploader("Word (.docx) 파일 업로드", type="docx")

if uploaded_file:
    st.success("파일 업로드 완료! 변환을 시작합니다.")
    data = extract_info_from_docx(uploaded_file)

    if data:
        csv_data = to_csv(data)
        st.download_button(
            label="📥 CSV 다운로드",
            data=csv_data,
            file_name="converted.csv",
            mime="text/csv"
        )
        st.dataframe(pd.DataFrame(data, columns=["A열(숫자)", "B열(1행1열)", "C열(2행1열)"]))
    else:
        st.warning("표 또는 숫자를 추출할 수 없습니다. Word 파일을 확인해 주세요.")
