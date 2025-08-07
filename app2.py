#1.streamlit 웹 앱
import streamlit as st
from parse_docx import extract_text_from_docx
from openai_client import extract_fields_from_text
from excel_writer3 import write_to_excel
import json
import tempfile
import os
from azure.storage.blob import BlobServiceClient
from dotenv import load_dotenv
import re
from datetime import datetime

load_dotenv()

connect_str = os.getenv("AZURE_STORAGE_CONNECTION_STRING")
container_name = os.getenv("AZURE_BLOB_CONTAINER_NAME")
# 
TEMPLATE_PATH = r".\CapturePlan_Template.xlsx"
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
OUTPUT_PATH = f"{timestamp}.xlsx"

st.title("📄 RFP → 캡처플랜 자동 작성!")

uploaded_file = st.file_uploader("RFP 문서를 업로드하세요 (.docx)", type="docx")

# 업로드 파일 Blob Storage에 저장하기
def upload_to_blob(local_file_path, blob_name):
    blob_service_client = BlobServiceClient.from_connection_string(connect_str)
    blob_client = blob_service_client.get_blob_client(container=container_name, blob=blob_name)

    with open(local_file_path, "rb") as data:
        blob_client.upload_blob(data, overwrite=True)

def parse_key_value(text):
    result = {}
    for line in text.splitlines():
        if ":" in line:
            key, value = line.split(":", 1)
            result[key.strip()] = value.strip()
    return result

#
if uploaded_file:
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        tmp.write(uploaded_file.getbuffer())
        tmp_path = tmp.name
        
        # 업로드된 RFP를 blob에 저장
        upload_to_blob(tmp_path, f"rfp_docs/{uploaded_file.name}")

        # 결과 엑셀도 저장
        #upload_to_blob(OUTPUT_PATH, "results/결과_캡처플랜.xlsx")

    # 1. 텍스트 추출
    text = extract_text_from_docx(tmp_path)
    st.subheader("📃 문서 내용 미리보기")
    st.text(text[:100])

    # 2. GPT 호출
    with st.spinner("🤖 문서 분석 중..."):
        raw_response = extract_fields_from_text(text, deployment_name="yhj-gpt-4.1-mini")
        try:
            extracted = json.loads(raw_response)
        except Exception:
            extracted = parse_key_value(raw_response)
        if extracted:
            st.success("✅ 분석 완료!")
            st.json(extracted)
            st.write(TEMPLATE_PATH)
            write_to_excel(TEMPLATE_PATH, OUTPUT_PATH, extracted)
            with open(OUTPUT_PATH, "rb") as f:
                st.download_button("📥 결과 엑셀 다운로드", f, file_name=OUTPUT_PATH)
        else:
            st.error("❌ GPT 응답을 파싱할 수 없습니다.")
            st.text(raw_response)