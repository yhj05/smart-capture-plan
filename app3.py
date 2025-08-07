#1.streamlit 웹 앱
import streamlit as st
from parse_docx import extract_text_from_docx
from openai_client import extract_fields_from_text
from excel_writer4 import write_to_excel
import json
import tempfile
import os
from azure.storage.blob import BlobServiceClient
from dotenv import load_dotenv
import re
from datetime import datetime
from openpyxl import load_workbook
from io import BytesIO

# 환경 변수 불러오기
load_dotenv()

connect_str = os.getenv("AZURE_STORAGE_CONNECTION_STRING")
container_name = "rfp-uploads"
template_blob_name = "CapturePlan_Template.xlsx"

st.title("📄 RFP → 캡처플랜 자동 작성!")

uploaded_file = st.file_uploader("RFP 문서를 업로드하세요 (.docx)", type="docx")

# 업로드: Blob에 저장
def upload_to_blob(local_file_path, blob_name):
    blob_service_client = BlobServiceClient.from_connection_string(connect_str)
    blob_client = blob_service_client.get_blob_client(container=container_name, blob=blob_name)
    with open(local_file_path, "rb") as data:
        blob_client.upload_blob(data, overwrite=True)

# 다운로드: Blob에서 템플릿 로딩
def download_template_from_blob():
    blob_service_client = BlobServiceClient.from_connection_string(connect_str)
    blob_client = blob_service_client.get_blob_client(container=container_name, blob=template_blob_name)
    blob_data = blob_client.download_blob().readall()
    return load_workbook(BytesIO(blob_data))  # openpyxl 워크북 객체 반환

# GPT 응답이 JSON 형식 아닐 때 파싱용
def parse_key_value(text):
    result = {}
    for line in text.splitlines():
        if ":" in line:
            key, value = line.split(":", 1)
            result[key.strip()] = value.strip()
    return result

#
if uploaded_file:
    uploaded_filename = os.path.splitext(uploaded_file.name)[0]  # 'some_rfp' 추출
    output_filename = f"{uploaded_filename}_captureplan.xlsx"
    output_local_path = os.path.join(tempfile.gettempdir(), output_filename)

    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        tmp.write(uploaded_file.getbuffer())
        tmp_path = tmp.name
        
        # RFP 문서 Blob에 저장
        upload_to_blob(tmp_path, f"rfp_docs/{uploaded_file.name}")

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

            # 3. 엑셀 템플릿 다운로드 (Blob에서)
            template_wb = download_template_from_blob()

            # 4. 엑셀 작성
            write_to_excel(template_wb, output_local_path, extracted)

            # 5. Blob에 결과 엑셀 업로드
            upload_to_blob(output_local_path, f"results/{output_filename}")

            # 6. 다운로드 버튼 제공
            with open(output_local_path, "rb") as f:
                st.download_button("📥 결과 엑셀 다운로드", f, file_name=output_filename)
        else:
            st.error("❌ GPT 응답을 파싱할 수 없습니다.")
            st.text(raw_response)
