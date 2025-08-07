
# 📄 RFP → 캡처플랜 자동화 웹 앱

이 프로젝트는 RFP(제안요청서) 문서를 업로드하면 GPT를 통해 내용을 분석하고, 캡처플랜 엑셀 문서를 자동으로 생성해주는 **Streamlit 기반 웹 애플리케이션**입니다.  
분석된 데이터는 지정된 Azure Blob Storage에 저장되며, 템플릿 파일도 Blob에서 자동으로 불러옵니다.

---

## 🧩 주요 기능

- `.docx` 형식의 RFP 문서를 업로드
- GPT를 통한 RFP 주요 항목 추출
- 캡처플랜 엑셀 템플릿에 자동 입력
- 엑셀 결과물 다운로드 및 Azure Blob Storage에 저장

---

## 📁 프로젝트 구조

smart_captureplan/
├── app3.py # Streamlit 웹 앱 실행 파일
├── excel_writer4.py # 엑셀 작성 로직
├── parse_docx.py # docx 텍스트 추출
├── openai_client.py # GPT 응답 처리
├── .env # 환경 변수 설정 파일


---

## ⚙️ 설치 및 실행 방법

### 1. Python 환경 설정

```bash
python -m venv venv
source venv/bin/activate   # (Windows: venv\Scripts\activate)
pip install -r requirements.txt
```


### 2. ▶️ 실행
bash
streamlit run app3.py

브라우저에서 자동으로 열리지 않으면 다음 주소로 이동하세요:
arduino
https://yhj-wrbapp-004-c4bjfta2egauhagp.canadacentral-01.azurewebsites.net/


### 3. 📦 요구사항 (requirements.txt 실제 사용한 패키지)
streamlit
openai
python-dotenv
python-docx
openpyxl
azure-storage-blob

### 4. 🌐 Azure 리소스 구성 요약
Blob Storage
    Container: rfp-uploads
    업로드 폴더: rfp_docs/
    결과물 폴더: results/
    템플릿 파일: CapturePlan_Template.xlsx
Azure OpenAI
    모델 배포 이름: yhj-gpt-4.1-mini 


---

## 📌 참고 사항
템플릿 파일의 병합 셀 위치가 변경되면 excel_writer4.py도 함께 수정해야 합니다.
GPT 응답이 JSON 형식이 아닐 경우에도 Key-Value 파싱을 통해 대응합니다.

