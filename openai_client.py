#3. GPT가 파싱데이터 매칭해줌
import os
from openai import AzureOpenAI
from dotenv import load_dotenv
import os

load_dotenv()


client = AzureOpenAI(
    api_key=os.getenv("AZURE_OPENAI_KEY"),
    api_version=os.getenv("AZURE_OPENAI_API_VERSION"),
    azure_endpoint=os.getenv("AZURE_OPENAI_ENDPOINT")
)

def extract_fields_from_text(text: str, deployment_name="yhj-gpt-4.1-mini") -> dict:
    system_prompt = """당신은 전문 RFP 분석가입니다. 아래 문서에서 다음 정보를 반드시 JSON 형식(예시: {"Project명": "...", ...})으로 추출하세요. 각 항목은 반드시 '항목명: 값' 형태로 출력하세요. :
- Project명 
- Account
- 고객사 주관 조직
- 사업 개요
- 예상 수행 시작 기간
- 예상 수행 종료 기간
- 제안서 마감
- 제안 설명회
- 사업 총 규모

[주의사항]
- 항목이 문서에 없을 경우, 'N/A'로 표기하세요.
- 반드시 JSON 형식으로만 답변하세요. 예시:
{
  "Project명": "스마트캡처플랜",
  "Account": "모두의연구소",
  ...
}
"""

    user_prompt = f"""
다음 텍스트에서 정보를 추출해주세요:
{text}
"""

    response = client.chat.completions.create(
        model=deployment_name,
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt}
        ],
        temperature=0.4
    )

    reply = response.choices[0].message.content.strip()
    
    return reply


