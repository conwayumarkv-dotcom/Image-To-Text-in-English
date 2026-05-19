import streamlit as st
from docx import Document
from docx.shared import Pt, RGBColor
from io import BytesIO
import time
import re
import threading
from PIL import Image
from google import genai
from google.genai import types

# 1. 페이지 기본 설정 및 디자인
st.set_page_config(page_title="Image To Text in English", page_icon="📝", layout="centered")

st.markdown("""
    <style>
    .stApp { background-color: #FDFBF6; }
    .main-title { font-size: 42px !important; font-weight: 800; color: #5C715E; text-align: center; margin-bottom: 8px; }
    .sub-title { font-size: 18px; color: #7A7A7A; text-align: center; margin-bottom: 10px; }
    .author-footer { font-size: 13px; color: #A3A3A3; text-align: right; margin-bottom: 45px; }
    div.stButton > button:first-child { background-color: #94A69A; color: white; width: 100%; border-radius: 10px; padding: 12px; }
    .percent-text { font-size: 20px; font-weight: 700; color: #5C715E; }
    </style>
""", unsafe_allow_html=True)

st.markdown('<p class="main-title">Image To Text in English</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-title">사진 속 지문을 인식하여 편집 가능한 워드 문서(.docx)로 변환합니다.</p>', unsafe_allow_html=True)

# 🛠️ 비동기 API 워커 함수
def gemini_api_worker(client, pil_image, prompt, result):
    try:
        # 선생님 환경에서 정상 작동이 검증된 2.5-flash 모델 적용
        response = client.models.generate_content(
            model='gemini-2.5-flash',
            contents=[pil_image, prompt],
            config=types.GenerateContentConfig(temperature=0.0)
        )
        result["text"] = response.text
        result["status"] = "success"
    except Exception as e:
        result["status"] = "error"
        result["error"] = str(e).upper()

try:
    api_key = st.secrets["GEMINI_API_KEY"]
    client = genai.Client(api_key=api_key)
    
    uploaded_files = st.file_uploader("사진을 업로드하세요 (복수 선택 가능)", type=["jpg", "jpeg", "png"], accept_multiple_files=True)

    if uploaded_files and st.button("Word 파일로 변환하기 ✨"):
        doc = Document()
        style = doc.styles['Normal']
        style.font.name = 'Arial'
        
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        success_count = 0
        quota_blocked = False
        
        for idx, file in enumerate(uploaded_files):
            # 이미지 최적화
            raw_img = Image.open(file)
            if raw_img.mode != 'RGB': raw_img = raw_img.convert('RGB')
            raw_img.thumbnail((1200, 1200), Image.Resampling.LANCZOS)
            
            result = {"status": "pending", "text": None, "error": None}
            prompt = "Extract text. Titles with [HEADING]. Dialogue with [NAME] Speaker: . Remove line numbers. Continuous paragraphs. Pure text only."
            
            status_text.text(f"⏳ [{idx+1}/{len(uploaded_files)}] '{file.name}' 지문을 분석 중입니다...")
            
            # 스레드 시작 (백그라운드에서 AI 분석)
            thread = threading.Thread(target=gemini_api_worker, args=(client, raw_img, prompt, result))
            thread.start()
            
            # 🛠️ 부드러운 진행률 애니메이션 루프
            start_p = int((idx / len(uploaded_files)) * 100)
            end_p = int(((idx + 1) / len(uploaded_files)) * 100)
            
            p = start_p
            while thread.is_alive():
                # 목표치의 98%까지만 가상으로 채우며 사용자 대기 시간 지루함 방지
                if p < end_p - 2:
                    p += 1
                    progress_bar.progress(p)
                time.sleep(0.05) # 애니메이션 속도 제어
            
            thread.join()
            
            if result["status"] == "success":
                success_count += 1
                progress_bar.progress(end_p) # 완료 시 해당 파일 목표 퍼센트 꽉 채우기
                
                # 워드 조립
                p_src = doc.add_paragraph()
                p_src.add_run(f"Source: {file.name}").font.color.rgb = RGBColor(128, 128, 128)
                
                for line in result["text"].split('\n'):
                    if not line.strip(): continue
                    para = doc.add_paragraph()
                    
                    if line.startswith("[HEADING]"):
                        run = para.add_run(line.replace("[HEADING]", "").strip())
                        run.bold = True
                        run.font.size = Pt(13)
                    elif line.startswith("[NAME]"):
                        content = line.replace("[NAME]", "").strip()
                        match = re.match(r"^([^:]+:)(.*)$", content)
                        if match:
                            para.add_run(match.group(1)).bold = True
                            para.add_run(match.group(2))
                        else:
                            para.add_run(content)
                    else:
                        para.add_run(line.replace("**", ""))
                doc.add_page_break()
                
            else:
                # 429 에러(하루 20회 한도) 필터링
                if "429" in result["error"] or "RESOURCE_EXHAUSTED" in result["error"] or "QUOTA" in result["error"]:
                    quota_blocked = True
                    break
                else:
                    st.error(f"❌ '{file.name}' 분석 실패: {result['error']}")

        if success_count > 0:
            if quota_blocked:
                st.warning("⚠️ 하루 구글 무료 이용 한도(20회)가 소진되었습니다. 지금까지 변환 성공한 파일들만 저장합니다.")
            else:
                progress_bar.progress(100)
                status_text.text("🎉 선택하신 모든 영어 지문이 워드 파일로 완성되었습니다!")
            
            buffer = BytesIO()
            doc.save(buffer)
            st.download_button("📥 워드 다운로드", data=buffer.getvalue(), file_name="Converted_Texts.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            
        elif quota_blocked and success_count == 0:
            status_text.empty()
            st.error("⚠️ 오늘 제공되는 구글의 무료 변환 한도(20회)가 모두 소진되어 지금은 변환을 시작할 수 없습니다. 내일 다시 이용해 주세요.")
        else:
            status_text.empty()
            st.error("❌ 변환된 지문이 없습니다.")

except Exception as e:
    st.error(f"시스템 에러: {str(e)}")
