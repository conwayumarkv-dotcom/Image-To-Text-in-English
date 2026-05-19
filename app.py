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

# 🛠️ 여러 모델을 순차적으로 시도하는 똑똑한 워커 함수
def gemini_api_worker_with_fallback(client, pil_image, prompt, result):
    # 사용할 모델 후보군 (빠르고 무료 할당량이 넉넉한 순서)
    models_to_try = ['gemini-1.5-flash-8b', 'gemini-1.5-flash', 'gemini-1.5-pro']
    
    last_error = None
    for model in models_to_try:
        try:
            # temperature를 0으로 설정하여 무작위성 배제 및 토큰 절약
            response = client.models.generate_content(
                model=model, 
                contents=[pil_image, prompt],
                config=types.GenerateContentConfig(temperature=0.0)
            )
            result["text"] = response.text
            result["status"] = "success"
            result["used_model"] = model  # 성공한 모델 기록
            return
        except Exception as e:
            error_str = str(e).upper()
            last_error = e
            
            # 할당량(Quota) 문제라면 다음 모델로 넘어가서 재시도
            if "429" in error_str or "RESOURCE_EXHAUSTED" in error_str or "QUOTA" in error_str:
                time.sleep(1) # 아주 짧은 대기 후 다음 모델 시도
                continue 
            else:
                # 다른 치명적 에러라면 즉시 중단
                break
                
    result["status"] = "error"
    result["error"] = str(last_error)

try:
    api_key = st.secrets["GEMINI_API_KEY"]
    client = genai.Client(api_key=api_key)
    
    uploaded_files = st.file_uploader("사진을 업로드하세요", type=["jpg", "jpeg", "png"], accept_multiple_files=True)

    if uploaded_files and st.button("Word 파일로 변환하기 ✨"):
        doc = Document()
        style = doc.styles['Normal']
        style.font.name = 'Arial'
        
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        success_count = 0
        
        for idx, file in enumerate(uploaded_files):
            raw_img = Image.open(file)
            if raw_img.mode != 'RGB': raw_img = raw_img.convert('RGB')
            raw_img.thumbnail((1200, 1200), Image.Resampling.LANCZOS)
            
            result = {"status": "pending", "text": None, "error": None, "used_model": None}
            prompt = "Extract text. Titles with [HEADING]. Dialogue with [NAME] Speaker: . Remove line numbers. Continuous paragraphs. Pure text only."
            
            # 스레드 시작
            thread = threading.Thread(target=gemini_api_worker_with_fallback, args=(client, raw_img, prompt, result))
            thread.start()
            
            start_p = int((idx / len(uploaded_files)) * 100)
            end_p = int(((idx + 1) / len(uploaded_files)) * 100)
            
            p = start_p
            while thread.is_alive():
                if p < end_p - 2:
                    p += 1
                    progress_bar.progress(p)
                time.sleep(0.05)
            thread.join()
            
            if result["status"] == "success":
                success_count += 1
                progress_bar.progress(end_p)
                
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
                st.error(f"❌ '{file.name}' 분석 실패 (모든 모델 한도 초과): {result['error']}")

        if success_count > 0:
            buffer = BytesIO()
            doc.save(buffer)
            st.download_button("📥 워드 다운로드", data=buffer.getvalue(), file_name="output.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

except Exception as e:
    st.error(f"시스템 에러: {str(e)}")
