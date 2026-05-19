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

# 🛠️ 백그라운드 API 워커
def gemini_api_worker(client, model, pil_image, prompt, result):
    try:
        # 모델 호출 (안정적인 gemini-2.0-flash 사용)
        response = client.models.generate_content(model=model, contents=[pil_image, prompt])
        result["text"] = response.text
        result["status"] = "success"
    except Exception as e:
        result["status"] = "error"
        result["error"] = str(e)

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
        
        # 🛠️ 모델명 변경 (1.5가 안 잡힐 경우 2.0으로 시도)
        model_name = 'gemini-2.0-flash'
        
        for idx, file in enumerate(uploaded_files):
            # 이미지 전처리
            raw_img = Image.open(file)
            if raw_img.mode != 'RGB': raw_img = raw_img.convert('RGB')
            raw_img.thumbnail((1200, 1200), Image.Resampling.LANCZOS)
            
            # API 호출용 스레드
            result = {"status": "pending", "text": None, "error": None}
            thread = threading.Thread(target=gemini_api_worker, args=(client, model_name, raw_img, "Extract text...", result))
            thread.start()
            
            # 애니메이션 루프
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
                progress_bar.progress(end_p)
                # 워드 조립
                for line in result["text"].split('\n'):
                    if not line.strip(): continue
                    para = doc.add_paragraph()
                    para.add_run(line.replace("**", ""))
                doc.add_page_break()
            else:
                st.error(f"❌ '{file.name}' 분석 실패: {result['error']}")

        # 다운로드
        buffer = BytesIO()
        doc.save(buffer)
        st.download_button("📥 워드 다운로드", data=buffer.getvalue(), file_name="output.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

except Exception as e:
    st.error(f"시스템 에러: {str(e)}")
