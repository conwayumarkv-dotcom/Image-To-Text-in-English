import streamlit as st
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from io import BytesIO
import time
import re
from PIL import Image
from google import genai
from google.genai import types

# 1. 페이지 기본 설정 및 디자인
st.set_page_config(
    page_title="Image To Text in English",
    page_icon="📝",
    layout="centered"
)

# 세련되고 직관적인 커스텀 CSS
st.markdown("""
    <style>
    .stApp { background-color: #FDFBF6; }
    .main-title { font-size: 42px !important; font-weight: 800; color: #5C715E; text-align: center; margin-bottom: 8px; }
    .sub-title { font-size: 18px; color: #7A7A7A; text-align: center; margin-bottom: 10px; }
    .author-footer { font-size: 14px; color: #B2B2B2; text-align: right; margin-bottom: 45px; }
    div.stButton > button:first-child { background-color: #94A69A; color: white; width: 100%; border-radius: 10px; padding: 12px; }
    .percent-text { font-size: 20px; font-weight: 700; color: #5C715E; }
    </style>
""", unsafe_allow_html=True)

st.markdown('<p class="main-title">Image To Text in English</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-title">사진 속 지문을 인식하여 편집 가능한 워드 문서(.docx)로 변환합니다.</p>', unsafe_allow_html=True)
st.markdown('<div class="author-footer">(Made by Manju)</div>', unsafe_allow_html=True)

try:
    api_key = st.secrets["GEMINI_API_KEY"]
    client = genai.Client(api_key=api_key)
    
    uploaded_files = st.file_uploader("변환할 영어 지문 사진을 업로드하세요", type=["jpg", "jpeg", "png"], accept_multiple_files=True)

    if uploaded_files and st.button("Word 파일로 변환하기 ✨"):
        doc = Document()
        style = doc.styles['Normal']
        style.font.name = 'Arial'
        style.font.size = Pt(11)
        
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        # 🛠️ [핵심 개선] 하루 한도 1,500회인 1.5-flash로 모델 교체
        model_name = 'gemini-1.5-flash'
        
        for idx, file in enumerate(uploaded_files):
            file_bytes = file.read()
            status_text.text(f"⏳ [{idx+1}/{len(uploaded_files)}] '{file.name}' 분석 중...")
            
            try:
                # 이미지 최적화 (토큰 절약)
                raw_img = Image.open(BytesIO(file_bytes))
                if raw_img.mode != 'RGB': raw_img = raw_img.convert('RGB')
                raw_img.thumbnail((1200, 1200), Image.Resampling.LANCZOS)
                img_buffer = BytesIO()
                raw_img.save(img_buffer, format="JPEG", quality=85)
                pil_image = Image.open(BytesIO(img_buffer.getvalue()))
                
                prompt = """
                Extract English text.
                - Titles: [HEADING] Title
                - Dialogue: [NAME] Name: Dialogue
                - Merge split lines, remove line numbers, no markdown, pure text only.
                """
                
                response = client.models.generate_content(
                    model=model_name,
                    contents=[pil_image, prompt]
                )
                extracted_text = response.text
                
                # 워드 문서 작성
                doc.add_paragraph(f"Source: {file.name}").runs[0].font.color.rgb = RGBColor(128, 128, 128)
                
                for line in extracted_text.split('\n'):
                    line = line.strip()
                    if not line: continue
                    p = doc.add_paragraph()
                    
                    if line.startswith("[HEADING]"):
                        run = p.add_run(line.replace("[HEADING]", "").strip())
                        run.bold = True
                        run.font.size = Pt(13)
                    elif line.startswith("[NAME]"):
                        content = line.replace("[NAME]", "").strip()
                        match = re.match(r"^([^:]+:)(.*)$", content)
                        if match:
                            p.add_run(match.group(1)).bold = True
                            p.add_run(match.group(2))
                        else:
                            p.add_run(content)
                    else:
                        p.add_run(line.replace("**", ""))
                
                doc.add_page_break()
                progress_bar.progress((idx + 1) / len(uploaded_files))
                
            except Exception as e:
                st.error(f"❌ '{file.name}' 변환 실패: {str(e)}")
                continue

        # 다운로드 버튼
        docx_buffer = BytesIO()
        doc.save(docx_buffer)
        docx_buffer.seek(0)
        st.download_button("📥 워드 파일 다운로드", data=docx_buffer, file_name="Converted_Texts.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

except KeyError:
    st.error("🔒 설정 오류: API Key가 등록되지 않았습니다.")
except Exception as e:
    st.error(f"시스템 오류 발생: {str(e)}")
