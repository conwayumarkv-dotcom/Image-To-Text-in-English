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
    .main-title {
        font-size: 42px !important; 
        font-weight: 800;
        color: #5C715E;
        text-align: center;
        margin-bottom: 8px;
        letter-spacing: -1.5px;
        white-space: nowrap;
    }
    .sub-title {
        font-size: 18px;
        color: #7A7A7A;
        text-align: center;
        margin-bottom: 10px;
        font-weight: 400;
        letter-spacing: -0.5px;
    }
    .author-footer {
        font-size: 13px;
        color: #A3A3A3; 
        text-align: right;
        margin-bottom: 45px;
        padding-right: 5px;
        font-family: 'Arial', sans-serif;
    }
    .stFileUploader {
        border-radius: 15px !important;
        background-color: white;
        box-shadow: 0 4px 15px rgba(0,0,0,0.05);
        padding: 10px;
    }
    div.stButton > button:first-child {
        background-color: #94A69A;
        color: white;
        border-radius: 10px;
        border: none;
        padding: 12px 40px;
        font-weight: 600;
        font-size: 16px;
        width: 100%;
        transition: all 0.3s ease;
        box-shadow: 0 2px 6px rgba(0,0,0,0.05);
    }
    div.stButton > button:first-child:hover { background-color: #7E8F83; }
    .status-box {
        padding: 25px;
        border-radius: 20px;
        background-color: white;
        margin-top: 30px;
        box-shadow: 0 4px 20px rgba(0,0,0,0.08);
        text-align: center;
    }
    .percent-text {
        font-size: 20px;
        font-weight: 700;
        color: #5C715E;
        margin-bottom: 8px;
        text-align: left;
    }
    </style>
""", unsafe_allow_html=True)

st.markdown('<p class="main-title">Image To Text in English</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-title">사진 속 지문을 인식하여 편집 가능한 워드 문서(.docx)로 변환합니다.</p>', unsafe_allow_html=True)
st.markdown('<div class="author-footer">© TOP English Academy. All rights reserved.</div>', unsafe_allow_html=True)

try:
    api_key = st.secrets["GEMINI_API_KEY"]
    client = genai.Client(api_key=api_key)
    
    uploaded_files = st.file_uploader(
        "변환할 영어 지문 사진을 업로드하세요 (복수 선택 가능)", 
        type=["jpg", "jpeg", "png"],
        accept_multiple_files=True
    )

    if uploaded_files:
        st.write(f"📂 **{len(uploaded_files)}개**의 파일이 선택되었습니다.")
        
        if st.button("Word 파일로 변환하기 ✨"):
            
            doc = Document()
            style = doc.styles['Normal']
            style.font.name = 'Arial'
            style._element.rPr.get_or_add_rFonts().set(qn('w:ascii'), 'Arial')
            style._element.rPr.get_or_add_rFonts().set(qn('w:hAnsi'), 'Arial')
            style.font.size = Pt(11)
            
            percent_display = st.empty()  
            progress_bar = st.progress(0)  
            status_text = st.empty()       
            
            total_files = len(uploaded_files)
            
            # ⚡ 토큰 효율성 및 속도가 최적화된 최신 플래그십 플래시 모델
            model_name = 'gemini-2.5-flash'
            
            success_count = 0     
            quota_blocked = False
            last_extracted_text = "" 
            
            for idx, file in enumerate(uploaded_files):
                file_bytes = file.read()
                extracted_text = "" 
                
                try:
                    raw_img = Image.open(BytesIO(file_bytes))
                    if raw_img.mode != 'RGB':
                        raw_img = raw_img.convert('RGB')
                    # 📉 이미지 해상도를 적절히 제한하여 입력 이미지 토큰 절약
                    raw_img.thumbnail((1200, 1200), Image.Resampling.LANCZOS)
                    
                    compressed_buffer = BytesIO()
                    raw_img.save(compressed_buffer, format="JPEG", quality=85)
                    pil_image = Image.open(BytesIO(compressed_buffer.getvalue()))
                except Exception:
                    try:
                        pil_image = Image.open(BytesIO(file_bytes))
                    except Exception:
                        st.error(f"❌ '{file.name}' 이미지를 로드하는 과정에서 에러가 발생했습니다.")
                        continue
                
                # ✂️ [토큰 다이어트] 불필요한 수식어를 빼고 구조를 단순화하여 입력 토큰을 절약한 프롬프트
                prompt = """
                Extract English text from image.
                Rules:
                - Add '[HEADING]' before main titles.
                - Add '[NAME]' before speaker names (e.g. [NAME] Name:).
                - Remove all line numbers completely.
                - Use continuous paragraphs; ignore image hard line breaks.
                - No markdown. Output ONLY raw extracted text.
                """
                
                status_text.text(f"⏳ [{idx+1}/{total_files}] '{file.name}' 사진 속 영어 지문을 분석 중입니다...")
                
                current_percent = int((idx / total_files) * 100)
                target_percent = int(((idx + 1) / total_files) * 100)
                
                for p in range(current_percent, min(target_percent, current_percent + 5)):
                    percent_display.markdown(f'<p class="percent-text">⏳ 변환 진행률: {p}%</p>', unsafe_allow_html=True)
                    progress_bar.progress(p)
                    time.sleep(0.02)

                max_retries = 3
                
                for attempt in range(max_retries):
                    try:
                        response = client.models.generate_content(
                            model=model_name,
                            contents=[pil_image, prompt],
                            config=types.GenerateContentConfig(temperature=0.0)
                        )
                        extracted_text = response.text
                        last_extracted_text = extracted_text
                        break 
                        
                    except Exception as api_err:
                        error_msg = str(api_err).upper()
                        
                        if "429" in error_msg or "RESOURCE_EXHAUSTED" in error_msg or "QUOTA" in error_msg:
                            quota_blocked = True
                            break 
                            
                        if attempt < max_retries - 1:
                            status_text.text(f"⏳ 서버 응답 지연. 다시 연결 중입니다... ({attempt+1}/{max_retries})")
                            time.sleep(2)
                            continue
                        else:
                            st.error(f"❌ '{file.name}' 변환 실패 상세 원인: {str(api_err)}")
                            break

                if quota_blocked:
                    break

                if extracted_text:
                    try:
                        success_count += 1
                        
                        for p in range(min(target_percent, current_percent + 5), target_percent + 1):
                            percent_display.markdown(f'<p class="percent-text">⏳ 변환 진행률: {p}%</p>', unsafe_allow_html=True)
                            progress_bar.progress(p)
                            time.sleep(0.01)
                        
                        status_text.text(f"✅ [{idx+1}/{total_files}] 지문 변환 및 서식 정리 완료!")
                        
                        p_src = doc.add_paragraph()
                        r_src = p_src.add_run(f"▪ Source: {file.name}")
                        r_src.font.size = Pt(10)
                        r_src.font.color.rgb = RGBColor(128, 128, 128)  
                        
                        paragraphs = extracted_text.split('\n')
                        for para_text in paragraphs:
                            clean_text = para_text.strip()
                            if not clean_text:
                                continue
                            
                            p_tag = doc.add_paragraph()
                            
                            if clean_text.startswith("[HEADING]"):
                                heading_content = clean_text.replace("[HEADING]", "").strip()
                                run = p_tag.add_run(heading_content)
                                run.bold = True
                                run.font.size = Pt(13) 
                                p_tag.paragraph_format.space_before = Pt(12) 
                                p_tag.paragraph_format.space_after = Pt(6)   
                                
                            elif clean_text.startswith("[NAME]"):
                                name_content = clean_text.replace("[NAME]", "").strip()
                                match = re.match(r"^([^:]+:)(.*)$", name_content)
                                if match:
                                    name_part = match.group(1)   
                                    dialogue_part = match.group(2) 
                                    r_name = p_tag.add_run(name_part)
                                    r_name.bold = True
                                    p_tag.add_run(dialogue_part)
                                else:
                                    p_tag.add_run(name_content)
                                    
                            else:
                                plain_content = clean_text.replace("**", "")
                                p_tag.add_run(plain_content)
                        
                        if idx < total_files - 1:
                            doc.add_page_break()
                                        
                    except Exception as word_err:
                        st.warning(f"⚠️ '{file.name}' 문서 디자인 조립 중 경미한 지연이 생겨 안전하게 다음으로 패스합니다.")
                        continue
                    
                    if idx < total_files - 1:
                        time.sleep(1.0)

            if success_count > 0:
                if quota_blocked:
                    st.warning("⚠️ 하루 무료 제공량이 모두 소진되어 작업이 중단되었습니다. 현재까지 성공적으로 변환된 파일만 저장합니다.")
                else:
                    percent_display.markdown('<p class="percent-text" style="color:#0D9488;">🎉 변환 진행률: 100%</p>', unsafe_allow_html=True)
                    progress_bar.progress(100)
                    status_text.text("🎉 선택하신 모든 영어 지문이 워드 파일로 완성되었습니다!")
                
                docx_buffer = BytesIO()
                doc.save(docx_buffer)
                docx_buffer.seek(0)
                
                st.markdown('<div class="status-box">', unsafe_allow_html=True)
                st.download_button(
                    label="📥 변환된 Word 파일 다운로드",
                    data=docx_buffer,
                    file_name="Converted_English_Texts.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                st.markdown('</div>', unsafe_allow_html=True)
                
            elif quota_blocked and success_count == 0:
                status_text.empty()
                st.error("⚠️ 오늘 구글의 하루 제공량을 모두 소진하여 더 이상 변환할 수 없습니다.")
            elif not last_extracted_text: 
                status_text.empty()
                st.error("❌ 변환된 지문이 없습니다. (위의 상세 에러 로그를 확인해 주세요)")

except KeyError:
    st.error("🔒 설정 오류: Streamlit Secrets에 API Key를 등록해 주세요.")
