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
            
            # 레이아웃 고정용 칸 초기화
            percent_display = st.empty()  
            progress_bar = st.progress(0)  
            status_text = st.empty()       
            
            percent_display.markdown('<p class="percent-text">⏳ 문서 생성률: 0%</p>', unsafe_allow_html=True)
            progress_bar.progress(0.0)
            status_text.text("🔄 인공지능 교사가 지문 판독을 시작합니다. 잠시만 기다려주세요...")
            
            total_files = len(uploaded_files)
            
            # [404 에러 완벽 해결] 현재 구글 SDK v1beta 버전 환경에서 오류 없는 정식 표준 모델로 매핑
            model_name = 'gemini-2.0-flash'
            
            success_count = 0     
            quota_blocked = False
            api_error_occurred = False
            last_extracted_text = "" 
            
            progress_per_file = 100.0 / total_files
            
            for idx, file in enumerate(uploaded_files):
                file_bytes = file.read()
                extracted_text = "" 
                
                start_p = int(idx * progress_per_file)
                mid_p = int((idx + 0.6) * progress_per_file)
                end_p = int((idx + 1) * progress_per_file)
                
                status_text.text(f"📝 [{idx+1}/{total_files}] '{file.name}' 파일을 읽어오는 중입니다...")
                for p in range(start_p, min(mid_p, 99) + 1):
                    percent_display.markdown(f'<p class="percent-text">⏳ 문서 생성률: {p}%</p>', unsafe_allow_html=True)
                    progress_bar.progress(p / 100.0)
                    time.sleep(0.01)
                
                try:
                    # [토큰 다이어트 1] 화질은 유지하되 픽셀 압축을 최적화하여 인공지능이 소모하는 이미지 토큰 절반 이하로 절감
                    raw_img = Image.open(BytesIO(file_bytes))
                    if raw_img.mode != 'RGB':
                        raw_img = raw_img.convert('RGB')
                    raw_img.thumbnail((900, 900), Image.Resampling.LANCZOS)
                    
                    compressed_buffer = BytesIO()
                    raw_img.save(compressed_buffer, format="JPEG", quality=75)
                    pil_image = Image.open(BytesIO(compressed_buffer.getvalue()))
                except Exception:
                    try:
                        pil_image = Image.open(BytesIO(file_bytes))
                    except Exception:
                        continue
                
                # [토큰 다이어트 2] 지시 프롬프트를 핵심 요약하여 전송 토큰 최소화
                prompt = """
                Extract English text from image.
                - Add '[HEADING]' before titles.
                - Add '[NAME]' before speaker names (e.g. [NAME] Tom:).
                - Delete all line numbers.
                - Keep paragraphs continuous.
                - Raw plain text only, no markdown.
                """

                status_text.text(f"🤖 [{idx+1}/{total_files}] 인공지능 교사가 지문을 깨끗하게 정렬하는 중입니다...")
                
                # 불필요한 재시도 횟수를 줄여 서버 한도 초과 원천 봉쇄
                max_retries = 2
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
                        if any(x in error_msg for x in ["429", "RESOURCE_EXHAUSTED", "QUOTA", "LIMIT_EXCEEDED"]):
                            quota_blocked = True
                            break 
                        
                        if "404" in error_msg or "NOT_FOUND" in error_msg:
                            api_error_occurred = True
                            break
                            
                        if attempt < max_retries - 1:
                            time.sleep(1.0)
                            continue
                        else:
                            break

                if quota_blocked or api_error_occurred:
                    break

                if extracted_text:
                    try:
                        success_count += 1
                        
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
                                        
                    except Exception:
                        continue
                
                for p in range(mid_p + 1, min(end_p, 99) + 1):
                    percent_display.markdown(f'<p class="percent-text">⏳ 문서 생성률: {p}%</p>', unsafe_allow_html=True)
                    progress_bar.progress(p / 100.0)
                    time.sleep(0.01)

            # 3. 결과 출력 및 완료 연출
            if success_count > 0:
                current_p = int(progress_bar.progress) if hasattr(progress_bar, 'progress') else 90
                for p in range(current_p, 101):
                    percent_display.markdown(f'<p class="percent-text">⏳ 문서 생성률: {p}%</p>', unsafe_allow_html=True)
                    progress_bar.progress(p / 100.0)
                    time.sleep(0.005)
                
                percent_display.markdown('<p class="percent-text" style="color:#0D9488;">🎉 변환 완료: 100%</p>', unsafe_allow_html=True)
                status_text.text("🎉 모든 영어 지문이 깨끗한 워드 파일 문서로 완성되었습니다!")
                
                if quota_blocked:
                    st.warning("⚠️ 오늘 준비된 무료 사용량이 도중에 소진되어, 판독에 성공한 지문들 위주로 우선 정렬되었습니다.")

                docx_buffer = BytesIO()
                doc.save(docx_buffer)
                docx_buffer.seek(0)
                
                st.download_button(
                    label="📥 변환된 Word 파일 다운로드",
                    data=docx_buffer,
                    file_name="Converted_English_Texts.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                
            elif quota_blocked and success_count == 0:
                percent_display.empty()
                progress_bar.empty()
                st.error("⚠️ 오늘 제공되는 구글 인공지능의 하루 무료 한도를 모두 소진했습니다. 내일 다시 시도하시거나 유료 API 계정 전환이 필요합니다.")
            elif api_error_occurred:
                percent_display.empty()
                progress_bar.empty()
                st.error("❌ 서버 환경 모델명 설정에 오류가 있습니다. 개발 환경 라이브러리 검토가 필요합니다.")
            elif not last_extracted_text: 
                percent_display.empty()
                progress_bar.empty()
                st.error("❌ 사진에서 영어 글자를 전혀 찾지 못했습니다. 사진이 흐리거나 어둡지 않은지 확인 후 다시 업로드해 주세요.")

except KeyError:
    st.error("🔒 설정 오류: 프로그램 관리자 설정(Streamlit Secrets)에 구글 인증키가 올바르게 등록되지 않았습니다.")
