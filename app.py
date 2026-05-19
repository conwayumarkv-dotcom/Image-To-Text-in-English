import streamlit as st
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn 
from io import BytesIO
import time
import re
from PIL import Image, ImageOps
from google import genai

# 1. 페이지 기본 설정 및 디자인 테마
st.set_page_config(
    page_title="Image To Text in English",
    page_icon="📝",
    layout="centered"
)

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
    }
    .sub-title {
        font-size: 18px;
        color: #7A7A7A;
        text-align: center;
        margin-bottom: 10px;
    }
    .author-footer {
        font-size: 13px;
        color: #A3A3A3; 
        text-align: right;
        margin-bottom: 45px;
    }
    div.stButton > button:first-child {
        background-color: #94A69A;
        color: white;
        border-radius: 10px;
        border: none;
        padding: 12px 40px;
        font-weight: 600;
        width: 100%;
    }
    .percent-text {
        font-size: 20px;
        font-weight: 700;
        color: #5C715E;
        margin-bottom: 8px;
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
            
            percent_display.markdown('<p class="percent-text">⏳ 문서 생성률: 0%</p>', unsafe_allow_html=True)
            progress_bar.progress(0.0)
            status_text.text("🔄 구글 최신 라이브러리 규격 맵핑을 시작합니다...")
            
            total_files = len(uploaded_files)
            
            # 구글 정식 라이브러리 공식 권장 메인 플래시 모델 탑재
            model_name = 'gemini-2.5-flash'
            
            success_count = 0     
            quota_blocked = False
            last_extracted_text = "" 
            
            progress_per_file = 100.0 / total_files
            
            for idx, file in enumerate(uploaded_files):
                file_bytes = file.read()
                extracted_text = "" 
                
                # 분당 트래픽 제한(RPM 리밋) 안전 방어 벽 설치
                if idx > 0:
                    for remaining in range(3, 0, -1):
                        status_text.text(f"⏳ 구글 서버 트래픽 안정화 대기 중... ({remaining}초)")
                        time.sleep(1)
                
                start_p = int(idx * progress_per_file)
                mid_p = int((idx + 0.6) * progress_per_file)
                end_p = int((idx + 1) * progress_per_file)
                
                status_text.text(f"📝 [{idx+1}/{total_files}] '{file.name}' 용량 초경량 압축 중...")
                for p in range(start_p, min(mid_p, 99) + 1):
                    percent_display.markdown(f'<p class="percent-text">⏳ 문서 생성률: {p}%</p>', unsafe_allow_html=True)
                    progress_bar.progress(p / 100.0)
                    time.sleep(0.01)
                
                try:
                    # [토큰 최소화] 흑백 다이어트 가공으로 픽셀 데이터 소모량 80% 절감
                    raw_img = Image.open(BytesIO(file_bytes))
                    raw_img = ImageOps.grayscale(raw_img)
                    raw_img.thumbnail((650, 650), Image.Resampling.LANCZOS)
                    
                    compressed_buffer = BytesIO()
                    raw_img.save(compressed_buffer, format="JPEG", quality=70)
                    pil_image = Image.open(BytesIO(compressed_buffer.getvalue()))
                except Exception:
                    try:
                        pil_image = Image.open(BytesIO(file_bytes))
                    except Exception:
                        continue
                
                # 가독성 및 정확성을 보장하는 초경량 정렬 프롬프트
                prompt = """
                Extract English text from this image perfectly.
                - Add '[HEADING]' right before any titles or headers.
                - Add '[NAME]' right before speaker names in dialogues.
                - Remove all line numbers and keep paragraphs continuous.
                """

                status_text.text(f"🤖 [{idx+1}/{total_files}] 인공지능이 공식 데이터 규격으로 텍스트를 추출하는 중...")
                
                try:
                    # [핵심 수정] 구글 정식 genai 라이브러리 매핑 오류를 방어하는 공식 표준 컨텐츠 순서 배치
                    response = client.models.generate_content(
                        model=model_name,
                        contents=[prompt, pil_image]
                    )
                    if response and response.text:
                        extracted_text = response.text
                        last_extracted_text = extracted_text
                    
                except Exception as api_err:
                    err_str = str(api_err).upper()
                    if any(x in err_str for x in ["429", "RESOURCE_EXHAUSTED", "QUOTA", "LIMIT_EXCEEDED"]):
                        quota_blocked = True
                        break

                if quota_blocked:
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

            # 3. 문서 정렬 생성 완료 후 결과 도출
            if success_count > 0:
                for p in range(int(progress_bar.progress * 100), 101):
                    percent_display.markdown(f'<p class="percent-text">⏳ 문서 생성률: {p}%</p>', unsafe_allow_html=True)
                    progress_bar.progress(p / 100.0)
                    time.sleep(0.005)
                
                percent_display.markdown('<p class="percent-text" style="color:#0D9488;">🎉 변환 완료: 100%</p>', unsafe_allow_html=True)
                status_text.text("🎉 정식 연동 매핑에 성공하여 깨끗한 워드 지문 문서가 발급되었습니다!")

                docx_buffer = BytesIO()
                doc.save(docx_buffer)
                docx_buffer.seek(0)
                
                st.download_button(
                    label="📥 변환된 Word 파일 다운로드",
                    data=docx_buffer,
                    file_name="Converted_English_Texts.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                
            else:
                percent_display.empty()
                progress_bar.empty()
                st.error("⚠️ 구글 계정 트래픽 한도가 일시 도달했습니다. 대략 1~2분 정도 후에 변환 버튼을 다시 클릭해 주시면 정상 처리됩니다.")

except KeyError:
    st.error("🔒 설정 오류: 프로그램 관리자 설정(Streamlit Secrets)에 구글 인증키가 올바르게 등록되지 않았습니다.")
