import streamlit as st
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn 
from io import BytesIO
import time
import re
from PIL import Image, ImageOps
from google import genai

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

try:
    # 1. API 키 로드 및 클라이언트 생성
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
            model_name = 'gemini-2.5-flash'
            success_count = 0     
            
            # 실제 구글 서버가 응답한 에러 메시지를 가감 없이 담아내기 위한 변수
            raw_error_message = None 
            
            progress_per_file = 100.0 / total_files
            
            for idx, file in enumerate(uploaded_files):
                file_bytes = file.read()
                extracted_text = "" 
                
                # 디레이팅 트래픽 가드: 무조건 3초 쉬고 호출
                if idx > 0:
                    for remaining in range(3, 0, -1):
                        status_text.text(f"⏳ 구글 서버 안정화 대기 중... ({remaining}초)")
                        time.sleep(1)
                
                start_p = int(idx * progress_per_file)
                status_text.text(f"📝 [{idx+1}/{total_files}] '{file.name}' 초정밀 이미지 전처리 중...")
                progress_bar.progress(start_p / 100.0)
                
                try:
                    # 이미지 600px 흑백 다이어트 가드 적용 (데이터 유실 최소화)
                    raw_img = Image.open(BytesIO(file_bytes))
                    raw_img = ImageOps.grayscale(raw_img)
                    raw_img.thumbnail((600, 600), Image.Resampling.LANCZOS)
                    
                    compressed_buffer = BytesIO()
                    raw_img.save(compressed_buffer, format="JPEG", quality=70)
                    pil_image = Image.open(BytesIO(compressed_buffer.getvalue()))
                except Exception:
                    pil_image = Image.open(BytesIO(file_bytes))
                
                prompt = """
                Extract English text from this image perfectly.
                - Add '[HEADING]' right before any titles or headers.
                - Add '[NAME]' right before speaker names in dialogues.
                - Remove all line numbers and keep paragraphs continuous.
                """

                status_text.text(f"🤖 [{idx+1}/{total_files}] 연동된 API 채널을 통해 지문 판독 중...")
                
                try:
                    response = client.models.generate_content(
                        model=model_name,
                        contents=[prompt, pil_image]
                    )
                    if response and response.text:
                        extracted_text = response.text
                except Exception as api_err:
                    # 필터링 없이 리얼 시스템 에러 원본 그대로 획득
                    raw_error_message = str(api_err)
                    break

                if raw_error_message:
                    break

                if extracted_text:
                    success_count += 1
                    doc.add_paragraph(f"▪ Source: {file.name}").runs[0].font.color.rgb = RGBColor(128, 128, 128)
                    
                    for para_text in extracted_text.split('\n'):
                        clean_text = para_text.strip()
                        if not clean_text:
                            continue
                        
                        p_tag = doc.add_paragraph()
                        if clean_text.startswith("[HEADING]"):
                            heading_content = clean_text.replace("[HEADING]", "").strip()
                            run = p_tag.add_run(heading_content)
                            run.bold = True
                            run.font.size = Pt(13)
                        elif clean_text.startswith("[NAME]"):
                            name_content = clean_text.replace("[NAME]", "").strip()
                            match = re.match(r"^([^:]+:)(.*)$", name_content)
                            if match:
                                r_name = p_tag.add_run(match.group(1))
                                r_name.bold = True
                                p_tag.add_run(match.group(2))
                            else:
                                p_tag.add_run(name_content)
                        else:
                            p_tag.add_run(clean_text.replace("**", ""))
                    
                    if idx < total_files - 1:
                        doc.add_page_break()
                
                progress_bar.progress(int((idx + 1) * progress_per_file) / 100.0)

            # 결과 처리 바인딩
            if success_count > 0:
                percent_display.markdown('<p class="percent-text" style="color:#0D9488;">🎉 변환 완료: 100%</p>', unsafe_allow_html=True)
                status_text.text("🎉 새 API 연동 채널을 통해 문서 전송 및 정렬 작업이 정상 완료되었습니다!")

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
                # 고정 경고를 지우고 실제 시스템 에러 로그를 명확히 바인딩
                st.error(f"❌ 구글 통신 장애 발생! 구글 서버가 반환한 리얼 메시지입니다:\n\n`{raw_error_message}`")

except KeyError:
    st.error("🔒 설정 오류: Secrets 관리자 화면에 GEMINI_API_KEY 설정이 되어있는지 다시 확인해 주세요.")
