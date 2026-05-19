import streamlit as st
from docx import Document
from docx.shared import Pt, RGBColor
from io import BytesIO
import time
import re  # 🛠️ NameError 해결을 위해 반드시 필요한 정규표현식 라이브러리
from PIL import Image  # 🛠️ 고용량 이미지 압축 및 전처리를 위한 라이브러리
from google import genai
from google.genai import types

# 1. 페이지 기본 설정 및 디자인
st.set_page_config(
    page_title="Image To Text in English",
    page_icon="📝",
    layout="centered"
)

# 세련되고 직관적인 커스텀 CSS (크림/그린 테마 정자체 세팅)
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

# UI 상단 타이틀 및 설명 구조 정제
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
            style.font.size = Pt(11)
            
            percent_display = st.empty()  
            progress_bar = st.progress(0)  
            status_text = st.empty()       
            
            total_files = len(uploaded_files)
            model_name = 'gemini-2.5-flash'
            success_count = 0     
            
            for idx, file in enumerate(uploaded_files):
                file_bytes = file.read()
                
                # 🛠️ 고용량 사진 데이터 해상도 다이어트 및 포맷 최적화 (토큰 한도 초과 방지)
                try:
                    raw_img = Image.open(BytesIO(file_bytes))
                    if raw_img.mode != 'RGB':
                        raw_img = raw_img.convert('RGB')
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
                
                prompt = """
                이 사진 속의 영어 지문 텍스트를 상식적이고 가독성 높은 문맥에 맞춰 추출해줘.
                - 사진의 메인 제목이나 큰 단원 소제목이 있다면 텍스트 제일 앞에 '[HEADING]' 이라는 태그를 붙여줘. (예: [HEADING] Into a New World of Storytelling)
                - 대화 내용 구조인 경우에만, 대화 주체 이름 뒤에 콜론(:)을 붙이고 이름 앞에 '[NAME]' 태그를 붙여줘. (예: [NAME] Mike: Hey, guys!)
                
                [⚠️ 중요 규칙: 가독성 및 줄바꿈 지시]
                - 영어 원문 텍스트 중간이나 우측 가장자리에 적혀 있는 '5', '10', '15'와 같은 행 번호(Line Numbers) 표시는 지문 본문 단어와 꼬이지 않도록 완벽하게 무시하고 걷어내줘.
                - 원본 사진의 줄바꿈을 억지로 따라 하지 마. 하나의 긴 단락(Paragraph)은 인위적으로 끊지 말고 쭉 이어서 하나의 유기적인 문단으로 합쳐서 완성해줘. 
                - 오직 문맥상 새로운 이야기나 단락이 시작될 때만 자연스럽게 줄바꿈을 적용해줘.
                - 본문 단어에 임의로 마크다운 별표(**)나 진하게 설정을 넣지 마. 오직 순수한 텍스트만 출력해줘.
                - 결과물은 오직 추출된 텍스트만 보여주고, 다른 부연 설명은 하지 마.
                """
                
                status_text.text(f"⏳ [{idx+1}/{total_files}] '{file.name}' 사진 속 영어 지문을 깨끗하게 읽어오는 중입니다...")
                
                # 진행률 계산 및 부드러운 전진 효과
                current_percent = int((idx / total_files) * 100)
                target_percent = int(((idx + 1) / total_files) * 100)
                
                for p in range(current_percent, min(target_percent, current_percent + 5)):
                    percent_display.markdown(f'<p class="percent-text">⏳ 변환 진행률: {p}%</p>', unsafe_allow_html=True)
                    progress_bar.progress(p)
                    time.sleep(0.02)

                # API 호출 수행
                try:
                    response = client.models.generate_content(
                        model=model_name,
                        contents=[pil_image, prompt]
                    )
                    extracted_text = response.text
                except Exception as api_err:
                    st.error(f"❌ '{file.name}' 구글 서버 연동 중 오류가 발생했습니다: {str(api_err)}")
                    continue

                if extracted_text:
                    try:
                        success_count += 1
                        
                        # 파일 분석 완료 후 해당 구간 채우기
                        for p in range(min(target_percent, current_percent + 5), target_percent + 1):
                            percent_display.markdown(f'<p class="percent-text">⏳ 변환 진행률: {p}%</p>', unsafe_allow_html=True)
                            progress_bar.progress(p)
                            time.sleep(0.01)
                        
                        status_text.text(f"✅ [{idx+1}/{total_files}] 지문 변환 및 서식 정리 완료!")
                        
                        # 사진 출처 표기 서식
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
                            
                            # 스타일링 1: 소제목/제목 구조 처리
                            if clean_text.startswith("[HEADING]"):
                                heading_content = clean_text.replace("[HEADING]", "").strip()
                                run = p_tag.add_run(heading_content)
                                run.bold = True
                                run.font.size = Pt(13) 
                                p_tag.paragraph_format.space_before = Pt(12) 
                                p_tag.paragraph_format.space_after = Pt(6)   
                                
                            # 스타일링 2: 대화문 주체 구조 처리
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
                                    
                            # 스타일링 3: 일반 본문 문장 구조 처리
                            else:
                                plain_content = clean_text.replace("**", "")
                                p_tag.add_run(plain_content)
                                        
                        doc.add_page_break()
                        
                    except Exception as word_err:
                        st.warning(f"⚠️ '{file.name}' 문서 디자인 조립 중 경미한 지연이 생겨 안전하게 다음으로 패스합니다.")
                        continue
                    
                    # 연속 요청 간격을 주어 트래픽 분산
                    if idx < total_files - 1:
                        time.sleep(1.0)

            if success_count > 0:
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
            else:
                status_text.empty()
                st.error("❌ 변환된 지문이 없습니다. 파일 상태나 API Key 설정을 확인해 주세요.")

except KeyError:
    st.error("🔒 설정 오류: Streamlit Secrets에 API Key를 등록해 주세요.")
