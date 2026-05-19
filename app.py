import streamlit as st
from docx import Document
from docx.shared import Pt, RGBColor
from io import BytesIO
import time
import re
from google import genai
from google.genai import types
from google.genai.errors import APIError, ClientError, ServerError

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
        font-size: 14px;
        color: #B2B2B2; 
        text-align: right;
        margin-bottom: 45px;
        padding-right: 5px;
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
st.markdown('<div class="author-footer">(Made by Manju)</div>', unsafe_allow_html=True)

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
            
            current_percent = 0
            success_count = 0     
            quota_blocked = False 
            
            for idx, file in enumerate(uploaded_files):
                image_bytes = file.read()
                
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
                
                status_text.text(f"⏳ [{idx+1}/{total_files}] '{file.name}' 지문을 AI 서버에서 분석하는 중...")
                
                # ---------------------------------------------------------
                # 🛠️ [핵심 개선] 부드러운 진행률 상승을 위한 가상 애니메이션 루프 적용
                # ---------------------------------------------------------
                # 이번 파일 작업 완료 시 도달해야 할 실제 최종 목표 %
                real_target_percent = int(((idx + 1) / total_files) * 100)
                if idx == total_files - 1:
                    real_target_percent = 100

                # AI 호출 직전, 다음 목표치 직전(-2%)까지 부드럽게 먼저 채우는 가상 루프
                # 예: 0% -> 98%까지 약 4~5초간 일정하게 올라가며 사용자의 지루함을 방지합니다.
                virtual_target = max(current_percent, real_target_percent - 2)
                
                extracted_text = ""
                max_retries = 3  
                
                for attempt in range(max_retries):
                    try:
                        # 1단계: 가상으로 70% 정도까지 아주 빠르게 먼저 도달하게 유도 (체감 속도 업)
                        mid_target = current_percent + int((virtual_target - current_percent) * 0.7)
                        for p in range(current_percent, mid_target + 1):
                            percent_display.markdown(f'<p class="percent-text">⏳ 변환 진행률: {p}%</p>', unsafe_allow_html=True)
                            progress_bar.progress(p)
                            time.sleep(0.01)
                        
                        current_percent = mid_target
                        
                        # 2단계: 나머지 98% 지점까지는 AI 연산 속도(약 3~4초)에 맞춰 초당 약 10~15%씩 서서히 증가
                        # AI 서버 응답이 이 루프 중간이나 끝난 직후에 떨어지도록 완충 작용을 합니다.
                        response = None
                        
                        # 가상으로 조금씩 전진하는 비동기식 연출 루프 효과
                        for p in range(current_percent, virtual_target + 1):
                            percent_display.markdown(f'<p class="percent-text">⏳ 변환 진행률: {p}%</p>', unsafe_allow_html=True)
                            progress_bar.progress(p)
                            
                            # 루프 극초반에 딱 한 번 API 실제 요청을 보냄 (Non-blocking 느낌으로 연출)
                            if response is None:
                                response = client.models.generate_content(
                                    model=model_name,
                                    contents=[types.Part.from_bytes(data=image_bytes, mime_type=file.type), prompt]
                                )
                            
                            # API 응답을 이미 받았으므로 루프 가속을 위해 아주 짧은 딜레이만 할당
                            time.sleep(0.04) 
                        
                        extracted_text = response.text
                        current_percent = virtual_target
                        break  
                        
                    except (APIError, ClientError, ServerError) as e:
                        error_str = str(e).upper()
                        if "LIMIT: 20" in error_str:
                            quota_blocked = True
                            break
                            
                        if attempt < max_retries - 1:
                            for remaining in range(8, 0, -1):
                                status_text.text(f"⏳ 처리하는데 시간이 걸리니 조금만 기다려주세요.. ({remaining}초)")
                                time.sleep(1)
                        else:
                            st.error(f"❌ '{file.name}' 구글 서버 일시 지연으로 실패")
                            
                if quota_blocked:
                    st.warning("⚠️ 오늘 사용 가능한 구글 무료 한도(하루 20장)를 모두 소진했습니다. 현재까지 변환 성공한 지문들로만 워드 문서를 저장합니다.")
                    break

                if extracted_text:
                    try:
                        success_count += 1
                        
                        # 3단계: AI 데이터 처리가 끝나면 목표했던 진짜 %까지 완벽하게 도달시킴
                        for p in range(current_percent, real_target_percent + 1):
                            percent_display.markdown(f'<p class="percent-text">⏳ 변환 진행률: {p}%</p>', unsafe_allow_html=True)
                            progress_bar.progress(p)
                            time.sleep(0.01)
                        
                        current_percent = real_target_percent
                        status_text.text(f"✅ [{idx+1}/{total_files}] 정제 완료!")
                        
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
                        st.warning(f"⚠️ '{file.name}' 문서 서식 스타일링 적용 중 경미한 지연이 있습니다. 텍스트 추출은 유지됩니다.")
                        continue
                    
                    # 💡 파일 간 간격 딜레이 리팩토링 (기존 6초 대기를 부드러운 스태거 방식으로 통합)
                    if idx < total_files - 1:
                        steps = 30 
                        for step in range(steps):
                            sec_left = 3 - (step // 10)
                            status_text.text(f"⏳ 다음 지문으로 이동 중입니다.. ({sec_left}초)")
                            time.sleep(0.1)

            if success_count > 0:
                if not quota_blocked:
                    percent_display.markdown('<p class="percent-text" style="color:#0D9488;">🎉 변환 진행률: 100%</p>', unsafe_allow_html=True)
                    progress_bar.progress(100)
                    status_text.text("🎉 모든 영어 지문이 성공적으로 변환되었습니다!")
                else:
                    status_text.text(f"⚠️ 구글 무료 한도로 인해 {success_count}개의 지문만 먼저 변환 완료되었습니다.")
                
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
                st.error("❌ 현재 구글 무료 서버의 접속량이 너무 많아 일시적으로 응답이 불가능합니다. 잠시 후 다시 시도해 주세요.")

except KeyError:
    st.error("🔒 설정 오류: Streamlit Secrets에 API Key를 등록해 주세요.")
