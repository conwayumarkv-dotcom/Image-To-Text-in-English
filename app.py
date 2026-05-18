import streamlit as st
from docx import Document
from docx.shared import Pt
from io import BytesIO
import time
import re
from google import genai
from google.genai import types
from google.genai.errors import APIError, ClientError, ServerError # 💡 503 에러 방어를 위해 ServerError 추가

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
                이 사진 속의 영어 지문 텍스트를 정확하게 추출해줘.
                - 사진의 메인 제목이나 소제목이 있다면 텍스트 제일 앞에 '[HEADING]' 이라는 태그를 붙여줘.
                - 대화 내용(Dialogue) 구조인 경우, 대화의 주체를 나타내는 이름 뒤에 콜론(:)을 붙이고, 이름 앞에 '[NAME]' 태그를 붙여줘. (예: [NAME] Mike: Hey, guys!)
                - 일반 본문 문장은 아무 태그 없이 문맥에 맞게 가독성 높은 단락으로 나누어줘.
                - 원본에서 **진하게** 처리된 강조 키워드가 있다면 마크다운 기호(**)를 유지해줘.
                - 결과물은 오직 추출된 텍스트만 보여주고, 다른 부연 설명은 하지 마.
                """
                
                status_text.text(f"⏳ [{idx+1}/{total_files}] '{file.name}' 지문을 AI 서버에서 분석하는 중...")
                
                extracted_text = ""
                max_retries = 3  # 503 이나 429 에러 시 최대 3번까지 끈질기게 다시 시도
                
                for attempt in range(max_retries):
                    try:
                        response = client.models.generate_content(
                            model=model_name,
                            contents=[types.Part.from_bytes(data=image_bytes, mime_type=file.type), prompt]
                        )
                        extracted_text = response.text
                        break  # 성공 시 재시도 루프 탈출
                        
                    except (APIError, ClientError, ServerError) as e:
                        # 429 한도초과 또는 503 서버 혼잡 에러가 발생한 경우 감지
                        error_str = str(e).upper()
                        
                        # 만약 진짜 하루치 용량(20장)이 완전히 다 끝난 거라면 루프 폭파 (낙하산 작동)
                        if "LIMIT: 20" in error_str:
                            quota_blocked = True
                            break
                            
                        # 그 외 일시적 트래픽 과부하(503, 일반 429)는 친절하게 대기 연출 작동
                        if attempt < max_retries - 1:
                            for remaining in range(8, 0, -1):
                                status_text.text(f"⏳ 처리하는데 시간이 걸리니 조금만 기다려주세요.. ({remaining}초)")
                                time.sleep(1)
                        else:
                            # 3번 다 실패한 경우에만 예외를 기록하고 다음 사진으로 이동
                            st.error(f"❌ '{file.name}' 구글 서버 일시 지연으로 실패 (나중에 다시 시도해 주세요)")
                            
                # 하루 제한 마감 낙하산이 펼쳐졌다면 파일 순회 즉시 탈출
                if quota_blocked:
                    st.warning("⚠️ 오늘 사용 가능한 구글 무료 한도(하루 20장)를 모두 소진했습니다. 현재까지 변환 성공한 지문들로만 워드 문서를 저장합니다.")
                    break

                # 텍스트 추출에 성공한 경우 워드 조립 및 게이지 처리
                if extracted_text:
                    success_count += 1
                    target_percent = int(((idx + 1) / total_files) * 100)
                    if idx == total_files - 1:
                        target_percent = 100
                    
                    for p in range(current_percent, target_percent + 1):
                        percent_display.markdown(f'<p class="percent-text">⏳ 변환 진행률: {p}%</p>', unsafe_allow_html=True)
                        progress_bar.progress(p)
                        time.sleep(0.01)
                    
                    current_percent = target_percent
                    status_text.text(f"✅ [{idx+1}/{total_files}] 정제 완료!")
                    
                    # 워드 문서 작성
                    doc.add_heading(f"Source: {file.name}", level=3)
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
                            run.font.size = Pt(14)
                            
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
                            parts = clean_text.split('**')
                            for i, part in enumerate(parts):
                                run = p_tag.add_run(part)
                                if i % 2 == 1:
                                    run.bold = True
                                    
                    doc.add_page_break()
                    
                    # 다음 파일 전 휴식 간격 타임라인 연동
                    if idx < total_files - 1:
                        steps = 60 
                        for step in range(steps):
                            sec_left = 6 - (step // 10)
                            status_text.text(f"⏳ 처리하는데 시간이 걸리니 조금만 기다려주세요.. ({sec_left}초)")
                            time.sleep(0.1)

            # 결과 다운로드 화면 출력
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
                    label=f"📥 변환된 Word 파일 다운로드 ({success_count}개 지문 포함)",
                    data=docx_buffer,
                    file_name="Converted_English_Texts.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                st.markdown('</div>', unsafe_allow_html=True)
            else:
                st.error("❌ 현재 구글 무료 서버의 접속량이 너무 많아 일시적으로 응답이 불가능합니다. 잠시 후 다시 시도해 주세요.")

except KeyError:
    st.error("🔒 설정 오류: Streamlit Secrets에 API Key를 등록해 주세요.")
