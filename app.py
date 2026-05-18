import streamlit as st
from docx import Document
from docx.shared import Pt
from io import BytesIO
import time
import re
from google import genai
from google.genai import types
from google.genai.errors import APIError, ClientError

# 1. 페이지 기본 설정 및 디자인
st.set_page_config(
    page_title="Image To Text in English",
    page_icon="📝",
    layout="centered"
)

# 세련되고 직관적인 커스텀 CSS (기울임꼴 제거 반영)
st.markdown("""
    <style>
    /* 전체 배경색: 포근하고 부드러운 크림톤 */
    .stApp { 
        background-color: #FDFBF6; 
    }
    
    /* 메인 타이틀: 한 줄로 깔끔하게 배치 및 폰트 최적화 */
    .main-title {
        font-size: 42px !important; 
        font-weight: 800;
        color: #5C715E; /* 차분한 뮤티드 그린 */
        text-align: center;
        margin-bottom: 8px;
        letter-spacing: -1.5px;
        white-space: nowrap; /* 무조건 한 줄 유지 */
    }
    
    /* 서브 타이틀: 직관적인 기능 중심 설명 */
    .sub-title {
        font-size: 18px;
        color: #7A7A7A;
        text-align: center;
        margin-bottom: 10px;
        font-weight: 400;
        letter-spacing: -0.5px;
    }
    
    /* 💡 [수정 완료] (Made by Manju) 우측 하단 서명 스타일 */
    .author-footer {
        font-size: 14px;
        color: #B2B2B2; /* 은은한 밝은 그레이 */
        text-align: right;
        margin-bottom: 45px;
        /* 기울임꼴(italic) 속성을 제거했습니다 */
        padding-right: 5px;
    }
    
    /* 파일 업로드 뷰어 영역 디자인 */
    .stFileUploader {
        border-radius: 15px !important;
        background-color: white;
        box-shadow: 0 4px 15px rgba(0,0,0,0.05);
        padding: 10px;
    }
    
    /* 변환 버튼 스타일 (가로 꽉 차게 변경하여 시원한 느낌) */
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
    div.stButton > button:first-child:hover { 
        background-color: #7E8F83; 
    }
    
    /* 결과 및 다운로드 박스 디자인 */
    .status-box {
        padding: 25px;
        border-radius: 20px;
        background-color: white;
        margin-top: 30px;
        box-shadow: 0 4px 20px rgba(0,0,0,0.08);
        text-align: center;
    }
    
    /* 실시간 퍼센트 숫자 강조 스타일 */
    .percent-text {
        font-size: 20px;
        font-weight: 700;
        color: #5C715E;
        margin-bottom: 8px;
        text-align: left;
    }
    </style>
""", unsafe_allow_html=True)

# UI 상단 타이틀 레이아웃 구성
st.markdown('<p class="main-title">Image To Text in English</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-title">사진 속 지문을 인식하여 편집 가능한 워드 문서(.docx)로 변환합니다.</p>', unsafe_allow_html=True)
# 💡 정자체로 표시되도록 스타일이 수정되었습니다
st.markdown('<div class="author-footer">(Made by Manju)</div>', unsafe_allow_html=True)

# 2. API Key 자동 로드 섹션
try:
    api_key = st.secrets["GEMINI_API_KEY"]
    client = genai.Client(api_key=api_key)
    
    # 3. 파일 업로드 UI
    uploaded_files = st.file_uploader(
        "변환할 영어 지문 사진을 업로드하세요 (복수 선택 가능)", 
        type=["jpg", "jpeg", "png"],
        accept_multiple_files=True
    )

    if uploaded_files:
        st.write(f"📂 **{len(uploaded_files)}개**의 파일이 선택되었습니다.")
        
        if st.button("Word 파일로 변환하기 ✨"):
            
            # 기본 워드 문서 생성 및 스타일 초기화
            doc = Document()
            style = doc.styles['Normal']
            style.font.name = 'Arial'
            style.font.size = Pt(11)
            
            # 실시간 상태창 컨테이너 빌딩
            percent_display = st.empty()  
            progress_bar = st.progress(0)  
            status_text = st.empty()       
            
            total_files = len(uploaded_files)
            model_name = 'gemini-2.5-flash' # 이미지 인식 효율이 가장 뛰어난 모델 적용
            
            current_percent = 0
            success_count = 0     # 실 변환 성공 카운트
            quota_blocked = False # 무료 계정 하루 제한 도달 감지 플래그
            
            for idx, file in enumerate(uploaded_files):
                image_bytes = file.read()
                
                # 가독성 및 세부 구조 조정을 위한 AI 커스텀 프롬프트
                prompt = """
                이 사진 속의 영어 지문 텍스트를 정확하게 추출해줘.
                - 사진의 메인 제목이나 소제목이 있다면 텍스트 제일 앞에 '[HEADING]' 이라는 태그를 붙여줘. (예: [HEADING] Reports from the Battlefield)
                - 대화 내용(Dialogue) 구조인 경우, 대화의 주체를 나타내는 이름 뒤에 콜론(:)을 붙이고, 이름 앞에 '[NAME]' 태그를 붙여줘. (예: [NAME] Mike: Hey, guys!)
                - 일반 본문 문장은 아무 태그 없이 문맥에 맞게 가독성 높은 단락으로 나누어줘.
                - 원본에서 **진하게** 처리된 강조 키워드가 있다면 마크다운 기호(**)를 유지해줘.
                - 결과물은 오직 추출된 텍스트만 보여주고, 다른 부연 설명은 하지 마.
                """
                
                status_text.text(f"⏳ [{idx+1}/{total_files}] '{file.name}' 지문을 AI 서버에서 분석하는 중...")
                
                extracted_text = ""
                try:
                    response = client.models.generate_content(
                        model=model_name,
                        contents=[types.Part.from_bytes(data=image_bytes, mime_type=file.type), prompt]
                    )
                    extracted_text = response.text
                    
                except (APIError, ClientError) as e:
                    error_msg = str(e).upper()
                    # 무료 제공 한도 초과 에러 감지 시의 방어 기전
                    if "QUOTA EXCEEDED" in error_msg or "RESOURCE_EXHAUSTED" in error_msg:
                        # 6초간 부드럽게 카운트다운을 표시하며 대기
                        for remaining in range(6, 0, -1):
                            status_text.text(f"⏳ 처리하는데 시간이 걸리니 조금만 기다려주세요.. ({remaining}초)")
                            time.sleep(1)
                        
                        try:
                            # 대기 후 최종 재시도
                            response = client.models.generate_content(
                                model=model_name,
                                contents=[types.Part.from_bytes(data=image_bytes, mime_type=file.type), prompt]
                            )
                            extracted_text = response.text
                        except Exception:
                            # 재시도마저 거부당한 경우 (완전 하루 마감) -> 낙하산 작동
                            quota_blocked = True
                            st.warning("⚠️ 오늘 사용 가능한 구글 무료 한도(하루 20장)를 모두 소진했습니다. 현재까지 변환 성공한 지문들로만 워드 문서를 저장합니다.")
                            break
                    else:
                        st.error(f"❌ '{file.name}' 처리 중 오류 발생: {str(e)}")
                        continue
                except Exception as general_err:
                    st.error(f"❌ '{file.name}' 시스템 예외 오류: {str(general_err)}")
                    continue

                # 텍스트 데이터 획득 성공 시 실시간 바인딩 작업 진행
                if extracted_text:
                    success_count += 1
                    target_percent = int(((idx + 1) / total_files) * 100)
                    if idx == total_files - 1:
                        target_percent = 100
                    
                    # 💡 멈춤/점프 현상 완벽 방지: 텍스트 추출 완료 후 다음 퍼센트까지 1%씩 부드럽게 흐름 연동
                    for p in range(current_percent, target_percent + 1):
                        percent_display.markdown(f'<p class="percent-text">⏳ 변환 진행률: {p}%</p>', unsafe_allow_html=True)
                        progress_bar.progress(p)
                        time.sleep(0.01)
                    
                    current_percent = target_percent
                    status_text.text(f"✅ [{idx+1}/{total_files}] 분석 및 데이터 정제 완료!")
                    
                    # 워드 문서 빌딩
                    doc.add_heading(f"Source: {file.name}", level=3)
                    
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
                            run.font.size = Pt(14)
                            
                        # 스타일링 2: 대화문 주체(이름만 진하게) 구조 처리
                        elif clean_text.startswith("[NAME]"):
                            name_content = clean_text.replace("[NAME]", "").strip()
                            # 정규식을 이용해 '이름:' 파트만 정확하게 슬라이싱
                            match = re.match(r"^([^:]+:)(.*)$", name_content)
                            if match:
                                name_part = match.group(1)   
                                dialogue_part = match.group(2) 
                                
                                # 이름 부분만 볼드(진하게) 처리
                                r_name = p_tag.add_run(name_part)
                                r_name.bold = True
                                
                                # 대화 본문은 기본 폰트 처리
                                p_tag.add_run(dialogue_part)
                            else:
                                p_tag.add_run(name_content)
                                
                        # 스타일링 3: 일반 본문 문장 구조 처리
                        else:
                            parts = clean_text.split('**')
                            for i, part in enumerate(parts):
                                run = p_tag.add_run(part)
                                if i % 2 == 1:
                                    run.bold = True
                                    
                    doc.add_page_break() # 깔끔하게 단원별 페이지 분할
                    
                    # 💡 6초 안전 휴식 타이머 구간 동안에도 퍼센트를 초 단위 타임라인과 완벽 연동
                    if idx < total_files - 1:
                        steps = 60 
                        # 6초간 퍼센트를 아주 조금씩 올려 안전함을 연출
                        percent_increment = (target_percent - current_percent) * 0.05 / steps
                        
                        for step in range(steps):
                            sec_left = 6 - (step // 10)
                            status_text.text(f"⏳ 처리하는데 시간이 걸리니 조금만 기다려주세요.. ({sec_left}초)")
                            time.sleep(0.1)

            # 4. 최종 마감 다운로드 팝업 노출 조건문
            if success_count > 0:
                if not quota_blocked:
                    percent_display.markdown('<p class="percent-text" style="color:#0D9488;">🎉 변환 진행률: 100%</p>', unsafe_allow_html=True)
                    progress_bar.progress(100)
                    status_text.text("🎉 모든 영어 지문이 성공적으로 변환되었습니다!")
                else:
                    status_text.text(f"⚠️ 구글 한도 도달로 인해 총 {total_files}개 중 {success_count}개만 변환되었습니다.")
                
                # 워드 데이터 바이트 스트림 변환
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
                st.error("❌ 오늘 사용 가능한 구글 무료 제공량(하루 20장)을 모두 초과하여 변환을 시작할 수 없습니다. 내일 다시 시도해 주세요.")

except KeyError:
    st.error("🔒 설정 오류: Streamlit Secrets에 API Key를 등록해 주세요.")
