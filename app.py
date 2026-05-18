import streamlit as st
from docx import Document
from io import BytesIO
import os
from google import genai
from google.genai import types

# 1. 페이지 기본 설정 및 디자인 (세련된 테마 적용)
st.set_page_config(
    page_title="스마트 영어 교재 텍스트 추출기",
    page_icon="📝",
    layout="centered"
)

# 30대 원장님 취향의 깔끔하고 세련된 커스텀 CSS
st.markdown("""
    <style>
    .main-title {
        font-size: 28px;
        font-weight: 700;
        color: #1E3A8A; /* 차분한 딥네이비 */
        margin-bottom: 5px;
    }
    .sub-title {
        font-size: 14px;
        color: #6B7280; /* 그레이 */
        margin-bottom: 25px;
    }
    .status-box {
        padding: 15px;
        border-radius: 10px;
        background-color: #F3F4F6;
        margin-top: 15px;
    }
    /* 버튼 스타일 조정 */
    div.stButton > button:first-child {
        background-color: #0D9488; /* 세련된 딥민트/에메랄드 */
        color: white;
        border-radius: 8px;
        border: none;
        padding: 10px 20px;
        font-weight: 600;
    }
    div.stButton > button:first-child:hover {
        background-color: #0F766E;
    }
    </style>
""", unsafe_allow_html=True)

# UI 상단 타이틀
st.markdown('<p class="main-title">📝 Smart Text Extractor</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-title">영어 지문 사진을 업로드하면 깔끔한 Word 파일로 변환해 드립니다.</p>', unsafe_allow_html=True)

# 2. API 키 설정 (Streamlit Secrets 연동 또는 사이드바 입력)
# GitHub 배포 시에는 Streamlit 전용 Secrets 기능에 입력하는 것이 안전합니다.
api_key = st.sidebar.text_input("Gemini API Key를 입력하세요", type="password")

if not api_key:
    st.info("💡 왼쪽 사이드바에 Gemini API Key를 입력하시면 서비스가 활성화됩니다.")
else:
    # 클라이언트 초기화
    client = genai.Client(api_key=api_key)

    # 3. 파일 업로드 UI (여러 장 동시에 드래그 앤 드롭 가능)
    uploaded_files = st.file_uploader(
        "영어 지문 사진을 선택하거나 이 자리로 끌어다 놓으세요 (복수 선택 가능)", 
        type=["jpg", "jpeg", "png"],
        accept_multiple_files=True
    )

    if uploaded_files:
        st.write(f"📂 총 **{len(uploaded_files)}개**의 파일이 선택되었습니다.")
        
        # 변환 시작 버튼
        if st.button("Word 파일로 변환하기 ✨"):
            
            # 워드 문서 객체 생성
            doc = Document()
            doc.add_heading("Extract Text from Images", level=1)
            
            # 진행 상황 바
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            # 각 이미지 순회하며 텍스트 추출 및 워드에 추가
            for idx, file in enumerate(uploaded_files):
                status_text.text(f"⏳ [{idx+1}/{len(uploaded_files)}] '{file.name}' 분석 중...")
                
                # 이미지 데이터 읽기
                image_bytes = file.read()
                
                # Gemini 모델을 이용한 OCR 및 텍스트 정제 요청
                try:
                    response = client.models.generate_content(
                        model='gemini-2.5-flash',
                        contents=[
                            types.Part.from_bytes(
                                data=image_bytes,
                                mime_type=file.type,
                            ),
                            "이 사진 속의 영어 지문 텍스트를 한 글자도 빠짐없이 그대로 추출해줘. 제목이 있다면 구별해 주고, 단락 구분을 명확히 해줘. 영어 외에 다른 불필요한 설명은 포함하지 마."
                        ]
                    )
                    
                    extracted_text = response.text
                    
                    # 워드 문서에 파일명과 추출된 텍스트 추가
                    doc.add_heading(f"📄 Source: {file.name}", level=2)
                    doc.add_paragraph(extracted_text)
                    doc.add_page_break()  # 다음 사진은 다음 페이지에 나오도록 분할
                    
                except Exception as e:
                    st.error(f"❌ '{file.name}' 처리 중 오류 발생: {str(e)}")
                
                # 프로그레스 바 업데이트
                progress_bar.progress((idx + 1) / len(uploaded_files))
            
            status_text.text("✅ 모든 파일 변환 완료!")
            
            # 워드 파일을 메모리 상의 바이트 스트림으로 변환 (다운로드 버튼용)
            docx_buffer = BytesIO()
            doc.save(docx_buffer)
            docx_buffer.seek(0)
            
            # 4. 다운로드 UI 제공
            st.markdown('<div class="status-box">', unsafe_allow_html=True)
            st.success("🎉 Word 파일 작성이 완료되었습니다! 아래 버튼을 눌러 다운로드하세요.")
            st.download_button(
                label="📥 변환된 Word 파일 다운로드",
                data=docx_buffer,
                file_name="영어지문_모음.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            st.markdown('</div>', unsafe_allow_html=True)