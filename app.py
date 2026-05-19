import streamlit as st
from docx import Document
from docx.shared import Pt, RGBColor
from io import BytesIO
import time
import re
import json
import threading
from PIL import Image
from google import genai
from google.genai import types
from google.genai.errors import APIError, ClientError, ServerError

# 1. 페이지 기본 설정 및 디자인
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

# 🛠️ [구조 혁신] 단어장 프로그램처럼 JSON 데이터를 안전하게 받아오는 워커 함수
def gemini_api_worker(client, model_name, pil_image, prompt, result_container):
    max_retries = 3
    for attempt in range(max_retries):
        try:
            response = client.models.generate_content(
                model=model_name,
                contents=[pil_image, prompt],
                config=types.GenerateContentConfig(
                    response_mime_type="application/json" # 🟢 단어장과 동일한 구조적 마이그레이션
                )
            )
            # 결과물을 안전하게 JSON파싱하여 저장
            result_container["json_data"] = json.loads(response.text)
            result_container["status"] = "success"
            return
        except Exception as e:
            error_str = str(e).upper()
            if ("LIMIT" in error_str or "QUOTA" in error_str or "429" in error_str or "EXHAUSTED" in error_str):
                result_container["status"] = "quota_error"
                return
            if attempt < max_retries - 1:
                time.sleep(1.5)
                continue
            result_container["status"] = "fail"
            return

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
            model_name = 'gemini-2.5-flash' # 다시 가볍고 빠른 flash 모델로 안정적 구동 가능
            
            current_percent = 0
            success_count = 0     
            quota_blocked = False 
            api_failed_completely = False 
            
            for idx, file in enumerate(uploaded_files):
                file_bytes = file.read()
                try:
                    pil_image = Image.open(BytesIO(file_bytes))
                except Exception:
                    st.error(f"❌ '{file.name}' 이미지를 읽어오는 중에 오류가 발생했습니다.")
                    continue
                
                # 🛠️ AI 오버헤드를 제로로 만드는 초경량 JSON 유도 프롬프트
                prompt = """
                Extract the English text from this image and structure it into the following JSON format.
                - Ignore textbook line numbers like '5', '10', '15'.
                - Merge split lines into a single, cohesive paragraph. Only start a new paragraph when the story/context naturally changes.
                - Do NOT include any markdown formatting like '**'.
                
                Respond ONLY with a JSON object inside this structure:
                {
                  "type": "heading" or "dialogue" or "normal",
                  "content": [
                    {
                      "role": "If type is dialogue, put speaker name here (e.g., Mike). Otherwise leave it empty.",
                      "text": "The extracted sentence or paragraph text here."
                    }
                  ]
                }
                """
                
                status_text.text(f"⏳ [{idx+1}/{total_files}] '{file.name}' 사진 속 영어 지문을 깨끗하게 읽어오는 중입니다...")
                
                real_target_percent = int(((idx + 1) / total_files) * 100)
                if idx == total_files - 1:
                    real_target_percent = 100

                virtual_target = max(current_percent, real_target_percent - 3)
                
                worker_result = {"status": "pending", "json_data": None}
                
                api_thread = threading.Thread(
                    target=gemini_api_worker,
                    args=(client, model_name, pil_image, prompt, worker_result)
                )
                api_thread.start()
                
                ui_progress = float(current_percent)
                while api_thread.is_alive():
                    if ui_progress < float(virtual_target):
                        ui_progress += 0.5
                        if ui_progress > float(virtual_target):
                            ui_progress = float(virtual_target)
                        percent_display.markdown(f'<p class="percent-text">⏳ 변환 진행률: {int(ui_progress)}%</p>', unsafe_allow_html=True)
                        progress_bar.progress(int(ui_progress))
                    time.sleep(0.04)
                
                if worker_result["status"] == "quota_error":
                    quota_blocked = True
                elif worker_result["status"] == "fail":
                    api_failed_completely = True
                
                if quota_blocked:
                    if success_count > 0:
                        st.warning("⚠️ 하루 이용 한도를 모두 소진했습니다. 현재까지 성공한 지문들로만 워드 문서를 저장합니다.")
                    break
                    
                if api_failed_completely:
                    break

                if worker_result["status"] == "success" and worker_result["json_data"]:
                    json_data = worker_result["json_data"]
                    try:
                        success_count += 1
                        
                        for p in range(int(ui_progress), real_target_percent + 1):
                            percent_display.markdown(f'<p class="percent-text">⏳ 변환 진행률: {p}%</p>', unsafe_allow_html=True)
                            progress_bar.progress(p)
                            time.sleep(0.01)
                        
                        current_percent = real_target_percent
                        status_text.text(f"✅ [{idx+1}/{total_files}] 지문 변환 및 서식 정리 완료!")
                        
                        # 사진 출처 표기 서식
                        p_src = doc.add_paragraph()
                        r_src = p_src.add_run(f"▪ Source: {file.name}")
                        r_src.font.size = Pt(10)
                        r_src.font.color.rgb = RGBColor(128, 128, 128)  
                        
                        # 🛠️ [파이썬 후처리] 깨끗하게 받아온 JSON 데이터를 기반으로 워드 파일 스타일링 조립
                        text_type = json_data.get("type", "normal")
                        contents = json_data.get("content", [])
                        
                        for item in contents:
                            p_tag = doc.add_paragraph()
                            item_text = item.get("text", "").strip()
                            if not item_text:
                                continue
                                
                            if text_type == "heading":
                                run = p_tag.add_run(item_text)
                                run.bold = True
                                run.font.size = Pt(13) 
                                p_tag.paragraph_format.space_before = Pt(12) 
                                p_tag.paragraph_format.space_after = Pt(6)
                                
                            elif text_type == "dialogue":
                                role = item.get("role", "").strip()
                                if role:
                                    r_name = p_tag.add_run(f"{role}: ")
                                    r_name.bold = True
                                p_tag.add_run(item_text)
                                
                            else:
                                p_tag.add_run(item_text)
                                        
                        doc.add_page_break()
                        
                    except Exception:
                        st.warning(f"⚠️ '{file.name}' 서식을 다듬는 과정에서 경미한 지연이 있습니다.")
                        continue
                    
                    if idx < total_files - 1:
                        time.sleep(0.2)

            if success_count > 0:
                if not quota_blocked:
                    percent_display.markdown('<p class="percent-text" style="color:#0D9488;">🎉 변환 진행률: 100%</p>', unsafe_allow_html=True)
                    progress_bar.progress(100)
                    status_text.text("🎉 선택하신 모든 영어 지문이 워드 파일로 멋지게 변환되었습니다!")
                
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
                if quota_blocked:
                    status_text.empty()
                    st.error("⚠️ 오늘 제공되는 구글의 무료 변환 한도가 모두 소진되어 지금은 변환을 시작할 수 없습니다. 내일 다시 이용해 주세요.")
                else:
                    status_text.empty()
                    st.error("❌ 구글 서버 연결이 일시적으로 원활하지 않습니다. 잠시 후 다시 시도해 주세요.")

except KeyError:
    st.error("🔒 설정 오류: Streamlit Secrets에 API Key를 등록해 주세요.")
