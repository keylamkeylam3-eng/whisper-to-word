# ============================================
# MP3 語音轉文字 Word 產生器
# Streamlit + OpenAI Whisper
# ============================================

import io
import os
import tempfile

import streamlit as st
import whisper
from docx import Document
from docx.shared import Pt

# --------------------------------------------
# 頁面設定
# --------------------------------------------
st.set_page_config(
    page_title="MP3 語音轉文字 Word 產生器",
    page_icon="🎧",
    layout="centered"
)

st.title("🎧 MP3 語音轉文字 Word 產生器")
st.write(
    "上傳 MP3 檔案，系統會使用 Whisper 進行語音辨識，"
    "並產生帶有時間標記與標點符號的 Word 文件。"
)

# --------------------------------------------
# 建立暫存資料夾
# --------------------------------------------
TEMP_DIR = "temp_audio"
os.makedirs(TEMP_DIR, exist_ok=True)

# --------------------------------------------
# 選擇 Whisper 模型
# --------------------------------------------
MODEL_OPTIONS = {
    "base（較快，適合測試）": "base",
    "small（較準確，建議使用）": "small",
    "medium（更準確，較慢）": "medium"
}

selected_model = st.selectbox(
    "選擇辨識模型",
    options=list(MODEL_OPTIONS.keys()),
    index=1,
    key="model_selector"
)

MODEL_NAME = MODEL_OPTIONS[selected_model]

# --------------------------------------------
# 選擇語言
# --------------------------------------------
LANGUAGE_OPTIONS = {
    "自動偵測語言（建議）": None,
    "普通話／中文": "zh",
    "英文": "en"
}

selected_language = st.selectbox(
    "選擇語音語言",
    options=list(LANGUAGE_OPTIONS.keys()),
    index=0,
    key="language_selector"
)

LANGUAGE = LANGUAGE_OPTIONS[selected_language]

# --------------------------------------------
# 上傳音訊檔案
# --------------------------------------------
uploaded_file = st.file_uploader(
    "請上傳 MP3 或其他音訊檔案",
    type=["mp3", "wav", "m4a", "ogg", "flac"],
    key="audio_uploader"
)

# --------------------------------------------
# 按鈕：開始轉換
# --------------------------------------------
if uploaded_file is not None:

    st.success(f"已成功上傳檔案：{uploaded_file.name}")

    if st.button(
        "開始轉換為 Word 文件",
        type="primary",
        key="convert_button"
    ):
        try:
            # 顯示處理進度
            progress_bar = st.progress(0, text="正在準備音訊檔案……")

            # 儲存上傳檔案到暫存資料夾
            original_filename = uploaded_file.name
            file_extension = os.path.splitext(original_filename)[1]

            temp_audio_path = os.path.join(
                TEMP_DIR,
                f"uploaded_audio{file_extension}"
            )

            with open(temp_audio_path, "wb") as f:
                f.write(uploaded_file.getvalue())

            progress_bar.progress(
                10,
                text="正在載入 Whisper 模型……"
            )

            # 載入 Whisper 模型
            model = whisper.load_model(MODEL_NAME)

            progress_bar.progress(
                30,
                text="正在進行語音辨識，請稍候……"
            )

            # 執行語音辨識
            result = model.transcribe(
                temp_audio_path,
                language=LANGUAGE,
                word_timestamps=True,
                verbose=False
            )

            progress_bar.progress(
                80,
                text="正在建立 Word 文件……"
            )

            # 刪除暫存音訊檔
            if os.path.exists(temp_audio_path):
                os.remove(temp_audio_path)

            segments = result.get("segments", [])

            if not segments:
                progress_bar.empty()
                st.warning(
                    "未能從音訊檔辨識到文字，"
                    "請確認音訊內容是否清晰或重新上傳檔案。"
                )
                st.stop()

            # --------------------------------------------
            # 時間格式轉換
            # --------------------------------------------
            def format_time(seconds):
                hours = int(seconds // 3600)
                minutes = int((seconds % 3600) // 60)
                secs = int(seconds % 60)
                return f"{hours:02d}:{minutes:02d}:{secs:02d}"

            # --------------------------------------------
            # 建立 Word 文件
            # --------------------------------------------
            doc = Document()

            style = doc.styles["Normal"]
            style.font.name = "Microsoft JhengHei"
            style.font.size = Pt(12)

            doc.add_heading("語音轉文字紀錄", level=0)
            doc.add_paragraph(f"來源檔案：{original_filename}")
            doc.add_paragraph("")

            transcript_text = ""

            for seg in segments:
                start_time = format_time(seg["start"])
                end_time = format_time(seg["end"])
                text = seg["text"].strip()

                if text:
                    line = f"[{start_time} - {end_time}] {text}"

                    doc.add_paragraph(line)
                    transcript_text += line + "\n"

            # 將 Word 檔寫入記憶體
            docx_buffer = io.BytesIO()
            doc.save(docx_buffer)
            docx_buffer.seek(0)

            progress_bar.progress(
                100,
                text="轉換完成！"
            )

            st.success("語音轉文字完成！")

            # --------------------------------------------
            # 顯示辨識結果
            # --------------------------------------------
            st.subheader("辨識結果")

            st.text_area(
                "文字內容",
                value=transcript_text,
                height=350,
                key="transcript_area"
            )

            # --------------------------------------------
            # 下載 Word 檔
            # --------------------------------------------
            st.subheader("下載文件")

            base_filename = os.path.splitext(original_filename)[0]

            docx_filename = base_filename + "_語音轉文字.docx"
            txt_filename = base_filename + "_語音轉文字.txt"

            st.download_button(
                label="⬇️ 下載 Word 檔",
                data=docx_buffer,
                file_name=docx_filename,
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "wordprocessingml.document"
                ),
                use_container_width=True,
                key="download_docx_button"
            )

            st.download_button(
                label="⬇️ 下載文字檔（TXT）",
                data=transcript_text.encode("utf-8"),
                file_name=txt_filename,
                mime="text/plain",
                use_container_width=True,
                key="download_txt_button"
            )

        except Exception as e:
            st.error("轉換過程中發生錯誤。")
            st.exception(e)

else:
    st.info("請先上傳 MP3 或其他支援的音訊檔案。")

# --------------------------------------------
# 頁尾說明
# --------------------------------------------
st.divider()
st.caption(
    "提示：較長的錄音或較大的模型會需要較長處理時間。"
    "建議先使用 base 或 small 模型測試。"
)
    "提示：較長的錄音或較大的模型會需要較長處理時間。"
    "建議先使用 base 或 small 模型測試。"
)
