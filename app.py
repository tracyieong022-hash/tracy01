import streamlit as st
import whisper
from docx import Document
import tempfile
import os
from io import BytesIO

# 頁面設定
st.set_page_config(
    page_title="語音轉文字 → Word",
    page_icon="🎙️",
    layout="centered"
)

st.title("🎙️ 語音轉文字 → 產生 Word")
st.markdown("上傳 MP3（或其他音訊），自動轉成文字並下載 Word 檔。")

# 側邊欄設定
with st.sidebar:
    st.header("設定")
    model_size = st.selectbox(
        "Whisper 模型",
        options=["tiny", "base", "small"],
        index=1,
        help="tiny 最快、準確度較低；base 平衡；small 較準但更慢"
    )
    language = st.selectbox(
        "語言",
        options=["自動偵測", "中文 (zh)", "英文 (en)"],
        index=0
    )
    st.markdown("---")
    st.caption("注意：Streamlit Cloud 為 CPU 環境，長音檔或較大模型可能超時。")

# 上傳檔案
uploaded_file = st.file_uploader(
    "上傳音訊檔（建議 MP3，大小建議 < 25MB）",
    type=["mp3", "wav", "m4a", "ogg", "flac"]
)

# 載入模型（使用 cache，避免每次都重新載入）
@st.cache_resource
def load_whisper_model(size: str):
    return whisper.load_model(size)

if uploaded_file is not None:
    st.audio(uploaded_file, format=f"audio/{uploaded_file.type.split('/')[-1]}")

    if st.button("開始轉換", type="primary"):
        with st.spinner("正在轉換中，請稍候（第一次載入模型較慢）..."):
            try:
                # 寫入暫存檔
                with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(uploaded_file.name)[1]) as tmp:
                    tmp.write(uploaded_file.getvalue())
                    tmp_path = tmp.name

                # 載入模型
                model = load_whisper_model(model_size)

                # 語言參數
                lang = None
                if language == "中文 (zh)":
                    lang = "zh"
                elif language == "英文 (en)":
                    lang = "en"

                # 轉寫
                result = model.transcribe(tmp_path, language=lang)
                text = result["text"].strip()

                # 清理暫存
                os.unlink(tmp_path)

                if not text:
                    st.warning("沒有辨識到任何文字。")
                else:
                    st.success("轉換完成！")
                    st.subheader("轉換結果")
                    st.text_area("文字內容", text, height=300)

                    # 產生 Word
                    doc = Document()
                    doc.add_heading("語音轉文字結果", level=1)
                    doc.add_paragraph(text)

                    buffer = BytesIO()
                    doc.save(buffer)
                    buffer.seek(0)

                    st.download_button(
                        label="📥 下載 Word 檔",
                        data=buffer,
                        file_name="語音轉文字結果.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )

            except Exception as e:
                st.error(f"轉換失敗：{str(e)}")
                st.exception(e)
else:
    st.info("請先上傳音訊檔。")
