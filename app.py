from io import BytesIO
from docx import Document
from google import genai
import streamlit as st

st.set_page_config(
    page_title="Gemini 逐字稿轉 Word 筆記",
    page_icon="📝",
    layout="centered",
)

st.title("📝 Gemini 逐字稿轉 Word 筆記")
st.caption("貼上逐字稿 ➡️ Gemini 自動整理 ➡️ 一鍵匯出 Word (.docx)")

# 取得 Gemini API Key（優先讀取 Streamlit Secrets，沒有則讓使用者在側邊欄輸入）
default_api_key = ""
try:
    default_api_key = st.secrets.get("GEMINI_API_KEY", "")
except Exception:
    default_api_key = ""

with st.sidebar:
    st.header("⚙️ Gemini 設定")
    api_key = (
        st.text_input(
            "Gemini API Key",
            value=default_api_key,
            type="password",
            help="可在 Google AI Studio 取得",
        )
        if not default_api_key
        else default_api_key
    )

    if default_api_key:
        st.success("已自動載入雲端 API Key")
    else:
        st.markdown("[前往 Google AI Studio 取得金鑰](https://aistudio.google.com/)")

    model_name = st.selectbox(
        "選擇 Gemini 模型",
        options=["gemini-3.8-flash", "gemini-4-hydrogen", "gemini-3.5-flash-lite"],
        index=0,
    )

    note_style = st.selectbox(
        "筆記整理模式",
        options=["通用重點筆記", "課堂講義筆記", "會議與待辦清單"],
        index=0,
    )

    doc_title = st.text_input("Word 文件標題", value="Gemini 錄音重點筆記")

transcript = st.text_area(
    "請貼上錄音逐字稿文字",
    height=260,
    placeholder="在這裡貼上逐字稿內容...",
)


def build_prompt(text: str, style: str) -> str:
    if style == "課堂講義筆記":
        structure = """
        # 課程核心摘要
        (用 2-3 句話概述本段課程重點)

        # 重要觀念與名詞解析
        - (列出核心概念、定義、專有名詞或原理說明)

        # 補充細節與實務舉例
        - (整理老師提到的案例、步驟、注意事項或延伸重點)

        # 複習重點檢核
        - (列出考前或課後最適合回顧的 3-5 個重點)
        """
    elif style == "會議與待辦清單":
        structure = """
        # 會議摘要
        (簡述討論背景與整體結論)

        # 討論重點與決議
        - (條列各議題的討論內容與已拍板決議)

        # 待辦事項 (Action Items)
        - (列出任務內容、負責對象與時程；若無明確資訊則註明「未提及」)
        """
    else:
        structure = """
        # 主題摘要
        (用一段話總結這段逐字稿的核心內容)

        # 重點整理
        - (去除口語贅字，條列出所有重要觀點、細節與知識點)

        # 待辦事項與結論
        - (列出後續行動、重要結論或提醒；若無則寫「無」)
        """

    return f"""
你是一位擅長整理資訊的筆記助手。請仔細閱讀以下錄音逐字稿，去除口語贅字與重複語句，保留關鍵事實、專有名詞與邏輯脈絡，並以繁體中文輸出。
請嚴格依照以下結構排版（使用 `#` 代表大段落標題，使用 `- ` 代表項目符號）：

{structure}

逐字稿內容：
{text}
""".strip()


def create_word_file(title: str, notes_text: str) -> BytesIO:
    doc = Document()
    doc.add_heading(title.strip() or "Gemini 錄音重點筆記", level=0)

    for raw_line in notes_text.split("\n"):
        line = raw_line.strip()
        if not line:
            continue
        if line.startswith("### "):
            doc.add_heading(line[4:].strip(), level=3)
        elif line.startswith("## "):
            doc.add_heading(line[3:].strip(), level=2)
        elif line.startswith("# "):
            doc.add_heading(line[2:].strip(), level=1)
        elif line.startswith("- ") or line.startswith("* "):
            doc.add_paragraph(line[2:].strip().replace("**", ""), style="List Bullet")
        else:
            doc.add_paragraph(line.replace("**", ""))

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


if st.button("✨ 用 Gemini 整理並產生 Word", type="primary", use_container_width=True):
    if not api_key:
        st.error("請先提供 Gemini API Key（可在左側輸入，或設定 st.secrets['GEMINI_API_KEY']）。")
    elif not transcript.strip():
        st.warning("請先貼上逐字稿文字再開始整理。")
    else:
        try:
            with st.spinner("Gemini 正在整理逐字稿並生成筆記..."):
                client = genai.Client(api_key=api_key)
                prompt = build_prompt(transcript, note_style)
                response = client.models.generate_content(
                    model=model_name,
                    contents=prompt,
                )
                notes_text = response.text or ""

            if not notes_text.strip():
                st.error("Gemini 沒有回傳文字內容，請稍後再試一次。")
            else:
                st.success("✅ 筆記整理完成！")

                with st.expander("預覽整理後的筆記", expanded=True):
                    st.markdown(notes_text)

                docx_buffer = create_word_file(doc_title, notes_text)
                safe_filename = (doc_title.strip() or "Gemini重點筆記") + ".docx"

                st.download_button(
                    label="📥 下載 Word 檔案 (.docx)",
                    data=docx_buffer,
                    file_name=safe_filename,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                )
        except Exception as exc:
            st.error(f"處理時發生錯誤：{exc}")
