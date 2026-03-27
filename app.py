from __future__ import annotations

from pathlib import Path
from typing import List

import streamlit as st

from ppt_translator import (
    PPTTranslationError,
    get_runtime_environment,
    translate_presentation,
)


st.set_page_config(
    page_title="PPT 中英翻译与自动排版",
    page_icon="📄",
    layout="centered",
)

RUNTIME_ENV = get_runtime_environment()
GEMINI_MODELS = [
    "gemini-2.5-flash",
    "gemini-2.5-pro",
    "gemini-2.0-flash",
]


PAGE_CSS = """
<style>
    .stApp {
        background:
            radial-gradient(circle at top left, rgba(0, 128, 96, 0.10), transparent 30%),
            linear-gradient(180deg, #f8fbf8 0%, #eef5f2 100%);
    }
    .block-container {
        max-width: 900px;
        padding-top: 2rem;
        padding-bottom: 3rem;
    }
    .hero-card {
        background: rgba(255, 255, 255, 0.92);
        border: 1px solid rgba(20, 82, 56, 0.08);
        border-radius: 24px;
        padding: 1.6rem 1.8rem;
        box-shadow: 0 18px 50px rgba(22, 50, 38, 0.10);
        margin-bottom: 1rem;
    }
    .hero-title {
        font-size: 2rem;
        font-weight: 700;
        color: #163226;
        margin-bottom: 0.35rem;
    }
    .hero-subtitle {
        color: #466356;
        line-height: 1.65;
        font-size: 1rem;
    }
    .tip-row {
        display: grid;
        grid-template-columns: repeat(3, 1fr);
        gap: 0.8rem;
        margin-top: 1.2rem;
    }
    .tip-card {
        background: #f4faf6;
        border: 1px solid rgba(34, 100, 70, 0.10);
        border-radius: 18px;
        padding: 0.9rem 1rem;
        color: #224a39;
        min-height: 84px;
    }
    .tip-card strong {
        display: block;
        color: #163226;
        margin-bottom: 0.2rem;
    }
    @media (max-width: 768px) {
        .tip-row {
            grid-template-columns: 1fr;
        }
    }
</style>
"""


def reset_output_if_file_changed(uploaded_file: st.runtime.uploaded_file_manager.UploadedFile) -> None:
    token = None
    if uploaded_file is not None:
        token = "{0}:{1}".format(uploaded_file.name, uploaded_file.size)

    if st.session_state.get("upload_token") != token:
        st.session_state["upload_token"] = token
        st.session_state.pop("translated_bytes", None)
        st.session_state.pop("download_name", None)
        st.session_state.pop("process_log", None)
        st.session_state.pop("last_error", None)


def build_download_name(source_name: str) -> str:
    path = Path(source_name)
    stem = path.stem or "translated"
    return "{0}_EN.pptx".format(stem)


def render_page() -> None:
    st.markdown(PAGE_CSS, unsafe_allow_html=True)
    st.markdown(
        """
        <div class="hero-card">
            <div class="hero-title">PPT 中英翻译与自动排版</div>
            <div class="hero-subtitle">
                上传中文 PPT，调用你自己的 Gemini API Key 完成逐页翻译，并尽量保留原始样式、图片、
                图表、背景与页面结构。处理完成后，网页会直接返回新的 .pptx 文件。
            </div>
            <div class="tip-row">
                <div class="tip-card">
                    <strong>1. 填入 API Key</strong>
                    Key 只在当前会话中使用，不写入文件。
                </div>
                <div class="tip-card">
                    <strong>2. 上传 .pptx</strong>
                    支持点击选择或拖拽到上传区。
                </div>
                <div class="tip-card">
                    <strong>3. 开始翻译</strong>
                    页面会显示逐页进度，完成后可直接下载。
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    api_key = st.text_input(
        "Gemini API Key",
        type="password",
        placeholder="AIza...",
        help="网页不会把 API Key 写入磁盘；刷新页面后需要重新输入。",
    )

    model = st.selectbox(
        "Gemini 模型",
        options=GEMINI_MODELS,
        index=0,
        help="默认使用 gemini-2.5-flash，兼顾速度与成本；如果 `pro` 遇到 429，优先切回 `flash`。",
    )

    uploaded_file = st.file_uploader(
        "上传 PPTX 文件",
        type=["pptx"],
        help="支持拖拽上传，仅处理 .pptx 文件。",
    )
    reset_output_if_file_changed(uploaded_file)

    if RUNTIME_ENV["ready"]:
        st.info("当前运行引擎: {0}".format(RUNTIME_ENV["message"]))
    else:
        st.error(RUNTIME_ENV["message"])

    if uploaded_file is not None:
        st.caption(
            "已选择文件: {0} ({1:.2f} MB)".format(
                uploaded_file.name,
                uploaded_file.size / 1024.0 / 1024.0,
            )
        )

    start_clicked = st.button(
        "开始翻译",
        type="primary",
        use_container_width=True,
        disabled=not api_key or uploaded_file is None or not RUNTIME_ENV["ready"],
    )

    status_placeholder = st.empty()
    progress_placeholder = st.empty()
    log_placeholder = st.empty()

    if start_clicked and uploaded_file is not None:
        logs: List[str] = []
        progress_bar = progress_placeholder.progress(0)

        def on_progress(current: int, total: int, detail: str) -> None:
            percent = 0 if total == 0 else min(int(current / float(total) * 100), 100)
            status_placeholder.info(
                "正在处理第 {0} 页，共 {1} 页\n\n{2}".format(current, total, detail)
            )
            progress_bar.progress(percent)
            logs.append("第 {0}/{1} 页: {2}".format(current, total, detail))
            log_placeholder.code("\n".join(logs[-12:]), language="text")

        try:
            translated_bytes = translate_presentation(
                uploaded_file.getvalue(),
                api_key=api_key,
                progress_callback=on_progress,
                model=model,
            )
            download_name = build_download_name(uploaded_file.name)
            st.session_state["translated_bytes"] = translated_bytes
            st.session_state["download_name"] = download_name
            st.session_state["process_log"] = logs
            st.session_state.pop("last_error", None)
            status_placeholder.success("翻译完成，新的 PPT 文件已经生成。")
            progress_bar.progress(100)
        except PPTTranslationError as exc:
            st.session_state["last_error"] = str(exc)
            progress_placeholder.empty()
            status_placeholder.error(str(exc))
        except Exception as exc:  # pragma: no cover
            st.session_state["last_error"] = str(exc)
            progress_placeholder.empty()
            status_placeholder.error(
                "处理过程中出现未预期错误: {0}".format(exc)
            )

    if st.session_state.get("translated_bytes"):
        st.download_button(
            "下载翻译后的 PPT",
            data=st.session_state["translated_bytes"],
            file_name=st.session_state["download_name"],
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True,
        )

    with st.expander("处理说明", expanded=False):
        st.markdown(
            """
            - 只会读取和替换形状与表格中的文本，不会主动压缩、重绘或替换图片、图表、背景。
            - 组合图形会递归遍历内部文本。
            - 纯数字或不含中文的内容会自动跳过，减少不必要的 Token 消耗。
            - 译文较长时会启用自动换行，并对文本框高度、表格行高和下方重叠元素做保守式微调。
            - 当前机器如果没有 `python-pptx`，会自动改走本机 PowerPoint 引擎。
            - 如果 Gemini 返回 429，通常是当前模型限流或额度不足，优先换成 `gemini-2.5-flash` 再试。
            """
        )

    if st.session_state.get("process_log"):
        with st.expander("最近一次处理日志", expanded=False):
            st.code("\n".join(st.session_state["process_log"]), language="text")


if __name__ == "__main__":
    render_page()
