from __future__ import annotations

import tempfile
from collections import Counter
from pathlib import Path

import streamlit as st

from checker import SUPPORTED, find_tesseract, run_check


st.set_page_config(page_title="文档内容批量核对", page_icon="🔎", layout="wide")


def save_upload(upload, folder: Path, prefix: str) -> Path:
    safe_name = Path(upload.name).name
    path = folder / f"{prefix}_{safe_name}"
    path.write_bytes(upload.getbuffer())
    return path


def result_table(results):
    return [
        {
            "待核对文件": item.source_file.split("_", 1)[-1],
            "位置": item.source_location,
            "待核对内容": item.source_text,
            "Word位置": item.reference_location,
            "Word参考内容": item.reference_text,
            "相似度(%)": item.similarity,
            "结论": item.result,
            "关键数字/日期差异": item.important_change,
        }
        for item in results
    ]


st.title("🔎 文档内容批量核对")
st.caption("以 Word 为基准，批量核对 PPT、图片和其他文档，并下载差异报告。")

with st.sidebar:
    st.header("核对设置")
    threshold = st.slider("疑似差异阈值", 60, 100, 88, 1, help="相似度达到此值时标为疑似差异。")
    enable_ocr = st.toggle("启用图片 OCR", value=True)
    if find_tesseract():
        st.success("OCR 已就绪")
    else:
        st.warning("OCR 引擎未安装，图片将被跳过。Docker 部署可自动安装。")
    st.divider()
    st.caption("建议将非绿色结果进行人工复核。上传文件仅用于当次处理。")

left, right = st.columns(2, gap="large")
with left:
    st.subheader("1. 基准文档")
    reference_upload = st.file_uploader(
        "上传基准 Word",
        type=["docx", "txt", "md"],
        accept_multiple_files=False,
        help="推荐使用 .docx。",
    )
with right:
    st.subheader("2. 待核对文件")
    source_uploads = st.file_uploader(
        "批量上传 PPT、Word 或图片",
        type=sorted(suffix.lstrip(".") for suffix in SUPPORTED),
        accept_multiple_files=True,
    )

ready = reference_upload is not None and bool(source_uploads)
run = st.button("开始批量核对", type="primary", use_container_width=True, disabled=not ready)

if run:
    try:
        with st.status("正在提取文字并核对…", expanded=True) as status:
            with tempfile.TemporaryDirectory(prefix="online_doc_check_") as temp_name:
                temp = Path(temp_name)
                reference = save_upload(reference_upload, temp, "reference")
                sources = [save_upload(upload, temp, f"source_{idx:03d}") for idx, upload in enumerate(source_uploads, 1)]
                output = temp / "reports"
                st.write(f"正在处理 {len(sources)} 个待核对文件…")
                xlsx, html_report, results, warnings = run_check(
                    reference, sources, output, threshold / 100, enable_ocr
                )
                st.session_state["last_result"] = {
                    "xlsx": xlsx.read_bytes(),
                    "html": html_report.read_bytes(),
                    "rows": result_table(results),
                    "counts": dict(Counter(item.result for item in results)),
                    "warnings": list(dict.fromkeys(warnings)),
                }
            status.update(label="核对完成", state="complete", expanded=False)
    except Exception as exc:
        st.error(f"处理失败：{exc}")

data = st.session_state.get("last_result")
if data:
    st.divider()
    st.subheader("核对结果")
    labels = ["一致", "疑似差异", "不一致", "未在基准文档中找到"]
    columns = st.columns(4)
    for column, label in zip(columns, labels):
        column.metric(label, data["counts"].get(label, 0))

    if data["warnings"]:
        with st.expander("处理提示", expanded=True):
            for warning in data["warnings"]:
                st.warning(warning)

    download_xlsx, download_html = st.columns(2)
    download_xlsx.download_button(
        "下载 Excel 差异报告",
        data["xlsx"],
        file_name="文档内容核对报告.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    download_html.download_button(
        "下载 HTML 浏览报告",
        data["html"],
        file_name="文档内容核对报告.html",
        mime="text/html",
        use_container_width=True,
    )

    choices = ["全部", "只看需复核项"] + labels
    view = st.segmented_control("筛选结果", choices, default="只看需复核项")
    rows = data["rows"]
    if view == "只看需复核项":
        rows = [row for row in rows if row["结论"] != "一致"]
    elif view and view != "全部":
        rows = [row for row in rows if row["结论"] == view]
    st.dataframe(rows, use_container_width=True, hide_index=True, height=520)

