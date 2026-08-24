# 文档内容批量核对工具

以 Word `.docx` 为基准，批量核对 PPT、Word、图片、TXT 和 Markdown 内容，生成可筛选的 Excel 报告和便于查看的 HTML 报告。

## 使用

1. 双击 `启动工具.bat`。
2. 选择一个基准 Word 文档。
3. 添加待核对的 PPT/图片/文档，也可一次添加整个文件夹。
4. 点击“开始批量核对”。
5. 在输出文件夹查看 `.xlsx` 和 `.html` 报告。

## 核对规则

- 自动忽略空格、换行、全半角和普通标点差异。
- 按文本块寻找 Word 中最相似的段落或表格行。
- 结果分为“一致”、“疑似差异”、“不一致”和“未在基准文档中找到”。
- 额外提取数字、日期、金额和百分比差异，便于优先人工复核。

## 图片 OCR

图片识别需要本机安装 Tesseract OCR，并建议加装简体中文语言包 `chi_sim`。未安装时，工具仍可正常核对 PPT 原生文字和 Office 文档，报告中会注明已跳过的图片。

## 命令行用法

```powershell
python checker.py --reference "基准.docx" --sources "待核对.pptx" "海报.png" --output "核对结果"
```

## 在线部署

双击 `启动在线版_本地预览.bat` 可先在本机浏览器预览网页版。

### Streamlit Community Cloud

1. 将本文件夹上传到一个 GitHub 仓库。
2. 在 Streamlit Community Cloud 新建应用，入口文件选择 `streamlit_app.py`。
3. `requirements.txt` 会安装 Python 依赖，`packages.txt` 会安装中英文 OCR。

### Docker（更适合内网或私有服务器）

```powershell
docker build -t document-checker .
docker run --rm -p 8501:8501 document-checker
```

然后访问 `http://localhost:8501`。可将同一个 Docker 镜像部署到云服务器、Render、Railway 或公司内网容器平台。

> 文档可能含有业务或个人信息。正式使用时建议优先部署在公司内网，或在网关增加登录认证，不要直接暴露为公开网址。

## 局限

- 目前支持 Office Open XML 格式（`.docx` / `.pptx`），不直接支持旧版 `.doc` / `.ppt`。
- OCR 质量受图片分辨率、字体和表格复杂度影响。
- “一致”表示标准化后文本一致；关键结论仍建议对报告中的非绿色项进行人工确认。

