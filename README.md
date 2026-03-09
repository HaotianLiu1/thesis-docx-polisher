# Thesis DOCX Polisher

中文 / English 双语说明。

## 中文说明

### 项目简介
`thesis-docx-polisher` 是一个用于论文初稿润色与“降AI表达”的可复用技能与脚本方案。它会按段调用 OpenAI 兼容接口，对段落做**小幅微调**，并在输出 Word 中进行差异标注：

- 新增内容：红色字体
- 删除内容：删除线
- 未修改内容：保持默认样式

### 主要特性
- 支持任意 `.docx` 学术文稿（论文/报告）
- 按一级标题自动定位章节范围
- 可配置开始/结束章节
- 可配置标题样式关键字与参考文献标题
- 内置失败重试、超时控制、失败日志输出
- 不内置任何明文 API Key

### 目录结构
```text
skills/
  thesis-docx-polisher/
    SKILL.md
    scripts/
      polish_docx_via_api.py
    references/
      api_and_prompt.md
      docx_workflow_notes.md
```

### 依赖
```bash
pip install python-docx requests
```

### 快速使用
```bash
python "skills/thesis-docx-polisher/scripts/polish_docx_via_api.py" \
  --input "你的论文.docx" \
  --output "你的论文_润色版.docx" \
  --base-url "<openai-compatible-base-url>" \
  --api-key "<your-api-key>" \
  --model "<model-name>" \
  --start-chapter 1 \
  --end-chapter 5 \
  --chapter-style-key "Heading 1" \
  --reference-title "参考文献" \
  --retries 2 \
  --retry-backoff 1.5 \
  --timeout 180 \
  --fail-log "failures.jsonl"
```

### 常用参数
- `--start-chapter` / `--end-chapter`：章节范围（按一级标题顺序，`--end-chapter 0` 表示直到参考文献）
- `--chapter-style-key`：一级标题样式关键字（默认 `Heading 1`）
- `--reference-title`：参考文献标题文本（默认 `参考文献`）
- `--skip-style-keys`：跳过的样式关键字，逗号分隔
- `--limit`：调试时仅处理前 N 个候选段落

### 注意事项
- 本仓库不包含受限制分发条款的第三方技能内容。
- 请自行确保你使用的 API 服务与论文数据处理符合所在机构规范。

---

## English

### Overview
`thesis-docx-polisher` is a reusable skill + script workflow for polishing thesis drafts and reducing over-AI writing style. It sends paragraphs to an OpenAI-compatible API with **minimal edits**, then writes a DOCX with visual diff markers:

- Inserted text: red font
- Deleted text: strikethrough
- Unchanged text: default style

### Features
- Works on general academic `.docx` documents
- Chapter-range processing based on first-level headings
- Configurable chapter/style/reference detection
- Retry/timeout/failure-log support
- No hardcoded API keys

### Dependencies
```bash
pip install python-docx requests
```

### Quick Start
```bash
python "skills/thesis-docx-polisher/scripts/polish_docx_via_api.py" \
  --input "thesis.docx" \
  --output "thesis_polished.docx" \
  --base-url "<openai-compatible-base-url>" \
  --api-key "<your-api-key>" \
  --model "<model-name>" \
  --start-chapter 1 \
  --end-chapter 5 \
  --chapter-style-key "Heading 1" \
  --reference-title "参考文献"
```

### Key Arguments
- `--start-chapter` / `--end-chapter`: chapter window by first-level heading order (`--end-chapter 0` = until reference section)
- `--chapter-style-key`: heading style key (default: `Heading 1`)
- `--reference-title`: reference section title text
- `--skip-style-keys`: comma-separated style keywords to skip
- `--fail-log`: write API-failed paragraph indexes to JSONL

### Notes
- This repo intentionally excludes third-party skill materials with restricted redistribution terms.
- Ensure your API usage and document handling comply with your institutional policies.
