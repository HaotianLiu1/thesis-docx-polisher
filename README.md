# Thesis DOCX Polisher Skill (Claude Code)

中文 / English

---

## 中文（如何作为 Claude Code 技能使用）

### 这是什么
这是一个可直接放入 **Claude Code skills 目录** 的技能（不是单纯脚本集合）。

- 技能名：`thesis-docx-polisher`
- 核心能力：对论文 `.docx` 逐段调用 OpenAI 兼容 API 进行轻量润色，并在输出文档中做差异标注：
  - 新增：红色
  - 删除：删除线
  - 未变：默认样式

> 适用于“论文初稿润色、降AI语气、保持原意微调”的场景。

---

### 目录结构（必须保持）

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

---

### 1）如何导入到别人的 Claude Code

将本仓库里的 `skills/thesis-docx-polisher` 目录复制到对方项目中：

```text
<对方项目根目录>/.claude/skills/thesis-docx-polisher
```

也就是最终路径应为：

```text
<project>/.claude/skills/thesis-docx-polisher/SKILL.md
```

放好后，Claude Code 会自动发现该技能。

---

### 2）依赖安装

```bash
pip install python-docx requests
```

---

### 3）API 配置方式（推荐环境变量）

#### Linux/macOS (bash/zsh)
```bash
export OPENAI_BASE_URL="<openai-compatible-base-url>"
export OPENAI_API_KEY="<your-api-key>"
export OPENAI_MODEL="<model-name>"
```

#### Windows PowerShell
```powershell
$env:OPENAI_BASE_URL="<openai-compatible-base-url>"
$env:OPENAI_API_KEY="<your-api-key>"
$env:OPENAI_MODEL="<model-name>"
```

> 也可在命令行参数中显式传 `--base-url --api-key --model`。

---

### 4）如何在 Claude Code 中调用

有两种常见方式：

#### 方式 A：自然语言触发技能（推荐）
在 Claude Code 里直接说：

- “请用 thesis-docx-polisher 对 `xxx.docx` 做第一到第五章润色，输出 `xxx_润色.docx`，用我当前配置的 API。”

Claude 会根据 `SKILL.md` 自动使用该技能流程。

#### 方式 B：直接执行脚本（显式）
```bash
python ".claude/skills/thesis-docx-polisher/scripts/polish_docx_via_api.py" \
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

---

### 5）关键参数

- `--start-chapter` / `--end-chapter`：按一级标题顺序处理章节
  - `--end-chapter 0` = 一直到参考文献标题
- `--chapter-style-key`：一级标题样式关键字（默认 `Heading 1`）
- `--reference-title`：参考文献标题文本（默认 `参考文献`）
- `--skip-style-keys`：跳过样式关键字（逗号分隔）
- `--fail-log`：记录 API 失败段落（JSONL）
- `--limit`：调试时仅处理前 N 个候选段落

---

### 6）给使用者的建议（降AI但不过度改写）

- 模型提示已约束为“微调优先，不改变原意”。
- 建议先小范围试跑（`--limit`）确认风格后再全量。
- 若接口不稳定，开启重试与失败日志后补跑失败段落。

---

## English (Using this as a Claude Code Skill)

### What this is
This is a **drop-in Claude Code skill** (not just a standalone script).

- Skill name: `thesis-docx-polisher`
- Purpose: paragraph-level thesis polishing via an OpenAI-compatible API with DOCX diff styling:
  - insertions in red
  - deletions as strikethrough
  - unchanged text left as-is

---

### Required layout

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

---

### 1) Import into another Claude Code project

Copy `skills/thesis-docx-polisher` into:

```text
<target-project>/.claude/skills/thesis-docx-polisher
```

The key file must exist at:

```text
<target-project>/.claude/skills/thesis-docx-polisher/SKILL.md
```

Claude Code will discover the skill automatically.

---

### 2) Install dependencies

```bash
pip install python-docx requests
```

---

### 3) Configure API

#### Linux/macOS
```bash
export OPENAI_BASE_URL="<openai-compatible-base-url>"
export OPENAI_API_KEY="<your-api-key>"
export OPENAI_MODEL="<model-name>"
```

#### Windows PowerShell
```powershell
$env:OPENAI_BASE_URL="<openai-compatible-base-url>"
$env:OPENAI_API_KEY="<your-api-key>"
$env:OPENAI_MODEL="<model-name>"
```

You can also pass `--base-url --api-key --model` directly in CLI.

---

### 4) Invoke from Claude Code

#### A) Natural language invocation (recommended)
Ask Claude Code something like:

- “Use thesis-docx-polisher to polish chapters 1–5 in `thesis.docx` and output `thesis_polished.docx` with my current API config.”

#### B) Direct script execution
```bash
python ".claude/skills/thesis-docx-polisher/scripts/polish_docx_via_api.py" \
  --input "thesis.docx" \
  --output "thesis_polished.docx" \
  --base-url "<openai-compatible-base-url>" \
  --api-key "<your-api-key>" \
  --model "<model-name>" \
  --start-chapter 1 \
  --end-chapter 5
```

---

### 5) Important args

- `--start-chapter` / `--end-chapter`: chapter range by first-level heading order
- `--chapter-style-key`: heading style key (default `Heading 1`)
- `--reference-title`: reference section title text
- `--skip-style-keys`: comma-separated style keywords to skip
- `--fail-log`: JSONL log for failed paragraph calls
- `--limit`: process only first N candidate paragraphs (debug)

---

## License / Compliance

This repository intentionally excludes third-party skill materials with restricted redistribution terms.
Please ensure API usage and document processing comply with your local policy and institution rules.
