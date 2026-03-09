# Thesis DOCX Polisher Skill (Claude Code)

中文 / English

---

## 中文（如何作为 Claude Code 技能使用）

### 这是什么
这是一个可直接放入 **Claude Code skills 目录** 的技能。

- 技能名：`thesis-docx-polisher`
- 核心能力：对论文 `.docx` 按批次调用 OpenAI 兼容 API 进行轻量润色，并在输出文档中做差异标注：
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

### 1）如何导入到 Claude Code

将本仓库里的 `skills/thesis-docx-polisher` 目录复制到项目中：

```text
<项目根目录>/.claude/skills/thesis-docx-polisher
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

- “请用 thesis-docx-polisher 对 `xxx.docx` 做第一到第五章润色，每次调用api的段落数为x，输出 `xxx_润色.docx`，用我当前配置的 API。”

并且：
- 若你没有明确给出 `start-chapter / end-chapter / paragraphs-per-call`，技能会先反问你这 3 个参数，再执行。

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
  --paragraphs-per-call 5 \
  --chapter-style-key "Heading 1" \
  --reference-title "参考文献" \
  --retries 2 \
  --retry-backoff 1.5 \
  --timeout 180 \
  --sleep 0.3 \
  --fail-log "failures.jsonl"
```

---

### 5）关键参数

- `--start-chapter` / `--end-chapter`：按一级标题顺序处理章节
  - `--end-chapter 0` = 一直到参考文献标题
- `--paragraphs-per-call`：每次 API 调用处理段落数（批大小）
  - `1` = 兼容旧行为（每段一次调用）
  - 推荐 `3~8`，可显著减少请求次数
  - 过大可能导致超时、限流放大、或整批 JSON 解析失败
- `--chapter-style-key`：一级标题样式关键字（默认 `Heading 1`）
- `--reference-title`：参考文献标题文本（默认 `参考文献`）
- `--skip-style-keys`：跳过样式关键字（逗号分隔）
- `--fail-log`：记录 API 失败段落与回退状态（JSONL）
- `--limit`：调试时仅处理前 N 个候选段落

---

### 6）批处理协议与容错（核心）

请求侧以 `items` 发送：

```json
{"items":[{"id":0,"text":"段落A"},{"id":1,"text":"段落B"}]}
```

模型返回：

```json
{"items":[{"id":0,"need_edit":true,"revised_text":"..."},{"id":1,"need_edit":false,"revised_text":"..."}]}
```

容错规则：
- 非 JSON：整批回退原文。
- `id` 缺失/非法：仅该段回退原文。
- `revised_text` 为空：仅该段回退原文。

---

### 7）重试与限流

- 自动重试：请求异常、HTTP `429`、HTTP `5xx`
- 退避策略：指数退避 `retry_backoff * 2^attempt`
- 若接口返回 `Retry-After`，优先按其等待
- `--sleep`：按批次轻量节流

---

### 8）给使用者的建议（降AI但不过度改写）

- 模型提示已约束为“微调优先，不改变原意”。
- 提示词针对不同模型效果差异大，建议针对论文类型和模型调整 skills 中的提示词。
- 建议先小范围试跑（`--limit`）确认风格后再全量。
- 若接口不稳定，开启重试与失败日志后补跑失败段落。

---

## English (Using this as a Claude Code Skill)

### What this is
This is a **drop-in Claude Code skill** (not just a standalone script).

- Skill name: `thesis-docx-polisher`
- Purpose: batch thesis polishing via an OpenAI-compatible API with DOCX diff styling:
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

If `start/end chapter` or `paragraphs-per-call` is missing, the skill should ask first.

#### B) Direct script execution
```bash
python ".claude/skills/thesis-docx-polisher/scripts/polish_docx_via_api.py" \
  --input "thesis.docx" \
  --output "thesis_polished.docx" \
  --base-url "<openai-compatible-base-url>" \
  --api-key "<your-api-key>" \
  --model "<model-name>" \
  --start-chapter 1 \
  --end-chapter 5 \
  --paragraphs-per-call 5
```

---

### 5) Important args

- `--start-chapter` / `--end-chapter`: chapter range by first-level heading order
- `--paragraphs-per-call`: paragraphs per request (`1` keeps legacy mode, `3~8` recommended)
- `--chapter-style-key`: heading style key (default `Heading 1`)
- `--reference-title`: reference section title text
- `--skip-style-keys`: comma-separated style keywords to skip
- `--fail-log`: JSONL log for failed/fallback paragraphs
- `--limit`: process only first N candidate paragraphs (debug)

---

## License / Compliance

This repository intentionally excludes third-party skill materials with restricted redistribution terms.
Please ensure API usage and document processing comply with your local policy and institution rules.
