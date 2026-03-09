import argparse
import json
import os
import re
import sys
import time
from difflib import SequenceMatcher
from pathlib import Path

import requests
from docx import Document
from docx.oxml.ns import qn
from docx.shared import RGBColor

DEFAULT_BASE_URL = os.getenv("OPENAI_BASE_URL", "")
DEFAULT_API_KEY = os.getenv("OPENAI_API_KEY", "")
DEFAULT_MODEL = os.getenv("OPENAI_MODEL", "gpt-5.4")
DEFAULT_SKIP_STYLE_KEYS = ["Heading", "标题-图", "标题-表格", "标题-无编号", "参考文献", "TOC"]

SYSTEM_PROMPT = """
你是中文学术论文润色助手。请先判断段落是否存在以下问题：
- AI语气过重、表达机械
- 语句不通顺或衔接生硬
- 连接词过于套路化

如果没有明显问题，请尽量保持原文不变。
如果有问题，只做微调，严格遵守：
1) 不大幅改写，不改变原意，不新增事实。
2) 优先将过渡词和连接词替换为基础、常用词，表达尽量简单直接。
3) 避免机械连接，使用自然衔接；句式可适度变化（简单句/复合句/插入语混合），避免连续整齐短句。
4) 保留术语、缩写、编号、公式、图表编号、引用编号。

只输出 JSON：
{"need_edit": true/false, "revised_text": "修改后段落"}
不要输出任何额外说明。
""".strip()

TOKEN_RE = re.compile(
    r"\s+|[A-Za-z0-9]+(?:[-._/][A-Za-z0-9]+)*|[\u4e00-\u9fff]|[^\w\s]",
    flags=re.UNICODE,
)


class ApiClient:
    def __init__(self, base_url: str, api_key: str, model: str, timeout: int, retries: int, retry_backoff: float):
        self.url = base_url.rstrip("/") + "/v1/chat/completions"
        self.model = model
        self.timeout = timeout
        self.retries = retries
        self.retry_backoff = retry_backoff
        self.headers = {
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
        }


def try_parse_json(text: str):
    s = text.strip()
    if s.startswith("```"):
        s = re.sub(r"^```(?:json)?\s*", "", s)
        s = re.sub(r"\s*```$", "", s)

    try:
        obj = json.loads(s)
        if isinstance(obj, dict):
            return obj
    except Exception:
        pass

    decoder = json.JSONDecoder()
    for i, ch in enumerate(s):
        if ch != "{":
            continue
        try:
            obj, _ = decoder.raw_decode(s[i:])
            if isinstance(obj, dict):
                return obj
        except Exception:
            continue
    return None


def tokenize(text: str):
    return TOKEN_RE.findall(text)


def diff_chunks(old: str, new: str):
    a, b = tokenize(old), tokenize(new)
    sm = SequenceMatcher(a=a, b=b)
    out = []
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        oa = "".join(a[i1:i2])
        nb = "".join(b[j1:j2])
        if tag == "equal" and nb:
            out.append(("equal", nb))
        elif tag == "delete" and oa:
            out.append(("delete", oa))
        elif tag == "insert" and nb:
            out.append(("insert", nb))
        elif tag == "replace":
            if oa:
                out.append(("delete", oa))
            if nb:
                out.append(("insert", nb))

    merged = []
    for kind, chunk in out:
        if not chunk:
            continue
        if merged and merged[-1][0] == kind:
            merged[-1] = (kind, merged[-1][1] + chunk)
        else:
            merged.append((kind, chunk))
    return merged


def clear_paragraph_content(paragraph):
    pxml = paragraph._p
    for child in list(pxml):
        if child.tag != qn("w:pPr"):
            pxml.remove(child)


def write_diff(paragraph, old: str, new: str):
    chunks = diff_chunks(old, new)
    if not chunks:
        chunks = [("equal", old)]

    clear_paragraph_content(paragraph)
    for kind, text in chunks:
        run = paragraph.add_run(text)
        if kind == "insert":
            run.font.color.rgb = RGBColor(0xFF, 0x00, 0x00)
        elif kind == "delete":
            run.font.strike = True


def is_target_paragraph(text: str, style: str, min_len: int, skip_style_keys):
    stripped = text.strip()
    if not stripped:
        return False
    if any(key in style for key in skip_style_keys):
        return False
    if len(stripped) < min_len:
        return False
    return True


def heading1_indices(doc: Document, chapter_style_key: str):
    out = []
    for i, p in enumerate(doc.paragraphs):
        style = p.style.name if p.style is not None else ""
        if chapter_style_key in style:
            out.append(i)
    return out


def reference_index(doc: Document, reference_title: str):
    for i, p in enumerate(doc.paragraphs):
        if p.text.strip() == reference_title:
            return i
    return len(doc.paragraphs)


def chapter_range(doc: Document, start_chapter: int, end_chapter: int, chapter_style_key: str, reference_title: str):
    h1 = heading1_indices(doc, chapter_style_key)
    if not h1:
        raise RuntimeError(f"未找到一级标题样式关键字: {chapter_style_key}")
    if start_chapter < 1:
        raise RuntimeError("--start-chapter 必须 >= 1")
    if start_chapter > len(h1):
        raise RuntimeError(
            f"--start-chapter={start_chapter} 超出范围，文档仅检测到 {len(h1)} 个一级标题"
        )

    start = h1[start_chapter - 1]
    ref = reference_index(doc, reference_title)

    if end_chapter <= 0:
        end = ref
    elif end_chapter < start_chapter:
        raise RuntimeError("--end-chapter 不能小于 --start-chapter")
    elif end_chapter < len(h1):
        end = h1[end_chapter]
    else:
        end = ref

    return start, end, h1, ref


def call_openai_compatible(session, client: ApiClient, paragraph_text: str):
    payload = {
        "model": client.model,
        "temperature": 0.2,
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": paragraph_text},
        ],
    }

    last_error = None
    for attempt in range(client.retries + 1):
        try:
            resp = session.post(client.url, headers=client.headers, json=payload, timeout=client.timeout)
            if resp.status_code >= 500 and attempt < client.retries:
                time.sleep(client.retry_backoff * (2**attempt))
                continue
            resp.raise_for_status()

            data = resp.json()
            content = data.get("choices", [{}])[0].get("message", {}).get("content", "")
            parsed = try_parse_json(content)
            if not parsed:
                return False, paragraph_text, "json_parse_failed"

            revised = str(parsed.get("revised_text", "")).strip() or paragraph_text
            need_edit = bool(parsed.get("need_edit", False))

            if not need_edit or revised == paragraph_text:
                return False, paragraph_text, "ok"
            return True, revised, "ok"

        except Exception as exc:
            last_error = exc
            if attempt < client.retries:
                time.sleep(client.retry_backoff * (2**attempt))
                continue
            raise last_error


def print_analysis(doc: Document, chapter_indexes, ref, chapter_style_key):
    print(f"总段落数: {len(doc.paragraphs)}")
    print(f"一级标题数量: {len(chapter_indexes)}")
    print(f"一级标题样式关键字: {chapter_style_key}")
    for n, idx in enumerate(chapter_indexes, 1):
        title = doc.paragraphs[idx].text.strip()
        print(f"第{n}章标题段落[{idx}]: {title}")
    print(f"参考文献段落索引: {ref}")


def main():
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="ignore")

    parser = argparse.ArgumentParser(
        description="通过 OpenAI 兼容 API 对 DOCX 论文分段润色并标注差异（新增红色，删除删除线）"
    )
    parser.add_argument("--input", required=True, help="输入 .docx")
    parser.add_argument("--output", required=True, help="输出 .docx")
    parser.add_argument("--base-url", default=DEFAULT_BASE_URL, help="OpenAI 兼容接口基地址")
    parser.add_argument("--api-key", default=DEFAULT_API_KEY, help="OpenAI 兼容接口 key")
    parser.add_argument("--model", default=DEFAULT_MODEL, help="模型名")

    parser.add_argument("--start-chapter", type=int, default=1, help="开始章（从1开始）")
    parser.add_argument("--end-chapter", type=int, default=5, help="结束章（包含）。设为0表示直到参考文献")
    parser.add_argument("--chapter-style-key", default="Heading 1", help="一级标题样式关键字")
    parser.add_argument("--reference-title", default="参考文献", help="参考文献标题文本")
    parser.add_argument(
        "--skip-style-keys",
        default=",".join(DEFAULT_SKIP_STYLE_KEYS),
        help="需要跳过的样式关键字，逗号分隔",
    )

    parser.add_argument("--min-len", type=int, default=6, help="最小段落长度")
    parser.add_argument("--timeout", type=int, default=180, help="单次API超时秒")
    parser.add_argument("--retries", type=int, default=2, help="API失败重试次数")
    parser.add_argument("--retry-backoff", type=float, default=1.5, help="重试退避基数秒")
    parser.add_argument("--sleep", type=float, default=0.0, help="每次调用后休眠秒")
    parser.add_argument("--limit", type=int, default=0, help="调试：最多处理多少个候选段落")
    parser.add_argument("--fail-log", default="", help="保存失败段落日志(jsonl)")
    args = parser.parse_args()

    if not args.base_url:
        raise RuntimeError("缺少 --base-url（或环境变量 OPENAI_BASE_URL）")
    if not args.api_key:
        raise RuntimeError("缺少 --api-key（或环境变量 OPENAI_API_KEY）")

    skip_style_keys = [x.strip() for x in args.skip_style_keys.split(",") if x.strip()]

    in_path = Path(args.input)
    out_path = Path(args.output)

    try:
        doc = Document(str(in_path))
    except Exception as exc:
        raise RuntimeError(f"无法读取输入文件: {in_path} ({type(exc).__name__}: {exc})") from exc

    start, end, chapter_indexes, ref = chapter_range(
        doc,
        args.start_chapter,
        args.end_chapter,
        args.chapter_style_key,
        args.reference_title,
    )
    print_analysis(doc, chapter_indexes, ref, args.chapter_style_key)
    print(f"处理范围: [{start}, {end})")

    client = ApiClient(
        base_url=args.base_url,
        api_key=args.api_key,
        model=args.model,
        timeout=args.timeout,
        retries=args.retries,
        retry_backoff=args.retry_backoff,
    )

    session = requests.Session()
    total = edited = skipped = api_err = 0

    fail_fp = None
    if args.fail_log:
        fail_path = Path(args.fail_log)
        fail_path.parent.mkdir(parents=True, exist_ok=True)
        fail_fp = fail_path.open("w", encoding="utf-8")

    try:
        for i in range(start, end):
            p = doc.paragraphs[i]
            style = p.style.name if p.style is not None else ""
            old = p.text

            if not is_target_paragraph(old, style, args.min_len, skip_style_keys):
                skipped += 1
                continue

            total += 1

            try:
                changed, revised, _ = call_openai_compatible(session, client, old)
            except Exception as exc:
                api_err += 1
                item = {"paragraph_index": i, "error": f"{type(exc).__name__}: {exc}"}
                if fail_fp:
                    fail_fp.write(json.dumps(item, ensure_ascii=False) + "\n")
                print(f"[API失败][{i}] {type(exc).__name__}: {exc}")
                continue

            if changed:
                write_diff(p, old, revised)
                edited += 1

            if args.sleep > 0:
                time.sleep(args.sleep)

            if args.limit > 0 and total >= args.limit:
                print("达到 --limit，提前结束")
                break

            if total % 20 == 0:
                print(f"已处理: {total}，已修改: {edited}，API失败: {api_err}")
    finally:
        if fail_fp:
            fail_fp.close()

    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_path))

    print("\n处理完成")
    print(f"候选段落: {total}")
    print(f"实际修改: {edited}")
    print(f"跳过段落: {skipped}")
    print(f"API失败: {api_err}")
    print(f"输出文件: {out_path}")


if __name__ == "__main__":
    main()
