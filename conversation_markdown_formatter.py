# -*- coding: utf-8 -*-
"""对话记录 Markdown 美化模块（纯函数，无 IO、无第三方依赖）。

职责：
    把结构化对话消息列表转换为适合归档阅读的 Markdown 文档字符串。

    - 输入：标题 + 消息列表，消息结构沿用 cursor_cleaner 数据层：
        {type: 1|2, time, text, thinking, tools: [{name, status, detail}]}
      type: 1=用户, 2=助手
    - 输出：完整的 Markdown 字符串。

    正文结构按 TUI 右侧圆点索引分组：每条用户消息是一个二级标题
    （对应一个圆点），其后的助手消息以三级标题归入该分组。

本模块不负责读取数据库、决定导出路径或写入文件；
文件写入由调用方（export_conversation_md 等）在拿到返回值之后执行。

内容保真约定：
    - 不删除/合并/重排任何消息；
    - 不修改正文文字（仅统一代码块围栏、压缩多余空行、
      降低消息内部标题层级以防破坏文档结构）；
    - 同一输入恒产生同一输出。
"""

from datetime import datetime
from typing import Any, List, Optional

# 消息类型常量（与 cursor_cleaner.fetch_conversation 语义一致）
TYPE_USER = 1
TYPE_ASSISTANT = 2

_ROLE_NAMES = {TYPE_USER: "用户", TYPE_ASSISTANT: "助手"}
_EMPTY_MESSAGE_NOTE = "> 此消息为空。"

# 连续纯文本超过该行数时按段落插入空行（列表/引用/代码块不拆分）
_LONG_PARAGRAPH_LINES = 30


# =====================================================================
# 时间戳
# =====================================================================

def _as_local_datetime(value: Any) -> Optional[datetime]:
    """将毫秒/秒/微秒时间戳或 ISO-8601 时间转为本地时间，失败返回 None。"""
    if value in (None, ""):
        return None

    numeric: Optional[float] = None
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        numeric = float(value)
    elif isinstance(value, str):
        text = value.strip()
        try:
            numeric = float(text)
        except ValueError:
            iso_text = text[:-1] + "+00:00" if text.endswith(("Z", "z")) else text
            try:
                parsed = datetime.fromisoformat(iso_text)
            except ValueError:
                return None
            return parsed.astimezone() if parsed.tzinfo is not None else parsed.astimezone()

    if numeric is None:
        return None
    magnitude = abs(numeric)
    if magnitude >= 1e14:
        seconds = numeric / 1_000_000
    elif magnitude >= 1e11:
        seconds = numeric / 1_000
    else:
        seconds = numeric
    try:
        return datetime.fromtimestamp(seconds).astimezone()
    except (OverflowError, OSError, ValueError):
        return None


def format_timestamp(value: Any) -> str:
    """原始时间戳 -> 'YYYY-MM-DD HH:MM:SS'（本地时区）；无法解析返回 ''。"""
    dt = _as_local_datetime(value)
    return dt.strftime("%Y-%m-%d %H:%M:%S") if dt else ""


def format_date(value: Any) -> str:
    """原始时间戳 -> 'YYYY-MM-DD'；无法解析返回 ''。"""
    dt = _as_local_datetime(value)
    return dt.strftime("%Y-%m-%d") if dt else ""


# =====================================================================
# 消息规范化
# =====================================================================

def normalize_message(message: Any) -> dict:
    """把单条消息整理为固定字段结构，缺失字段安全兜底。"""
    if not isinstance(message, dict):
        message = {}
    mtype = message.get("type")
    if mtype not in (TYPE_USER, TYPE_ASSISTANT):
        mtype = TYPE_ASSISTANT
    tools = message.get("tools")
    if not isinstance(tools, list):
        tools = []
    return {
        "type": mtype,
        "time": message.get("time"),
        "text": message.get("text") or "",
        "thinking": message.get("thinking") or "",
        "tools": tools,
    }


# =====================================================================
# 代码块与正文排版
# =====================================================================

def _longest_backtick_run(text: str) -> int:
    """正文中最长连续反引号的长度，用于决定围栏长度。"""
    longest = 0
    current = 0
    for ch in text:
        if ch == "`":
            current += 1
            if current > longest:
                longest = current
        else:
            current = 0
    return longest


def _format_code_block(code: str, lang: str = "") -> str:
    """把代码片段包装为完整闭合的围栏代码块。

    围栏长度取内容中最长连续反引号数 + 1（至少 3），保证内容里
    的反引号不会提前闭合围栏。代码内容与语言标识原样保留。
    """
    longest = _longest_backtick_run(code)
    fence = "`" * max(3, longest + 1)
    opener = fence + lang
    return f"{opener}\n{code.rstrip()}\n{fence}"


def _lower_headings(line: str) -> str:
    """消息内部标题整体降低 2 级（# -> ###），不修改标题文字。

    ### 及更深的标题保持原样，避免超过 Markdown 支持的 6 级上限。
    """
    stripped = line.lstrip()
    hashes = len(stripped) - len(stripped.lstrip("#"))
    if 0 < hashes < 3 and stripped[hashes:hashes + 1] in (" ", ""):
        return "#" * (hashes + 2) + stripped[hashes:]
    return line


def _split_long_paragraphs(lines: List[str]) -> List[str]:
    """超过阈值行数且不含列表/引用标记的连续文本，按行分组插入空行。"""
    if len(lines) <= _LONG_PARAGRAPH_LINES:
        return lines
    in_block = False
    block: List[str] = []
    result: List[str] = []

    def flush_block() -> None:
        if not block:
            return
        if len(block) > _LONG_PARAGRAPH_LINES:
            for start in range(0, len(block), _LONG_PARAGRAPH_LINES):
                result.extend(block[start:start + _LONG_PARAGRAPH_LINES])
                if start + _LONG_PARAGRAPH_LINES < len(block):
                    result.append("")
        else:
            result.extend(block)
        block.clear()

    for line in lines:
        stripped = line.lstrip()
        if (
            stripped.startswith(("- ", "* ", "+ ", "> ", "#", "`", "|", "---"))
            or stripped[:1].isdigit() and ". " in stripped[:4]
        ):
            flush_block()
            result.append(line)
            continue
        if stripped == "":
            flush_block()
            result.append(line)
            continue
        block.append(line)
    flush_block()
    return result


def _format_message_body(text: str) -> str:
    """正文整体排版：规范化换行、压缩空行、代码块重围栏、标题降级。"""
    raw_lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")

    # 代码块状态机：fence 长度由内容决定，先扫描后重写
    fence_len = 0          # 0=不在代码块内，否则为当前围栏长度
    fence_lang = ""
    code_lines: List[str] = []
    out: List[str] = []

    def flush_code() -> None:
        if code_lines:
            out.append(_format_code_block("\n".join(code_lines), fence_lang))
            code_lines.clear()

    i = 0
    while i < len(raw_lines):
        line = raw_lines[i]
        stripped = line.lstrip()
        if fence_len == 0:
            # 尝试开启代码块：行首 >=3 个反引号
            if stripped.startswith("```") or stripped.startswith("~~~"):
                marker = "`" if stripped.startswith("```") else "~"
                count = len(stripped) - len(stripped.lstrip(marker))
                if count >= 3:
                    fence_len = count
                    fence_lang = stripped[count:].strip()
                    i += 1
                    continue
            out.append(_lower_headings(line))
        else:
            if stripped.startswith("`" * fence_len) or stripped.startswith("~" * fence_len):
                flush_code()
                fence_len = 0
                fence_lang = ""
                i += 1
                continue
            code_lines.append(line)
        i += 1
    if fence_len > 0:
        flush_code()
        fence_len = 0

    # 压缩连续空行（正文内部；围栏已重写，无需再关心原 fence）
    collapsed: List[str] = []
    blank = 0
    for line in out:
        if line.strip() == "":
            blank += 1
            if blank > 1:
                continue
        else:
            blank = 0
        collapsed.append(line)

    # 长纯文本段落分组
    return "\n".join(_split_long_paragraphs(collapsed)).rstrip()


def _format_tools(tools: List[dict]) -> List[str]:
    """工具调用摘要 -> 引用块行（内容不截断，供归档）。"""
    lines: List[str] = []
    for tool in tools:
        if not isinstance(tool, dict):
            continue
        name = str(tool.get("name") or "工具")
        status = tool.get("status")
        head = f"> 🔧 {name}" + (f" [{status}]" if status else "")
        lines.append(head)
        detail = tool.get("detail") or ""
        if detail:
            lines.append(f"> `{detail}`")
    return lines


# =====================================================================
# 单条消息
# =====================================================================

def format_message(message: Any, index: int) -> str:
    """把一条消息格式化为 Markdown 段落（含角色标题），index 从 1 起。

    标题级别对应 TUI 右侧圆点索引：用户消息是分组起点（二级标题，
    每个圆点一条），其后的助手消息归入该分组（三级标题）。
    """
    m = normalize_message(message)
    role = _ROLE_NAMES[m["type"]]
    level = "##" if m["type"] == TYPE_USER else "###"
    ts = format_timestamp(m["time"])
    heading = f"{level} {role} · 第 {index} 条" + (f" · {ts}" if ts else "")
    parts: List[str] = [heading]

    body = _format_message_body(m["text"])
    if body:
        parts.append(body)
    if m["thinking"]:
        parts.append(f"> 💭 思考: {m['thinking']}")
    if m["tools"]:
        parts.extend(_format_tools(m["tools"]))
    if len(parts) == 1:
        parts.append(_EMPTY_MESSAGE_NOTE)
    parts.append("")
    parts.append("---")
    return "\n".join(parts)


# =====================================================================
# 元数据
# =====================================================================

def _format_metadata(title: str, msgs: List[dict]) -> List[str]:
    """会话信息块：只有原始数据中存在的信息才写入。"""
    lines: List[str] = ["## 会话信息"]
    dates: List[str] = []
    roles: List[str] = []
    for raw in msgs:
        m = normalize_message(raw)
        role = _ROLE_NAMES[m["type"]]
        if role not in roles:
            roles.append(role)
        d = format_date(m["time"])
        if d and d not in dates:
            dates.append(d)
    if dates:
        lines.append(f"- **日期：** {dates[0]}")
    lines.append(f"- **消息数量：** {len(msgs)}")
    if roles:
        lines.append(f"- **参与者：** {'、'.join(roles)}")
    lines.append("")
    return lines


# =====================================================================
# 文档级
# =====================================================================

def build_markdown_document(title: str, msgs: List[dict]) -> str:
    """组装完整文档：标题 + 主题 + 会话信息 + 按用户消息分组的对话正文。

    正文不再使用统一的 "## 对话正文" 标题：每条用户消息（对应 TUI
    右侧的一个圆点）直接作为二级标题分组，其后的助手消息以三级
    标题归入该分组，直到下一条用户消息。
    """
    title = (title or "").strip() or "对话记录"
    parts: List[str] = [f"# {title}", ""]
    parts.append(f"> 对话主题：{title}")
    parts.append("")
    parts.extend(_format_metadata(title, msgs))
    for i, msg in enumerate(msgs, start=1):
        parts.append(format_message(msg, i))
    return "\n".join(parts).rstrip() + "\n"


def format_conversation_markdown(title: str = "", msgs: Optional[List[dict]] = None) -> str:
    """将结构化对话数据转换为可读的 Markdown 文档字符串。

    参数与 cursor_cleaner 数据层对齐：
        title: 会话标题（可空）
        msgs:  消息列表 [{type, time, text, thinking, tools}]（可空）
    空消息列表时输出占位说明，不抛异常。
    """
    if not msgs:
        return f"# {(title or '').strip() or '对话记录'}\n\n_（无聊天记录或已无正文数据）_\n"
    return build_markdown_document(title, msgs)