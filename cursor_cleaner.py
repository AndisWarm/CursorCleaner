# -*- coding: utf-8 -*-
"""
Cursor 会话清理工具（TUI + CLI）

TUI 用法:
    python cursor_cleaner.py

CLI 用法（自动化/测试）:
    python cursor_cleaner.py --op preview
    python cursor_cleaner.py --op delete-archived --yes

操作列表:
    preview         扫描并分类会话（归档/残留/孤儿/未归档）
    delete-archived 删除已归档会话 + 镜像残留 + 正文孤儿（核心功能）
    backup          备份当前数据库
    restore         从备份恢复
    wipe-all        清空全部会话（危险）
    purge-index     清理会话搜索索引 conversation-search.db

数据模型（state.vscdb）:
    - composerHeaders: 每行一个会话，isArchived=1 表示已归档
    - ItemTable['composer.composerHeaders']: 侧边栏镜像列表（含 isArchived）
    - cursorDiskKV: 会话正文（composerData:<id> / bubbleId:<id>:* / ...）

会话状态分类:
    ARCHIVED      表或镜像标记为已归档
    ACTIVE        表或镜像存在且未归档
    MIRROR_ONLY   仅在镜像中（表行已删，残留）
    CONTENT_ONLY  仅在正文中（表/镜像都已删，孤儿数据）
"""

import argparse
import glob
import json
import os
import re
import shutil
import sqlite3
import subprocess
import sys
import time
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Set, Tuple

def _configure_utf8_console() -> None:
    """统一 Windows 控制台及标准输入/输出编码，避免 Codex/PowerShell 乱码。"""
    if sys.platform.startswith("win"):
        try:
            import ctypes

            kernel32 = ctypes.windll.kernel32
            # 对真实控制台生效；输出被 Codex 捕获为管道时调用失败也无妨。
            kernel32.SetConsoleCP(65001)
            kernel32.SetConsoleOutputCP(65001)
        except (AttributeError, OSError):
            pass

    for stream in (sys.stdin, sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, OSError):
            # 某些嵌入式运行环境提供的流没有 reconfigure()。
            pass


_configure_utf8_console()

GLOBAL_STORAGE = os.path.join(os.environ["APPDATA"], "Cursor", "User", "globalStorage")
DB = os.path.join(GLOBAL_STORAGE, "state.vscdb")
SEARCH_INDEX = os.path.join(GLOBAL_STORAGE, "conversation-search.db")

# 与 composerId 关联的正文键前缀（key 的 UUID 段即 composerId）
SESSION_KEY_PATTERNS = ("composerData:", "composerVirtualRowHeights:", "bubbleId:", "checkpointId:", "ofsContent:")
UUID_RE = re.compile(r"[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}")

S_ARCHIVED = "archived"
S_ACTIVE = "active"
S_MIRROR_ONLY = "mirror-only"
S_CONTENT_ONLY = "content-only"


# =====================================================================
# 数据层（纯函数，不依赖 UI）
# =====================================================================

@dataclass
class Session:
    """一个会话的跨表汇总视图。"""
    composer_id: str
    table_archived: bool = False      # composerHeaders 表中 isArchived
    mirror_archived: bool = False     # 镜像列表中 isArchived
    in_table: bool = False
    in_mirror: bool = False
    content_keys: int = 0             # cursorDiskKV 中关联键数量
    name: str = ""
    last_updated: Optional[int] = None
    workspace_id: str = ""

    @property
    def status(self) -> str:
        if self.in_table or self.in_mirror:
            if self.table_archived or self.mirror_archived:
                return S_ARCHIVED
            return S_ACTIVE
        if self.content_keys:
            return S_CONTENT_ONLY
        return S_ACTIVE

    @property
    def display_name(self) -> str:
        n = (self.name or "").strip()
        if not n:
            n = self.composer_id[:8]
        return n[:60]


def open_db_ro() -> sqlite3.Connection:
    """只读连接（可读到 WAL 内未合并数据）。"""
    return sqlite3.connect(f"file:{DB}?mode=ro", uri=True)


def open_db_rw() -> sqlite3.Connection:
    return sqlite3.connect(DB)


def cursor_running() -> bool:
    """Windows 下检测 Cursor.exe 是否在运行。"""
    if not sys.platform.startswith("win"):
        return False
    try:
        out = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq Cursor.exe"],
            capture_output=True, text=True, timeout=15,
        ).stdout
        return "Cursor.exe" in out
    except Exception:
        return False


def scan() -> List[Session]:
    """扫描三处数据源，合并为会话清单。只读。"""
    if not os.path.exists(DB):
        raise FileNotFoundError(f"找不到 {DB}")

    con = open_db_ro()
    try:
        cur = con.cursor()

        # 1) 表
        table_map: Dict[str, bool] = {}
        for cid, arch in cur.execute("SELECT composerId, isArchived FROM composerHeaders"):
            table_map[cid] = bool(arch)

        # 2) 镜像
        mirror_map: Dict[str, dict] = {}
        row = cur.execute("SELECT value FROM ItemTable WHERE key='composer.composerHeaders'").fetchone()
        if row:
            raw = row[0].encode("utf-8", "replace") if isinstance(row[0], str) else row[0]
            try:
                for h in json.loads(raw).get("allComposers", []):
                    cid = h.get("composerId")
                    if cid:
                        mirror_map[cid] = h
            except json.JSONDecodeError:
                pass

        # 3) 正文键中出现的 composerId -> 键数
        content_map: Dict[str, int] = {}
        for (key,) in cur.execute(
            "SELECT key FROM cursorDiskKV WHERE key LIKE 'composerData:%' OR key LIKE "
            "'composerVirtualRowHeights:%' OR key LIKE 'bubbleId:%' OR key LIKE 'checkpointId:%' OR key LIKE 'ofsContent:%'"
        ):
            m = UUID_RE.search(key)
            if m:
                cid = m.group(0)
                content_map[cid] = content_map.get(cid, 0) + 1

        all_ids = set(table_map) | set(mirror_map) | set(content_map)
        sessions = []
        for cid in all_ids:
            h = mirror_map.get(cid, {})
            sessions.append(Session(
                composer_id=cid,
                table_archived=table_map.get(cid, False),
                mirror_archived=bool(h.get("isArchived", False)),
                in_table=cid in table_map,
                in_mirror=cid in mirror_map,
                content_keys=content_map.get(cid, 0),
                name=h.get("name", ""),
                last_updated=h.get("lastUpdatedAt"),
                workspace_id=h.get("workspaceIdentifier", ""),
            ))
        sessions.sort(key=lambda s: (s.last_updated or 0), reverse=True)
        return sessions
    finally:
        con.close()


def classify(sessions: List[Session]) -> Dict[str, List[Session]]:
    out = {S_ARCHIVED: [], S_ACTIVE: [], S_MIRROR_ONLY: [], S_CONTENT_ONLY: []}
    for s in sessions:
        out[s.status].append(s)
    return out


def count_keys_for(con: sqlite3.Connection, cid: str) -> int:
    cur = con.execute(
        "SELECT COUNT(*) FROM cursorDiskKV WHERE key IN (?, ?) OR key LIKE ? OR key LIKE ? OR key LIKE ?",
        (f"composerData:{cid}", f"composerVirtualRowHeights:{cid}",
         f"bubbleId:{cid}:%", f"checkpointId:{cid}:%", f"ofsContent:{cid}:%"),
    )
    return cur.fetchone()[0]


def fetch_conversation(cid: str) -> List[dict]:
    """
    还原一个会话的聊天记录（按时间顺序）。
    数据源: composerData 的 fullConversationHeadersOnly（消息索引）+ 各 bubbleId 键（消息正文）。
    返回每条消息的 dict: {type, time, text, thinking, tools}
      type: 1=用户, 2=AI
      tools: [{name, status, detail}]（工具调用摘要）
    任何一步解析失败都跳过该条，不抛异常。
    """
    con = open_db_ro()
    try:
        cur = con.cursor()
        row = cur.execute("SELECT value FROM cursorDiskKV WHERE key=?", (f"composerData:{cid}",)).fetchone()
        if not row:
            return []
        raw = row[0]
        raw = raw.encode("utf-8", "replace") if isinstance(raw, str) else raw
        try:
            cdata = json.loads(raw)
        except json.JSONDecodeError:
            return []
        headers = cdata.get("fullConversationHeadersOnly") or []
        if not isinstance(headers, list):
            return []

        messages = []
        for h in headers:
            bid = h.get("bubbleId")
            if not bid:
                continue
            mtype = h.get("type")
            created = h.get("createdAt", "")
            brow = cur.execute("SELECT value FROM cursorDiskKV WHERE key=?", (f"bubbleId:{cid}:{bid}",)).fetchone()
            if not brow:
                continue
            bv = brow[0]
            bv = bv.encode("utf-8", "replace") if isinstance(bv, str) else bv
            try:
                bobj = json.loads(bv)
            except json.JSONDecodeError:
                continue

            tools = []
            td = bobj.get("toolFormerData")
            if isinstance(td, dict):
                tools.append(_tool_summary(td))
            elif isinstance(td, list):
                for item in td:
                    if isinstance(item, dict):
                        tools.append(_tool_summary(item))

            messages.append({
                "type": mtype,
                "time": created,
                "text": bobj.get("text") or "",
                "thinking": _extract_text(bobj.get("thinking")),
                "tools": tools,
            })
        return messages
    finally:
        con.close()


def _extract_text(v) -> str:
    """thinking 字段可能是字符串或 {text: ...} 结构，统一提取文本。"""
    if isinstance(v, str):
        return v
    if isinstance(v, dict):
        t = v.get("text")
        if isinstance(t, str):
            return t
        return json.dumps(v, ensure_ascii=False)
    if isinstance(v, list):
        parts = [_extract_text(x) for x in v]
        return "\n".join(p for p in parts if p)
    return ""


def _tool_summary(t: dict) -> dict:
    """从工具调用对象提取 {name, status, detail}。"""
    name = t.get("name") or f"tool#{t.get('toolIndex', '?')}"
    status = t.get("status") or ""
    detail = ""
    for key in ("params", "rawArgs", "arguments"):
        v = t.get(key)
        if isinstance(v, str):
            detail = v
            break
    if not detail and isinstance(t.get("rawArgs"), str):
        detail = t["rawArgs"]
    return {"name": name, "status": status, "detail": (detail or "")[:300]}


def fmt_conversation_markdown(cid: str, name: str = "") -> str:
    """把会话聊天记录渲染为 Markdown 文本（供 TUI 展示）。"""
    msgs = fetch_conversation(cid)
    if not msgs:
        return f"# {name or cid}\n\n_（无聊天记录或已无正文数据）_"
    lines = [f"# {name or cid}", f"共 {len(msgs)} 条消息\n"]
    for m in msgs:
        t = "用户" if m["type"] == 1 else "助手"
        ts = (m.get("time") or "")[:19].replace("T", " ")
        lines.append(f"## {t}  {ts}\n")
        if m["text"]:
            lines.append(m["text"] + "\n")
        if m["thinking"]:
            th = m["thinking"]
            if len(th) > 500:
                th = th[:500] + "…（已截断）"
            lines.append(f"> 💭 思考: {th}\n")
        for tool in m["tools"]:
            st = f" [{tool['status']}]" if tool["status"] else ""
            lines.append(f"> 🔧 {tool['name']}{st}\n")
            if tool["detail"]:
                d = tool["detail"]
                if len(d) > 220:
                    d = d[:220] + "…"
                lines.append(f"> `{d}`\n")
    return "\n".join(lines)


def remove_keys_for(con: sqlite3.Connection, cid: str) -> int:
    """删除一个会话的所有正文键，返回删除行数。"""
    cur = con.execute(
        "DELETE FROM cursorDiskKV WHERE key IN (?, ?) OR key LIKE ? OR key LIKE ? OR key LIKE ?",
        (f"composerData:{cid}", f"composerVirtualRowHeights:{cid}",
         f"bubbleId:{cid}:%", f"checkpointId:{cid}:%", f"ofsContent:{cid}:%"),
    )
    return cur.rowcount


def rewrite_mirror(con: sqlite3.Connection, delete_ids: Set[str]) -> Tuple[int, int]:
    """Reconcile the composer header mirror without dropping active rows."""
    row = con.execute(
        "SELECT value FROM ItemTable WHERE key='composer.composerHeaders'"
    ).fetchone()
    if not row:
        return 0, 0

    original_value = row[0]
    raw = (
        original_value.encode("utf-8", "replace")
        if isinstance(original_value, str)
        else original_value
    )
    data = json.loads(raw)
    if not isinstance(data, dict) or not isinstance(data.get("allComposers"), list):
        raise ValueError("composer.composerHeaders 的 allComposers 格式无效")

    headers = data["allComposers"]
    before = len(headers)
    table_rows = con.execute(
        "SELECT composerId, workspaceId, createdAt, lastUpdatedAt, "
        "isArchived, isSubagent, value FROM composerHeaders"
    ).fetchall()

    table_meta = {}
    for cid, workspace_id, created_at, last_updated_at, archived, subagent, value in table_rows:
        parsed = None
        if value is not None:
            try:
                raw_value = value.encode("utf-8", "replace") if isinstance(value, str) else value
                candidate = json.loads(raw_value)
                if isinstance(candidate, dict):
                    parsed = candidate
            except (TypeError, json.JSONDecodeError):
                parsed = None
        table_meta[cid] = {
            "workspace_id": workspace_id,
            "created_at": created_at,
            "last_updated_at": last_updated_at,
            "archived": bool(archived),
            "subagent": bool(subagent),
            "value": parsed,
        }

    kept = []
    seen = set()
    for header in headers:
        if not isinstance(header, dict):
            kept.append(header)
            continue
        cid = header.get("composerId")
        if cid in delete_ids:
            continue
        normalized = dict(header)
        meta = table_meta.get(cid)
        if meta is not None:
            normalized["isArchived"] = meta["archived"]
        kept.append(normalized)
        if cid:
            seen.add(cid)

    # The table can contain active conversations that the mirror has not seen
    # yet. Add them so IDE/Agents does not hide them until a later reconciliation.
    missing = []
    for cid, meta in table_meta.items():
        if cid in delete_ids or cid in seen or meta["archived"] or meta["subagent"]:
            continue
        candidate = meta["value"]
        if not isinstance(candidate, dict):
            continue
        normalized = dict(candidate)
        normalized["composerId"] = cid
        normalized["isArchived"] = False
        if "createdAt" not in normalized and meta["created_at"] is not None:
            normalized["createdAt"] = meta["created_at"]
        if "lastUpdatedAt" not in normalized and meta["last_updated_at"] is not None:
            normalized["lastUpdatedAt"] = meta["last_updated_at"]
        if "workspaceIdentifier" not in normalized and meta["workspace_id"]:
            normalized["workspaceIdentifier"] = meta["workspace_id"]
        missing.append(normalized)

    missing.sort(
        key=lambda item: (
            item.get("lastUpdatedAt") or item.get("createdAt") or 0,
            item.get("composerId", ""),
        ),
        reverse=True,
    )
    data["allComposers"] = missing + kept
    after = len(data["allComposers"])

    serialized = json.dumps(data, ensure_ascii=False, separators=(",", ":"))
    payload = serialized if isinstance(original_value, str) else serialized.encode("utf-8")
    con.execute(
        "UPDATE ItemTable SET value=? WHERE key='composer.composerHeaders'",
        (payload,),
    )
    return before, after

def backup() -> List[str]:
    """备份数据库三件套到 .bak-<时间戳>，返回生成的文件列表。"""
    ts = time.strftime("%Y%m%d-%H%M%S")
    made = []
    for p in (DB, DB + "-wal", DB + "-shm"):
        if os.path.exists(p):
            dst = f"{p}.bak-{ts}"
            shutil.copy2(p, dst)
            made.append(dst)
    return made


def list_backups() -> List[str]:
    return sorted(glob.glob(DB + ".bak-*"), key=os.path.getmtime, reverse=True)


def delete_sessions(sessions: List[Session]) -> dict:
    """
    删除指定会话：表行 + 镜像条目 + 正文键，单事务提交，之后 VACUUM。
    返回统计信息。调用方负责前置备份。
    """
    ids = {s.composer_id for s in sessions}
    con = open_db_rw()
    try:
        cur = con.cursor()
        table_del = 0
        if ids:
            placeholders = ",".join("?" * len(ids))
            cur.execute(f"DELETE FROM composerHeaders WHERE composerId IN ({placeholders})", list(ids))
            table_del = cur.rowcount
        mirror_before, mirror_after = rewrite_mirror(con, ids)
        keys_del = sum(remove_keys_for(con, cid) for cid in ids)
        con.commit()

        print("正在压缩数据库（VACUUM）…")
        cur.execute("VACUUM")
        con.commit()
    finally:
        con.close()

    return {
        "sessions": len(ids),
        "table_rows": table_del,
        "mirror": (mirror_before, mirror_after),
        "keys": keys_del,
    }


# =====================================================================
# CLI 操作
# =====================================================================

def fmt_ts(ms: Optional[int]) -> str:
    if not ms:
        return "-"
    return time.strftime("%Y-%m-%d %H:%M", time.localtime(ms / 1000))


def print_report(classes: Dict[str, List[Session]]):
    for label, key in [("已归档", S_ARCHIVED), ("镜像残留(表已删)", S_MIRROR_ONLY),
                       ("正文孤儿(表/镜像已删)", S_CONTENT_ONLY), ("未归档", S_ACTIVE)]:
        lst = classes[key]
        print(f"{label}: {len(lst)}")
        for s in lst[:5]:
            print(f"    {s.display_name:<40} 键数={s.content_keys} 更新={fmt_ts(s.last_updated)}")
        if len(lst) > 5:
            print(f"    ... 共 {len(lst)} 条")


def op_preview(_args):
    sessions = scan()
    classes = classify(sessions)
    print(f"总会话: {len(sessions)}")
    print_report(classes)


def op_delete_archived(args):
    sessions = scan()
    classes = classify(sessions)
    targets = classes[S_ARCHIVED] + classes[S_MIRROR_ONLY] + classes[S_CONTENT_ONLY]
    if not targets:
        print("没有可清理的会话。")
        return
    key_cnt = sum(s.content_keys for s in targets)
    print(f"将删除 {len(targets)} 个会话（归档 {len(classes[S_ARCHIVED])}，镜像残留 "
          f"{len(classes[S_MIRROR_ONLY])}，孤儿 {len(classes[S_CONTENT_ONLY])}），正文键 {key_cnt} 个")
    if not args.yes:
        if input("确认删除？[y/N] ").strip().lower() != "y":
            print("已取消。")
            return
    require_closed(args.force)
    made = backup()
    print(f"已备份: {', '.join(made)}")
    stats = delete_sessions(targets)
    print(f"完成: 会话 {stats['sessions']}，表行 {stats['table_rows']}，"
          f"镜像 {stats['mirror'][0]} -> {stats['mirror'][1]}，正文键 {stats['keys']}")


def op_backup(_args):
    made = backup()
    print(f"已备份 {len(made)} 个文件:")
    for p in made:
        print("  ", p)


def op_restore(args):
    baks = list_backups()
    if not baks:
        print("没有找到备份（state.vscdb.bak-*）。")
        return
    print("可用备份（新到旧）:")
    for i, p in enumerate(baks, 1):
        ts = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(os.path.getmtime(p)))
        print(f"  [{i}] {os.path.basename(p)}  ({ts})")
    if args.op:
        idx = 0
    else:
        try:
            idx = int(input("选择要恢复的备份编号: ")) - 1
        except (ValueError, EOFError):
            print("已取消。")
            return
    if not (0 <= idx < len(baks)):
        print("[err] 编号无效。")
        return
    require_closed(args.force)
    tag = os.path.basename(baks[idx]).split("bak-", 1)[1]
    group = [p for p in (DB, DB + "-wal", DB + "-shm") if os.path.exists(f"{p}.bak-{tag}")]
    if not args.yes and input("恢复会覆盖当前数据库，继续？[y/N] ").strip().lower() != "y":
        print("已取消。")
        return
    for p in group:
        shutil.copy2(f"{p}.bak-{tag}", p)
        print(f"  已恢复 {os.path.basename(p)}")
    print("恢复完成，重新打开 Cursor 生效。")


def op_wipe_all(args):
    sessions = scan()
    if not sessions:
        print("数据库为空。")
        return
    print(f"将清空全部会话（共 {len(sessions)} 条，含未归档），此操作不可逆！")
    if not args.yes and input("确认清空全部会话？[y/N] ").strip().lower() != "y":
        print("已取消。")
        return
    require_closed(args.force)
    made = backup()
    print(f"已备份: {', '.join(made)}")
    stats = delete_sessions(sessions)
    print(f"完成: 会话 {stats['sessions']}，表行 {stats['table_rows']}，"
          f"镜像 {stats['mirror'][0]} -> {stats['mirror'][1]}，正文键 {stats['keys']}")


def op_purge_index(_args):
    if not os.path.exists(SEARCH_INDEX):
        print("没有找到搜索索引，无需清理。")
        return
    require_closed(False)
    os.remove(SEARCH_INDEX)
    print(f"已删除 {SEARCH_INDEX}，Cursor 下次启动会自动重建。")


def op_repair_mirror(args):
    """修复活动会话未出现在 composer.composerHeaders 镜像中的情况。"""
    require_closed(args.force)
    made = backup()
    con = open_db_rw()
    try:
        before, after = rewrite_mirror(con, set())
        con.commit()
    except Exception:
        con.rollback()
        raise
    finally:
        con.close()
    print(f"镜像已修复: {before} -> {after} 个条目")
    print(f"备份: {', '.join(made) if made else '无'}")


def require_closed(force: bool):
    if cursor_running() and not force:
        print("[err] Cursor 正在运行，请完全退出后重试（或加 --force 强行执行，有备份兜底）")
        sys.exit(1)


OPS = [
    ("preview", "扫描并分类会话", op_preview),
    ("delete-archived", "删除归档会话+残留+孤儿", op_delete_archived),
    ("backup", "备份当前数据库", op_backup),
    ("restore", "从备份恢复", op_restore),
    ("wipe-all", "清空全部会话（危险）", op_wipe_all),
    ("purge-index", "清理会话搜索索引", op_purge_index),
    ("repair-mirror", "修复 composerHeaders 镜像", op_repair_mirror),
]


# =====================================================================
# TUI（textual）
# =====================================================================

def _tui_imports():
    try:
        from textual.app import App, ComposeResult
        from textual.binding import Binding
        from textual.containers import Horizontal, Vertical
        from textual.screen import ModalScreen, Screen
        from textual.widgets import Button, DataTable, Footer, Header, Label, MarkdownViewer, Static
        return (App, ComposeResult, Binding, Horizontal, Vertical, ModalScreen, Screen,
                Button, DataTable, Footer, Header, Label, MarkdownViewer, Static)
    except ImportError:
        return None


def run_tui() -> int:
    mods = _tui_imports()
    if mods is None:
        print("[err] TUI 需要 textual：pip install textual")
        return 1
    global TUI_APP_CLASS
    TUI_APP_CLASS = _build_app_class(mods)
    TUI_APP_CLASS().run()
    return 0


TUI_APP_CLASS = None  # 测试可引用


def _build_app_class(mods):
    (App, ComposeResult, Binding, Horizontal, Vertical, ModalScreen, Screen,
     Button, DataTable, Footer, Header, Label, MarkdownViewer, Static) = mods

    STATUS_LABEL = {
        S_ARCHIVED: "归档",
        S_ACTIVE: "未归档",
        S_MIRROR_ONLY: "残留",
        S_CONTENT_ONLY: "孤儿",
    }

    class ConfirmModal(ModalScreen[bool]):
        """删除前确认弹窗。"""

        def __init__(self, message: str):
            super().__init__()
            self._message = message

        def compose(self) -> ComposeResult:
            with Vertical(id="confirm-box"):
                yield Label(self._message, id="confirm-text")
                with Horizontal(id="confirm-btns"):
                    yield Button("确认删除", variant="error", id="btn-yes")
                    yield Button("取消", variant="primary", id="btn-no")

        def on_button_pressed(self, event: Button.Pressed) -> None:
            self.dismiss(event.button.id == "btn-yes")

    class ChatScreen(Screen[None]):
        """会话聊天记录查看器（全屏，可滚动）。"""

        BINDINGS = [
            # ChatScreen 打开后，v 不应再冒泡到 CleanerApp，避免重复
            # push_screen 导致同一条记录被层层打开、返回次数增加。
            Binding("v", "ignore_view_chat", "", show=False),
            Binding("escape", "close", "返回"),
            Binding("q", "close", "返回"),
        ]

        def __init__(self, title: str, markdown: str):
            super().__init__()
            self._title = title
            self._markdown = markdown

        def compose(self) -> ComposeResult:
            yield Header(show_clock=True)
            # Markdown 只是渲染组件，不负责滚动；MarkdownViewer 继承
            # VerticalScroll，支持 PgUp/PgDn、方向键和鼠标滚轮。
            yield MarkdownViewer(
                self._markdown,
                show_table_of_contents=False,
                id="chat-md",
            )
            yield Static(f"[q/esc]返回  ↑↓/PgUp/PgDn滚动 — {self._title}", id="chat-footer")

        def on_mount(self) -> None:
            # MarkdownViewer 的滚动键绑定在外层，实际焦点放在其文档上，
            # 这样 PgUp/PgDn 会沿 DOM 冒泡到 MarkdownViewer 处理。
            viewer = self.query_one("#chat-md", MarkdownViewer)
            viewer.document.focus()

        def action_close(self) -> None:
            self.app.pop_screen()

        def action_ignore_view_chat(self) -> None:
            """聊天记录已打开时忽略重复的 v 操作。"""
            return

    class CleanerApp(App):
        BINDINGS = [
            Binding("space", "toggle", "勾选/取消"),
            Binding("a", "select_all", "全选当前筛选"),
            Binding("n", "select_none", "取消全选"),
            Binding("v", "view_chat", "查看聊天"),
            Binding("d", "delete_selected", "删除勾选"),
            Binding("b", "do_backup", "备份"),
            Binding("q", "quit", "退出"),
        ]

        CSS = """
        #status-bar { height: 1; background: $panel; color: $text; padding: 0 1; }
        #filters { height: 3; padding: 0 1; align: left middle; }
        #filters Label { margin-right: 2; }
        #filters .filter-btn { margin-right: 1; min-width: 14; }
        .active-filter { text-style: bold; background: $accent; }
        #sessions-table { height: 1fr; }
        #foot-hint { height: 1; background: $panel; color: $text-muted; padding: 0 1; }
        #confirm-box { width: 70; height: auto; padding: 1 2; border: thick $accent;
                       background: $surface; content-align: center middle; }
        #confirm-text { text-align: center; margin-bottom: 1; }
        #confirm-btns { align: center middle; }
        #confirm-btns Button { margin: 0 1; }
        #chat-md { height: 1fr; padding: 0 1; }
        #chat-footer { height: 1; background: $panel; color: $text-muted; padding: 0 1; }
        """

        def __init__(self):
            super().__init__()
            self.sessions: List[Session] = []
            self.sess_classes: Dict[str, List[Session]] = {}
            self.filter_key: str = "all"          # all / archived / mirror-only / content-only / active
            self.selected: Set[str] = set()       # composerId
            self.current_row: Optional[str] = None
            self.row_to_id: Dict[str, str] = {}
            self.checkbox_column_key = None
            self._suppress_row_selected: Optional[str] = None

        def compose(self) -> ComposeResult:
            yield Header(show_clock=True)
            yield Static(id="status-bar")
            with Horizontal(id="filters"):
                yield Label("筛选:")
                for key, text in [("all", "全部"), (S_ARCHIVED, "归档"),
                                  (S_MIRROR_ONLY, "残留"), (S_CONTENT_ONLY, "孤儿"), (S_ACTIVE, "未归档")]:
                    yield Button(text, id=f"f-{key}")
            yield DataTable(id="sessions-table")
            yield Static(id="foot-hint")
            yield Footer()

        def on_mount(self) -> None:
            table = self.query_one(DataTable)
            table.cursor_type = "row"
            table.zebra_stripes = True
            columns = table.add_columns("☐", "状态", "标题", "正文键", "最后更新", "ID")
            self.checkbox_column_key = columns[0]
            self.refresh_data()
            self.update_status()
            self.set_filter_buttons()

        # ---- 数据 ----

        def refresh_data(self) -> None:
            try:
                self.sessions = scan()
                self.sess_classes = classify(self.sessions)
                # 清理不存在的勾选
                alive = {s.composer_id for s in self.sessions}
                self.selected &= alive
            except Exception as e:
                self.notify(f"扫描失败: {e}", severity="error", timeout=5)
                self.sessions = []
                self.sess_classes = {k: [] for k in (S_ARCHIVED, S_ACTIVE, S_MIRROR_ONLY, S_CONTENT_ONLY)}
            self.rebuild_table()

        def visible_sessions(self) -> List[Session]:
            if self.filter_key == "all":
                return self.sessions
            return self.sess_classes.get(self.filter_key, [])

        def rebuild_table(self) -> None:
            table = self.query_one(DataTable)
            table.clear()
            self.row_to_id = {}
            for s in self.visible_sessions():
                row_key = table.add_row(
                    "☑" if s.composer_id in self.selected else "☐",
                    STATUS_LABEL[s.status],
                    s.display_name,
                    str(s.content_keys),
                    fmt_ts(s.last_updated),
                    s.composer_id[:13],
                    key=s.composer_id,
                )
                self.row_to_id[row_key.value] = s.composer_id
            self.update_status()

        def update_status(self) -> None:
            running = "● 运行中" if cursor_running() else "○ 未运行"
            total = len(self.sessions)
            arch = len(self.sess_classes[S_ARCHIVED])
            mirror = len(self.sess_classes[S_MIRROR_ONLY])
            orphan = len(self.sess_classes[S_CONTENT_ONLY])
            sel = len(self.selected)
            try:
                size = os.path.getsize(DB) / 1024 / 1024
            except OSError:
                size = 0
            self.query_one("#status-bar", Static).update(
                f"Cursor: {running}  |  库 {size:.1f} MB  |  会话 {total}（归档 {arch} / 残留 {mirror} / 孤儿 {orphan}）  |  已勾选 {sel}"
            )
            self.query_one("#foot-hint", Static).update(
                "[空格]勾选  [a]全选当前  [n]取消全选  [v]查看聊天  [d]删除勾选  [b]备份  [q]退出"
            )

        def set_filter_buttons(self) -> None:
            for key, text in [("all", "全部"), (S_ARCHIVED, "归档"),
                              (S_MIRROR_ONLY, "残留"), (S_CONTENT_ONLY, "孤儿"), (S_ACTIVE, "未归档")]:
                btn = self.query_one(f"#f-{key}", Button)
                btn.classes = "filter-btn" + (" active-filter" if self.filter_key == key else "")
                cnt = len(self.sessions) if key == "all" else len(self.sess_classes.get(key, []))
                btn.label = f"{text}({cnt})"

        # ---- 事件 ----

        def on_mouse_down(self, event) -> None:
            """首列复选框支持单击切换，而不是等待 DataTable 二次点击。"""
            if event.button != 1:
                return
            table = self.query_one(DataTable)
            if event.control is not table:
                return

            meta = event.style.meta or {}
            row_index = meta.get("row", -1)
            column_index = meta.get("column", -1)
            if row_index < 0 or column_index != 0 or row_index >= table.row_count:
                return

            row_key = table.ordered_rows[row_index].key.value
            # 如果光标已经在该行，DataTable 后续的 Click 事件会再发一个
            # RowSelected；这里记下来并在 on_data_table_row_selected 中消费掉，
            # 防止一次鼠标点击发生两次切换。
            self._suppress_row_selected = (
                row_key if table.cursor_coordinate.row == row_index else None
            )
            self.action_toggle(row_key)

        def on_button_pressed(self, event: Button.Pressed) -> None:
            bid = event.button.id or ""
            if bid.startswith("f-"):
                self.filter_key = bid[2:]
                self.rebuild_table()
                self.set_filter_buttons()
            elif bid == "btn-yes":
                pass  # 由 Modal 的 dismiss 处理

        def on_data_table_row_highlighted(self, event) -> None:
            self.current_row = event.row_key.value

        def on_data_table_row_selected(self, event) -> None:
            # 优先使用事件携带的 row_key，避免 current_row 在重绘或鼠标点击
            # 过程中被另一个高亮事件覆盖。
            row_key = event.row_key.value
            if self._suppress_row_selected == row_key:
                self._suppress_row_selected = None
                return
            self.action_toggle(row_key)

        # ---- 动作 ----

        def action_toggle(self, row_key: Optional[str] = None) -> None:
            target_row = row_key or self.current_row
            if not target_row:
                return
            cid = self.row_to_id.get(target_row)
            if cid is None:
                return
            if cid in self.selected:
                self.selected.discard(cid)
            else:
                self.selected.add(cid)
            # 只更新复选框单元格，不清空并重建整张表，避免 DataTable 将
            # 光标和滚动位置重置到第一行。
            table = self.query_one(DataTable)
            if self.checkbox_column_key is not None:
                table.update_cell(
                    target_row,
                    self.checkbox_column_key,
                    "☑" if cid in self.selected else "☐",
                )
            self.update_status()

        def action_select_all(self) -> None:
            for s in self.visible_sessions():
                self.selected.add(s.composer_id)
            self.rebuild_table()

        def action_select_none(self) -> None:
            self.selected.clear()
            self.rebuild_table()

        def action_do_backup(self) -> None:
            made = backup()
            self.notify(f"已备份 {len(made)} 个文件", timeout=3)

        def action_view_chat(self) -> None:
            # v 是 CleanerApp 的全局快捷键。打开聊天记录后，按键事件仍
            # 可能到达这里；此时直接忽略，避免重复压入 ChatScreen。
            if isinstance(self.screen, ChatScreen):
                return
            if not self.current_row:
                self.notify("先选中一行再查看（↑↓ 移动高亮）", timeout=3)
                return
            cid = self.row_to_id.get(self.current_row)
            if cid is None:
                return
            sess = next((s for s in self.sessions if s.composer_id == cid), None)
            title = sess.display_name if sess else cid
            try:
                md = fmt_conversation_markdown(cid, title)
            except Exception as e:
                self.notify(f"读取聊天记录失败: {e}", severity="error", timeout=5)
                return
            self.push_screen(ChatScreen(title, md))

        def action_delete_selected(self) -> None:
            if not self.selected:
                self.notify("没有勾选任何会话", timeout=3)
                return
            selected_sessions = [s for s in self.sessions if s.composer_id in self.selected]
            arch = sum(1 for s in selected_sessions if s.status == S_ARCHIVED)
            mirror = sum(1 for s in selected_sessions if s.status == S_MIRROR_ONLY)
            orphan = sum(1 for s in selected_sessions if s.status == S_CONTENT_ONLY)
            keys = sum(s.content_keys for s in selected_sessions)
            msg = (f"将删除 {len(selected_sessions)} 个会话\n"
                   f"（归档 {arch} / 镜像残留 {mirror} / 孤儿 {orphan}），正文键 {keys} 个\n"
                   f"删除前自动备份，且要求 Cursor 已退出。")

            def _ask(result: bool) -> None:
                if not result:
                    self.notify("已取消", timeout=2)
                    return
                if cursor_running():
                    self.notify("Cursor 正在运行！请先完全退出再删除。", severity="error", timeout=5)
                    return
                made = backup()
                try:
                    stats = delete_sessions(selected_sessions)
                except Exception as e:
                    self.notify(f"删除失败: {e}", severity="error", timeout=5)
                    return
                self.selected.clear()
                self.refresh_data()
                self.notify(
                    f"完成: 会话 {stats['sessions']}，镜像 {stats['mirror'][0]}→{stats['mirror'][1]}，正文键 {stats['keys']}（备份: {os.path.basename(made[0]) if made else '无'}）",
                    timeout=6,
                )

            self.push_screen(ConfirmModal(msg), _ask)

    return CleanerApp


# =====================================================================
# 入口
# =====================================================================

def main():
    global DB, SEARCH_INDEX
    ap = argparse.ArgumentParser(description="Cursor 会话维护工具（TUI / CLI）")
    ap.add_argument("--op", choices=[n for n, _, _ in OPS], help="直接执行 CLI 操作，跳过 TUI")
    ap.add_argument("--yes", action="store_true", help="跳过确认提示（配合 --op）")
    ap.add_argument("--force", action="store_true", help="跳过 Cursor 运行检测")
    ap.add_argument("--db", help=r"指定 state.vscdb 路径（默认 %%APPDATA%%\Cursor\...；测试用）")
    args = ap.parse_args()

    if args.db:
        DB = os.path.abspath(args.db)
        SEARCH_INDEX = os.path.join(os.path.dirname(DB), "conversation-search.db")

    if args.op:
        fn = next(fn for n, _, fn in OPS if n == args.op)
        try:
            fn(args)
        except KeyboardInterrupt:
            print("\n已中断。")
        return

    sys.exit(run_tui())


if __name__ == "__main__":
    main()
