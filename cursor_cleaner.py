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
from datetime import datetime
import gzip
import glob
import json
import os
import re
import shutil
import sqlite3
import subprocess
import sys
import time
import zlib
from dataclasses import dataclass, field
from typing import Any, Dict, Iterable, List, Optional, Set, Tuple

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

GLOBAL_STORAGE = os.path.join(
    os.environ.get("APPDATA", os.path.expanduser("~")),
    "Cursor",
    "User",
    "globalStorage",
)
DEFAULT_DB = os.path.join(GLOBAL_STORAGE, "state.vscdb")
DB = DEFAULT_DB
SEARCH_INDEX = os.path.join(GLOBAL_STORAGE, "conversation-search.db")

# Cursor 的不同登录/版本会把会话分别写到 globalStorage 和
# workspaceStorage/<workspace>/state.vscdb。旧版本/API key 场景通常只会
# 命中前者，因此原来只打开 DB 的实现会漏掉账号登录后的会话。
WORKSPACE_STORAGE = os.path.join(os.path.dirname(GLOBAL_STORAGE), "workspaceStorage")
DB_DISCOVERY_ENABLED = True

# 会话列表的 ItemTable key 在 Cursor 版本之间发生过变化。保留旧 key，
# 同时兼容账号登录版本使用的 composer.composerData。
COMPOSER_MIRROR_KEYS = (
    "composer.composerHeaders",
    "composer.composerData",
)

# 与 composerId 关联的正文键前缀（key 的 UUID 段即 composerId）
SESSION_KEY_PATTERNS = ("composerData:", "composerVirtualRowHeights:", "bubbleId:", "checkpointId:", "ofsContent:")
UUID_RE = re.compile(
    r"[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}",
    re.IGNORECASE,
)

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
    created_at: Optional[int] = None
    message_count: int = 0            # composerData 中可还原的消息头数量
    is_draft: bool = False            # Cursor 创建的草稿/占位 composer
    is_subagent: bool = False         # 子 agent composer，不应作为主会话重复展示
    hidden: bool = field(default=False, repr=False, compare=False)
    # 同一个 composer 的元数据和正文可能跨 global/workspace 数据库，
    # 记录来源便于诊断，也避免 UI 只能展示 globalStorage 的结果。
    source_paths: Set[str] = field(default_factory=set, repr=False, compare=False)

    @property
    def status(self) -> str:
        if self.in_table or self.in_mirror:
            if self.table_archived or self.mirror_archived:
                return S_ARCHIVED
            if self.in_mirror and not self.in_table:
                return S_MIRROR_ONLY
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

    @property
    def has_content(self) -> bool:
        """判断该记录是否包含可展示的正文，而不是只有空 composerData。"""
        return self.message_count > 0


def _unique_paths(paths: Iterable[str]) -> List[str]:
    """按规范化绝对路径去重，同时保留传入顺序。"""
    result: List[str] = []
    seen: Set[str] = set()
    for path in paths:
        if not path:
            continue
        absolute = os.path.abspath(path)
        key = os.path.normcase(absolute)
        if key in seen:
            continue
        seen.add(key)
        result.append(absolute)
    return result


def database_paths() -> List[str]:
    """返回当前要扫描的 Cursor state.vscdb 文件。

    默认同时扫描 globalStorage 和所有 workspaceStorage。传入 --db 后由
    main() 关闭自动发现，这样测试数据库/用户指定数据库仍然是单文件模式。
    """
    paths = [DB] if os.path.isfile(DB) else []
    if DB_DISCOVERY_ENABLED and os.path.normcase(os.path.abspath(DB)) == os.path.normcase(os.path.abspath(DEFAULT_DB)):
        pattern = os.path.join(WORKSPACE_STORAGE, "*", "state.vscdb")
        paths.extend(glob.glob(pattern))
    return _unique_paths(path for path in paths if os.path.isfile(path))


def _table_exists(con: sqlite3.Connection, table: str) -> bool:
    row = con.execute(
        "SELECT 1 FROM sqlite_master WHERE type='table' AND name=? LIMIT 1",
        (table,),
    ).fetchone()
    return row is not None


def _decode_json(value: Any) -> Any:
    """解析 SQLite 中的 JSON 文本/BLOB，兼容压缩过的历史值。"""
    if isinstance(value, memoryview):
        value = value.tobytes()
    if isinstance(value, (dict, list, int, float, bool)) or value is None:
        return value

    candidates: List[Any] = [value]
    if isinstance(value, (bytes, bytearray)):
        raw = bytes(value)
        for decoder in (gzip.decompress, zlib.decompress):
            try:
                decoded = decoder(raw)
            except (OSError, zlib.error):
                continue
            candidates.append(decoded)

    for candidate in candidates:
        try:
            if isinstance(candidate, (bytes, bytearray)):
                candidate = candidate.decode("utf-8-sig", "replace")
            if not isinstance(candidate, str):
                continue
            text = candidate.strip()
            if not text:
                continue
            parsed = json.loads(text)
            # 少数版本会把 JSON 再序列化成 JSON 字符串。
            if isinstance(parsed, str) and parsed.lstrip().startswith(("{", "[")):
                try:
                    return json.loads(parsed)
                except json.JSONDecodeError:
                    pass
            return parsed
        except (TypeError, UnicodeDecodeError, json.JSONDecodeError):
            continue
    return None


def open_db_ro(path: Optional[str] = None) -> sqlite3.Connection:
    """只读连接（可读到 WAL 内未合并数据）。"""
    db_path = os.path.abspath(path or DB)
    try:
        con = sqlite3.connect(f"file:{db_path}?mode=ro", uri=True)
        # 连接对象可能延迟到第一次语句执行时才触发 Windows 锁错误。
        con.execute("SELECT 1").fetchone()
        return con
    except sqlite3.Error as normal_error:
        try:
            con.close()
        except (UnboundLocalError, sqlite3.Error):
            pass
        # Windows 下 Cursor 运行时可能持有阻止普通只读连接的锁。immutable
        # 不申请 SQLite 锁，适合扫描；它可能看不到尚未 checkpoint 的 WAL，
        # 因此只作为普通 mode=ro 失败时的最后回退。
        try:
            con = sqlite3.connect(f"file:{db_path}?mode=ro&immutable=1", uri=True)
            con.execute("SELECT 1").fetchone()
            return con
        except sqlite3.Error:
            try:
                con.close()
            except (UnboundLocalError, sqlite3.Error):
                pass
            raise normal_error


def open_db_rw(path: Optional[str] = None) -> sqlite3.Connection:
    return sqlite3.connect(os.path.abspath(path or DB))


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


def _as_bool(value: Any) -> bool:
    if isinstance(value, str):
        return value.strip().lower() in {"1", "true", "yes", "y", "on"}
    return bool(value)


def _as_int(value: Any) -> Optional[int]:
    if isinstance(value, bool):
        return int(value)
    if isinstance(value, (int, float)):
        return int(value)
    if isinstance(value, str):
        try:
            return int(float(value.strip()))
        except ValueError:
            # 某些 Cursor 版本将 lastUpdatedAt/createdAt 写成 ISO-8601。
            parsed = _as_local_datetime(value)
            return int(parsed.timestamp() * 1000) if parsed else None
    return None


def _header_id(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def _extract_composer_headers(value: Any) -> List[dict]:
    """从 ItemTable 的不同版本 JSON 中提取 composer header。"""
    obj = _decode_json(value)
    found: Dict[str, dict] = {}
    visited: Set[int] = set()

    def visit(node: Any, depth: int = 0) -> None:
        if depth > 7:
            return
        if isinstance(node, dict):
            marker = id(node)
            if marker in visited:
                return
            visited.add(marker)

            cid = _header_id(node.get("composerId") or node.get("composerID"))
            if cid:
                current = found.setdefault(cid, {})
                for key, item in node.items():
                    if item not in (None, "", [], {}):
                        current[key] = item

            # 这些字段是 Cursor 不同版本使用过的会话列表容器。
            for key in (
                "allComposers",
                "composers",
                "composerHeaders",
                "conversationHeaders",
                "items",
            ):
                child = node.get(key)
                if isinstance(child, (dict, list)):
                    visit(child, depth + 1)

            # composer.composerData 也有过再包一层 data/state 的版本。
            for key, child in node.items():
                if key in {"data", "state", "value", "payload", "result"} or "composer" in key.lower():
                    if isinstance(child, (dict, list)):
                        visit(child, depth + 1)
        elif isinstance(node, list):
            for child in node:
                if isinstance(child, (dict, list)):
                    visit(child, depth + 1)

    visit(obj)
    return list(found.values())


def _read_composer_table(con: sqlite3.Connection) -> Dict[str, dict]:
    """读取 composerHeaders；没有该表/列时返回空映射。"""
    if not _table_exists(con, "composerHeaders"):
        return {}
    try:
        rows = con.execute("SELECT * FROM composerHeaders").fetchall()
        columns = [item[1] for item in con.execute("PRAGMA table_info(composerHeaders)").fetchall()]
    except sqlite3.DatabaseError:
        return {}

    result: Dict[str, dict] = {}
    for row in rows:
        record = dict(zip(columns, row))
        cid = _header_id(record.get("composerId") or record.get("composerID"))
        if not cid:
            continue
        parsed_value = _decode_json(record.get("value"))
        if isinstance(parsed_value, dict):
            record.update({k: v for k, v in parsed_value.items() if k not in record or record[k] in (None, "")})
        result[cid] = {
            "composerId": cid,
            "isArchived": _as_bool(record.get("isArchived")),
            "name": record.get("name") or record.get("title") or "",
            "lastUpdatedAt": record.get("lastUpdatedAt") or record.get("updatedAt"),
            "createdAt": record.get("createdAt"),
            "workspaceIdentifier": record.get("workspaceIdentifier") or record.get("workspaceId") or "",
            "isDraft": _as_bool(record.get("isDraft")),
            "isSubagent": _as_bool(record.get("isSubagent")),
            "value": parsed_value if isinstance(parsed_value, dict) else None,
        }
    return result


def _read_composer_mirror(con: sqlite3.Connection) -> Dict[str, dict]:
    """读取 ItemTable 中的 composer.composerHeaders/composerData。"""
    if not _table_exists(con, "ItemTable"):
        return {}
    try:
        rows = con.execute(
            "SELECT key, value FROM ItemTable WHERE key LIKE 'composer.%'"
        ).fetchall()
    except sqlite3.DatabaseError:
        return {}

    result: Dict[str, dict] = {}
    for key, value in rows:
        # 只把已知 composer 列表或包含 composerId 的 JSON 当作镜像，
        # 避免误把其它 composer 配置项当成会话。
        if key not in COMPOSER_MIRROR_KEYS and "composer" not in str(key).lower():
            continue
        for header in _extract_composer_headers(value):
            cid = _header_id(header.get("composerId") or header.get("composerID"))
            if not cid:
                continue
            current = result.setdefault(cid, {})
            for field_name in (
                "name",
                "title",
                "lastUpdatedAt",
                "updatedAt",
                "createdAt",
                "workspaceIdentifier",
                "workspaceId",
                "isArchived",
                "isDraft",
                "isSubagent",
            ):
                candidate = header.get(field_name)
                if candidate not in (None, ""):
                    if field_name in {"isArchived", "isSubagent"}:
                        current[field_name] = _as_bool(candidate)
                    else:
                        current[field_name] = candidate
            current["composerId"] = cid
    return result


def _conversation_headers(data: Any) -> List[dict]:
    """从 composerData 中取消息头，兼容旧/新字段名。"""
    obj = _decode_json(data)
    if isinstance(obj, list):
        return [item for item in obj if isinstance(item, dict)]
    if not isinstance(obj, dict):
        return []

    for key in (
        "fullConversationHeadersOnly",
        "conversationHeadersOnly",
        "fullConversationHeaders",
        "conversationHeaders",
        "headers",
        "messages",
        "bubbles",
        "conversationMap",
    ):
        value = _decode_json(obj.get(key))
        if isinstance(value, list):
            candidates = [item for item in value if isinstance(item, dict)]
            if candidates:
                return candidates
            continue
        if isinstance(value, dict):
            # 少数版本使用 {bubbleId: header} 而不是数组。
            candidates = list(value.values())
            if candidates and all(isinstance(item, dict) for item in candidates):
                return candidates

    # 账号登录的某些版本将数据包装在 data/state/payload 内。
    for key in ("data", "state", "payload", "conversation"):
        nested = obj.get(key)
        if isinstance(nested, (dict, list)):
            headers = _conversation_headers(nested)
            if headers:
                return headers
    return []


def _bubble_id(header: dict) -> str:
    return _header_id(
        header.get("bubbleId")
        or header.get("bubbleID")
        or header.get("messageId")
        or header.get("id")
    )


def _key_identifier(key: str, prefix: str) -> str:
    if not key.startswith(prefix):
        return ""
    rest = key[len(prefix):]
    if rest.startswith(":"):
        rest = rest[1:]
    return rest.split(":", 1)[0]


def _scan_content_db(con: sqlite3.Connection) -> Tuple[Dict[str, int], Dict[str, List[Any]]]:
    """扫描一个 DB 的正文键，返回 composerId -> 键数和 composerData。"""
    if not _table_exists(con, "cursorDiskKV"):
        return {}, {}
    try:
        rows = con.execute("SELECT key, value FROM cursorDiskKV").fetchall()
    except sqlite3.DatabaseError:
        return {}, {}

    counts: Dict[str, int] = {}
    composer_data: Dict[str, List[Any]] = {}
    unscoped_bubbles: Dict[str, int] = {}

    for key, value in rows:
        key = str(key)
        if key.startswith("composerData:"):
            cid = _key_identifier(key, "composerData:")
            if cid:
                counts[cid] = counts.get(cid, 0) + 1
                composer_data.setdefault(cid, []).append(_decode_json(value))
            continue

        if key.startswith(("composerVirtualRowHeights:", "checkpointId:", "ofsContent:")):
            for prefix in ("composerVirtualRowHeights:", "checkpointId:", "ofsContent:"):
                cid = _key_identifier(key, prefix)
                if cid:
                    counts[cid] = counts.get(cid, 0) + 1
                    break
            continue

        if key.startswith("bubbleId:"):
            parts = key.split(":")
            if len(parts) >= 3:
                cid = parts[1]
                counts[cid] = counts.get(cid, 0) + 1
            elif len(parts) == 2:
                # 新版本可能只有 bubbleId:<bubbleId>，之后通过
                # fullConversationHeadersOnly 反向关联 composerId。
                bid = parts[1]
                unscoped_bubbles[bid] = unscoped_bubbles.get(bid, 0) + 1

    for cid, payloads in composer_data.items():
        for payload in payloads:
            for header in _conversation_headers(payload):
                bid = _bubble_id(header)
                if bid and bid in unscoped_bubbles:
                    counts[cid] = counts.get(cid, 0) + unscoped_bubbles[bid]

    return counts, composer_data


def _workspace_identity(value: Any) -> str:
    """将 workspaceIdentifier/workspaceId 归一化，供跨库去重使用。"""
    if isinstance(value, dict):
        for key in ("id", "uri", "path", "fsPath", "external"):
            nested = value.get(key)
            if nested not in (None, ""):
                return _workspace_identity(nested)
        return ""
    text = str(value or "").strip()
    return re.sub(r"\s+", " ", text).casefold()


def _session_group_name(value: Any) -> str:
    text = str(value or "").strip()
    return re.sub(r"\s+", " ", text).casefold()


def _deduplicate_sessions(sessions: List[Session]) -> List[Session]:
    """标记 Cursor 生成的空草稿/占位副本，保留真正有正文的会话。

    Cursor 账号登录后可能同时留下一个可读的 composer 和一个只有
    composerData 的 draft/placeholder。两者通常拥有相同标题与工作区，
    但后者没有任何消息正文。归档、镜像残留和正文孤儿仍保留，便于
    清理；只对未归档的空副本做去重。
    """
    named_groups: Dict[Tuple[str, str], List[Session]] = {}
    for session in sessions:
        name = _session_group_name(session.name)
        if name:
            key = (_workspace_identity(session.workspace_id), name)
            named_groups.setdefault(key, []).append(session)

    hidden_ids: Set[str] = set()
    for group in named_groups.values():
        has_real_content = any(session.has_content for session in group)
        if not has_real_content:
            continue
        for session in group:
            if not session.has_content:
                hidden_ids.add(session.composer_id)

    for session in sessions:
        if (
            not session.has_content
            and (session.is_draft or session.is_subagent or not _session_group_name(session.name))
        ):
            # 空的草稿、子 agent 和无标题占位记录不是可查看的对话，
            # 也是造成“一条对话显示两次”的主要来源。这里只标记隐藏，
            # 让 delete-archived 仍可在 include_hidden 模式下清理它们。
            hidden_ids.add(session.composer_id)
    for session in sessions:
        session.hidden = session.composer_id in hidden_ids
    return sessions


def scan(include_hidden: bool = False) -> List[Session]:
    """扫描 globalStorage 与 workspaceStorage，合并为会话清单。只读。"""
    paths = database_paths()
    if not paths:
        raise FileNotFoundError(f"找不到 Cursor 会话数据库：{DB} 或 {WORKSPACE_STORAGE}")

    records: Dict[str, dict] = {}
    readable_db = False
    for path in paths:
        try:
            con = open_db_ro(path)
        except sqlite3.Error:
            # Cursor 运行时某个 state.vscdb 可能被独占锁定；不要让一个
            # workspace 失败阻断其它数据库中的账号会话。
            continue
        try:
            table_map = _read_composer_table(con)
            mirror_map = _read_composer_mirror(con)
            content_map, composer_data = _scan_content_db(con)
            readable_db = True
            # bubbleId:<composerId>:... 中还会出现 empty-state-draft、
            # _recentIds 等 Cursor 内部临时标识；只有 UUID 或已有元数据
            # 的 ID 才能作为真正会话展示，避免污染列表。
            content_ids = {
                cid for cid in content_map
                if UUID_RE.fullmatch(cid) or cid in table_map or cid in mirror_map
            }
            all_ids = set(table_map) | set(mirror_map) | content_ids

            for cid in all_ids:
                record = records.setdefault(
                    cid,
                    {
                        "table_archived": False,
                        "mirror_archived": False,
                        "in_table": False,
                        "in_mirror": False,
                        "content_keys": 0,
                        "name": "",
                        "created_at": None,
                        "last_updated": None,
                        "workspace_id": "",
                        "message_count": 0,
                        "is_draft": False,
                        "is_subagent": False,
                        "source_paths": set(),
                    },
                )
                record["source_paths"].add(path)

                table = table_map.get(cid)
                if table is not None:
                    record["in_table"] = True
                    record["table_archived"] = record["table_archived"] or _as_bool(table.get("isArchived"))
                    record["name"] = record["name"] or table.get("name") or table.get("title") or ""
                    created = _as_int(table.get("createdAt"))
                    if created is not None and (record["created_at"] is None or created < record["created_at"]):
                        record["created_at"] = created
                    updated = _as_int(table.get("lastUpdatedAt") or table.get("updatedAt"))
                    if updated is not None and (record["last_updated"] is None or updated > record["last_updated"]):
                        record["last_updated"] = updated
                    record["workspace_id"] = record["workspace_id"] or table.get("workspaceIdentifier") or table.get("workspaceId") or ""
                    record["is_draft"] = record["is_draft"] or _as_bool(table.get("isDraft"))
                    record["is_subagent"] = record["is_subagent"] or _as_bool(table.get("isSubagent"))

                mirror = mirror_map.get(cid)
                if mirror is not None:
                    record["in_mirror"] = True
                    record["mirror_archived"] = record["mirror_archived"] or _as_bool(mirror.get("isArchived"))
                    record["name"] = record["name"] or mirror.get("name") or mirror.get("title") or ""
                    created = _as_int(mirror.get("createdAt"))
                    if created is not None and (record["created_at"] is None or created < record["created_at"]):
                        record["created_at"] = created
                    updated = _as_int(mirror.get("lastUpdatedAt") or mirror.get("updatedAt"))
                    if updated is not None and (record["last_updated"] is None or updated > record["last_updated"]):
                        record["last_updated"] = updated
                    record["workspace_id"] = record["workspace_id"] or mirror.get("workspaceIdentifier") or mirror.get("workspaceId") or ""
                    record["is_draft"] = record["is_draft"] or _as_bool(mirror.get("isDraft"))
                    record["is_subagent"] = record["is_subagent"] or _as_bool(mirror.get("isSubagent"))

                record["content_keys"] += int(content_map.get(cid, 0))
                payloads = composer_data.get(cid, [])
                message_count = max(
                    (len(_conversation_headers(payload)) for payload in payloads),
                    default=0,
                )
                record["message_count"] = max(record["message_count"], message_count)
                for payload in payloads:
                    if not isinstance(payload, dict):
                        continue
                    record["name"] = record["name"] or payload.get("name") or payload.get("title") or ""
                    created = _as_int(payload.get("createdAt"))
                    if created is not None and (record["created_at"] is None or created < record["created_at"]):
                        record["created_at"] = created
                    updated = _as_int(payload.get("lastUpdatedAt") or payload.get("updatedAt"))
                    if updated is not None and (record["last_updated"] is None or updated > record["last_updated"]):
                        record["last_updated"] = updated
                    record["workspace_id"] = record["workspace_id"] or payload.get("workspaceIdentifier") or ""
                    record["is_draft"] = record["is_draft"] or _as_bool(payload.get("isDraft"))
                    record["is_subagent"] = record["is_subagent"] or _as_bool(
                        payload.get("isSubagent") or payload.get("isBestOfNSubcomposer")
                    )

        except sqlite3.Error:
            # 连接建立后首次查询仍可能因 Windows 文件锁失败。
            continue
        finally:
            con.close()

    if not readable_db:
        raise PermissionError(
            "无法打开 Cursor 会话数据库。请完全退出 Cursor 后重试，"
            "或检查 state.vscdb 文件权限。"
        )

    sessions = [
        Session(composer_id=cid, **record)
        for cid, record in records.items()
    ]
    sessions = _deduplicate_sessions(sessions)
    sessions.sort(key=lambda s: (s.last_updated or 0), reverse=True)
    return sessions if include_hidden else [session for session in sessions if not session.hidden]


def classify(sessions: List[Session], include_hidden: bool = False) -> Dict[str, List[Session]]:
    out = {S_ARCHIVED: [], S_ACTIVE: [], S_MIRROR_ONLY: [], S_CONTENT_ONLY: []}
    for s in sessions:
        if s.hidden and not include_hidden:
            continue
        out[s.status].append(s)
    return out


def count_keys_for(con: sqlite3.Connection, cid: str) -> int:
    if not _table_exists(con, "cursorDiskKV"):
        return 0
    cur = con.execute(
        "SELECT COUNT(*) FROM cursorDiskKV WHERE key IN (?, ?) OR key LIKE ? OR key LIKE ? OR key LIKE ?",
        (f"composerData:{cid}", f"composerVirtualRowHeights:{cid}",
         f"bubbleId:{cid}:%", f"checkpointId:{cid}:%", f"ofsContent:{cid}:%"),
    )
    return int(cur.fetchone()[0])


def _cursor_rows(con: sqlite3.Connection) -> List[Tuple[str, Any]]:
    if not _table_exists(con, "cursorDiskKV"):
        return []
    try:
        return [(str(key), value) for key, value in con.execute("SELECT key, value FROM cursorDiskKV")]
    except sqlite3.DatabaseError:
        return []


def _item_conversation_payloads(con: sqlite3.Connection, cid: str) -> List[Any]:
    """兼容没有 cursorDiskKV 的旧版/账号登录 ItemTable 会话数据。"""
    if not _table_exists(con, "ItemTable"):
        return []
    try:
        rows = con.execute(
            "SELECT key, value FROM ItemTable WHERE key LIKE 'composer.%' OR key LIKE '%aichat%'"
        ).fetchall()
    except sqlite3.DatabaseError:
        return []

    result: List[Any] = []
    visited: Set[int] = set()

    def visit(node: Any, depth: int = 0) -> None:
        if depth > 8:
            return
        if isinstance(node, dict):
            marker = id(node)
            if marker in visited:
                return
            visited.add(marker)
            node_cid = _header_id(node.get("composerId") or node.get("composerID"))
            if node_cid == cid and _conversation_headers(node):
                result.append(node)
                return
            for key, child in node.items():
                if key in {"data", "state", "payload", "conversation", "messages", "bubbles"} or "composer" in key.lower() or "chat" in key.lower():
                    if isinstance(child, (dict, list)):
                        visit(child, depth + 1)
        elif isinstance(node, list):
            for child in node:
                if isinstance(child, (dict, list)):
                    visit(child, depth + 1)

    for key, value in rows:
        if "composer" not in str(key).lower() and "aichat" not in str(key).lower():
            continue
        visit(_decode_json(value))
    return result


def _find_bubble_value(rows: List[Tuple[str, Any]], cid: str, bid: str) -> Any:
    """查找带 composerId 或不带 composerId 的 bubbleId 键。"""
    candidates = {
        f"bubbleId:{cid}:{bid}",
        f"bubbleId:{bid}",
        f"bubble:{cid}:{bid}",
        f"bubble:{bid}",
    }
    for key, value in rows:
        if key in candidates:
            return value

    # 账号登录版本可能在 bubbleId 后追加版本/分片标记；只比较最后一段
    # 或 key 中的独立段，避免按 SQL LIKE 误匹配相似 ID。
    for key, value in rows:
        if not key.startswith(("bubbleId:", "bubble:")):
            continue
        parts = key.split(":")
        if bid in parts[1:]:
            return value
    return None


def _message_type(header: dict, payload: Any) -> int:
    value: Any = header.get("type")
    if value in (None, "") and isinstance(payload, dict):
        value = payload.get("type") or payload.get("role") or payload.get("sender")
    if isinstance(value, (int, float)):
        return 1 if int(value) == 1 else 2
    text = str(value or "").lower()
    if text in {"1", "user", "human", "request", "prompt", "user_message"}:
        return 1
    return 2


def _message_text(payload: Any) -> str:
    """从 bubble/消息对象中提取正文，兼容 text/content/消息块结构。"""
    if isinstance(payload, str):
        return payload
    if isinstance(payload, list):
        parts = [_message_text(item) for item in payload]
        return "\n".join(part for part in parts if part)
    if not isinstance(payload, dict):
        return ""

    for key in ("text", "rawText", "markdown", "value"):
        value = payload.get(key)
        if isinstance(value, str) and value:
            return value
        if isinstance(value, (dict, list)):
            text = _message_text(value)
            if text:
                return text
    for key in ("content", "message", "userMessage", "assistantMessage", "response", "result"):
        value = payload.get(key)
        if isinstance(value, (str, dict, list)):
            text = _message_text(value)
            if text:
                return text
    return ""


def _message_tools(payload: Any) -> List[dict]:
    if not isinstance(payload, dict):
        return []
    result: List[dict] = []
    for key in (
        "toolFormerData",
        "toolFormerDataList",
        "tools",
        "toolCalls",
        "tool_calls",
    ):
        value = payload.get(key)
        if isinstance(value, dict):
            result.append(_tool_summary(value))
        elif isinstance(value, list):
            result.extend(_tool_summary(item) for item in value if isinstance(item, dict))
    return result


def fetch_conversation(cid: str, paths: Optional[Iterable[str]] = None) -> List[dict]:
    """
    还原一个会话的聊天记录（按时间顺序）。
    数据源: composerData 的消息索引 + 各 bubbleId 键；globalStorage 和
    workspaceStorage 都会读取。账号登录版本常见的 bubbleId:<id>（无
    composerId）也会通过消息索引反向关联。
    返回每条消息的 dict: {type, time, text, thinking, tools}
      type: 1=用户, 2=AI
      tools: [{name, status, detail}]（工具调用摘要）
    任何一步解析失败都跳过该条，不抛异常。
    """
    db_paths = _unique_paths(paths if paths is not None else database_paths())
    if not db_paths:
        return []

    messages_by_key: Dict[str, dict] = {}
    order: Dict[str, int] = {}
    next_order = 0

    loaded: List[Tuple[List[Tuple[str, Any]], List[Any]]] = []
    for path in db_paths:
        try:
            con = open_db_ro(path)
        except sqlite3.Error:
            continue
        try:
            rows = _cursor_rows(con)
            payloads = [
                _decode_json(value)
                for key, value in rows
                if key == f"composerData:{cid}" or (
                    key.startswith("composerData:") and _key_identifier(key, "composerData:") == cid
                )
            ]
            if not payloads:
                payloads = _item_conversation_payloads(con, cid)
            loaded.append((rows, payloads))
        finally:
            con.close()

    # 不同 Cursor 版本可能把 composerData 和 bubble 正文分散到不同
    # state.vscdb；查找 bubble 时使用全部数据库的键集合。
    all_rows: List[Tuple[str, Any]] = []
    for rows, _ in loaded:
        all_rows.extend(rows)

    for rows, payloads in loaded:
            for cdata in payloads:
                headers = _conversation_headers(cdata)
                for index, header in enumerate(headers):
                    bid = _bubble_id(header)
                    bubble_value = _find_bubble_value(all_rows, cid, bid) if bid else None
                    payload = _decode_json(bubble_value) if bubble_value is not None else header
                    if payload is None:
                        payload = header
                    if isinstance(payload, dict):
                        # 某些版本将真正的 bubble 包在 data/bubble 字段中。
                        for nested_key in ("bubble", "data"):
                            nested = payload.get(nested_key)
                            if isinstance(nested, dict) and not _message_text(payload) and _message_text(nested):
                                payload = nested
                                break

                    text = _message_text(payload)
                    if not text and payload is not header:
                        text = _message_text(header)
                    mtype = _message_type(header, payload)
                    created = (
                        header.get("createdAt")
                        or header.get("timestamp")
                        or (payload.get("createdAt") if isinstance(payload, dict) else None)
                        or (payload.get("timestamp") if isinstance(payload, dict) else None)
                        or ""
                    )
                    thinking = ""
                    if isinstance(payload, dict):
                        thinking = _extract_text(
                            payload.get("thinking")
                            or payload.get("reasoning")
                            or payload.get("analysis")
                        )
                    if not thinking and isinstance(header, dict):
                        thinking = _extract_text(header.get("thinking") or header.get("reasoning"))

                    # 既有 DB 与 workspace DB 可能各存一份相同 bubble；按
                    # bubbleId 去重，但优先保留正文更完整的一份。
                    key = f"bubble:{bid}" if bid else f"message:{index}:{mtype}:{created}:{text}"
                    item = {
                        "type": mtype,
                        "time": str(created) if created is not None else "",
                        "text": text,
                        "thinking": thinking,
                        "tools": _message_tools(payload),
                    }
                    if key not in order:
                        order[key] = next_order
                        next_order += 1
                        messages_by_key[key] = item
                    else:
                        old = messages_by_key[key]
                        old_score = len(old.get("text", "")) + len(old.get("thinking", ""))
                        new_score = len(item.get("text", "")) + len(item.get("thinking", ""))
                        if new_score > old_score or len(item.get("tools", [])) > len(old.get("tools", [])):
                            messages_by_key[key] = item
    messages = [messages_by_key[key] for key in sorted(order, key=order.get)]
    return messages


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
    for key in ("params", "rawArgs", "arguments", "input", "inputSchema"):
        v = t.get(key)
        if isinstance(v, str):
            detail = v
            break
        if isinstance(v, (dict, list)):
            detail = json.dumps(v, ensure_ascii=False)
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
        ts = fmt_message_ts(m.get("time"))
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
    if not _table_exists(con, "cursorDiskKV"):
        return 0

    # 先从 composerData 取出无 composerId 的 bubbleId，账号登录版本会
    # 使用 bubbleId:<bubbleId>，不能只靠 bubbleId:<composerId>:% 删除。
    bubble_ids: Set[str] = set()
    try:
        row = con.execute(
            "SELECT value FROM cursorDiskKV WHERE key=?",
            (f"composerData:{cid}",),
        ).fetchone()
        if row:
            for header in _conversation_headers(row[0]):
                bid = _bubble_id(header)
                if bid:
                    bubble_ids.add(bid)
    except sqlite3.DatabaseError:
        pass

    keys_to_delete: List[str] = []
    try:
        for (key_value,) in con.execute("SELECT key FROM cursorDiskKV"):
            key = str(key_value)
            if (
                key == f"composerData:{cid}"
                or key.startswith(f"composerData:{cid}:")
                or key == f"composerVirtualRowHeights:{cid}"
                or key.startswith(f"composerVirtualRowHeights:{cid}:")
                or key.startswith(f"bubbleId:{cid}:")
                or key == f"checkpointId:{cid}"
                or key.startswith(f"checkpointId:{cid}:")
                or key == f"ofsContent:{cid}"
                or key.startswith(f"ofsContent:{cid}:")
            ):
                keys_to_delete.append(key)
                continue
            if bubble_ids and key.startswith(("bubbleId:", "bubble:")):
                if any(bid in key.split(":")[1:] for bid in bubble_ids):
                    keys_to_delete.append(key)
    except sqlite3.DatabaseError:
        return 0

    if not keys_to_delete:
        return 0
    placeholders = ",".join("?" for _ in keys_to_delete)
    cur = con.execute(
        f"DELETE FROM cursorDiskKV WHERE key IN ({placeholders})",
        keys_to_delete,
    )
    return int(cur.rowcount)


def rewrite_mirror(con: sqlite3.Connection, delete_ids: Set[str]) -> Tuple[int, int]:
    """更新不同 Cursor 版本的会话列表镜像，不丢失活动会话。"""
    if not _table_exists(con, "ItemTable"):
        return 0, 0
    table_map = _read_composer_table(con)
    try:
        rows = con.execute(
            "SELECT key, value FROM ItemTable WHERE key LIKE 'composer.%'"
        ).fetchall()
    except sqlite3.DatabaseError:
        return 0, 0

    total_before = 0
    total_after = 0
    changed = False

    for item_key, original_value in rows:
        if item_key not in COMPOSER_MIRROR_KEYS and "composer" not in str(item_key).lower():
            continue
        data = _decode_json(original_value)
        if not isinstance(data, dict):
            continue

        # 绝大多数版本使用 allComposers；同时兼容 composers/items 容器。
        container_key = next(
            (
                key for key in ("allComposers", "composers", "composerHeaders", "items")
                if isinstance(data.get(key), list)
            ),
            None,
        )
        if container_key is None:
            continue
        headers = data[container_key]
        before = len(headers)
        kept: List[Any] = []
        seen: Set[str] = set()
        for header in headers:
            if not isinstance(header, dict):
                kept.append(header)
                continue
            header_cid = _header_id(header.get("composerId") or header.get("composerID"))
            if header_cid in delete_ids:
                continue
            normalized = dict(header)
            meta = table_map.get(header_cid)
            if meta is not None and "isArchived" in normalized:
                normalized["isArchived"] = _as_bool(meta.get("isArchived"))
            kept.append(normalized)
            if header_cid:
                seen.add(header_cid)

        # 兼容原有 repair-mirror 行为：表中存在但镜像缺少的活动会话重新补入。
        missing: List[dict] = []
        for composer_id, meta in table_map.items():
            if composer_id in delete_ids or composer_id in seen or _as_bool(meta.get("isArchived")) or _as_bool(meta.get("isSubagent")):
                continue
            candidate = meta.get("value")
            if not isinstance(candidate, dict):
                continue
            normalized = dict(candidate)
            normalized["composerId"] = composer_id
            normalized["isArchived"] = False
            if "createdAt" not in normalized and meta.get("createdAt") is not None:
                normalized["createdAt"] = meta["createdAt"]
            if "lastUpdatedAt" not in normalized and meta.get("lastUpdatedAt") is not None:
                normalized["lastUpdatedAt"] = meta["lastUpdatedAt"]
            if "workspaceIdentifier" not in normalized and meta.get("workspaceIdentifier"):
                normalized["workspaceIdentifier"] = meta["workspaceIdentifier"]
            missing.append(normalized)

        missing.sort(
            key=lambda item: (
                _as_int(item.get("lastUpdatedAt") or item.get("createdAt")) or 0,
                item.get("composerId", ""),
            ),
            reverse=True,
        )
        data[container_key] = missing + kept
        after = len(data[container_key])
        total_before += before
        total_after += after

        serialized = json.dumps(data, ensure_ascii=False, separators=(",", ":"))
        if isinstance(original_value, str):
            payload = serialized
        elif isinstance(original_value, memoryview):
            payload = serialized.encode("utf-8")
        else:
            payload = serialized.encode("utf-8")
        con.execute(
            "UPDATE ItemTable SET value=? WHERE key=?",
            (payload, item_key),
        )
        changed = True

    if not changed:
        return 0, 0
    return total_before, total_after

def backup() -> List[str]:
    """备份所有已发现数据库的三件套到 .bak-<时间戳>。"""
    ts = time.strftime("%Y%m%d-%H%M%S")
    made = []
    for db_path in database_paths():
        for p in (db_path, db_path + "-wal", db_path + "-shm"):
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
    table_del = 0
    mirror_before = 0
    mirror_after = 0
    keys_del = 0

    for db_path in database_paths():
        try:
            con = open_db_rw(db_path)
        except sqlite3.Error:
            continue
        try:
            cur = con.cursor()
            if ids and _table_exists(con, "composerHeaders"):
                placeholders = ",".join("?" * len(ids))
                cur.execute(f"DELETE FROM composerHeaders WHERE composerId IN ({placeholders})", list(ids))
                db_table_del = max(cur.rowcount, 0)
            else:
                db_table_del = 0
            table_del += db_table_del
            before, after = rewrite_mirror(con, ids)
            mirror_before += before
            mirror_after += after
            db_keys_del = 0
            for cid in ids:
                db_keys_del += remove_keys_for(con, cid)
            keys_del += db_keys_del
            con.commit()

            # workspaceStorage 中有些 state.vscdb 没有会话表，避免无意义
            # VACUUM；失败也不影响已提交的删除。
            if db_table_del or before != after or db_keys_del:
                try:
                    print(f"正在压缩数据库（VACUUM）：{db_path}…")
                    cur.execute("VACUUM")
                    con.commit()
                except sqlite3.DatabaseError:
                    pass
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

def _as_local_datetime(value: Any) -> Optional[datetime]:
    """将 Cursor 的毫秒/秒/微秒时间戳或 ISO-8601 时间转为本地时间。"""
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
            # 无时区的旧记录按当前机器本地时间解释；带 Z/offset 的
            # Cursor 记录统一转换到当前机器本地时区。
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


def fmt_ts(value: Any) -> str:
    dt = _as_local_datetime(value)
    return dt.strftime("%Y-%m-%d %H:%M") if dt else "-"


def fmt_message_ts(value: Any) -> str:
    dt = _as_local_datetime(value)
    return dt.strftime("%Y-%m-%d %H:%M:%S") if dt else ""


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
    all_sessions = scan(include_hidden=True)
    sessions = [session for session in all_sessions if not session.hidden]
    classes = classify(sessions)
    visible_count = len(sessions)
    hidden_count = len(all_sessions) - visible_count
    suffix = f"（隐藏空占位 {hidden_count}）" if hidden_count else ""
    print(f"总会话: {visible_count}{suffix}")
    print_report(classes)


def op_delete_archived(args):
    sessions = scan(include_hidden=True)
    # 隐藏的空归档/占位记录仍应由清理命令处理。
    classes = classify(sessions, include_hidden=True)
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
        ts = fmt_message_ts(os.path.getmtime(p) * 1000)
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
    sessions = scan(include_hidden=True)
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
    """修复活动会话未出现在 composer 镜像中的情况。"""
    require_closed(args.force)
    made = backup()
    before = after = 0
    for db_path in database_paths():
        con = open_db_rw(db_path)
        try:
            db_before, db_after = rewrite_mirror(con, set())
            con.commit()
            before += db_before
            after += db_after
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
            Binding("r", "refresh", "刷新"),
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
            self._refresh_timer = None
            self._last_scan_error: Optional[str] = None

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
            # Cursor 在另一个进程中归档/取消归档时不会主动通知本 TUI；
            # 定时重新扫描，避免必须退出后再次启动脚本才能看到状态变化。
            self._refresh_timer = self.set_interval(2.0, self.refresh_data_quiet)

        # ---- 数据 ----

        def _session_signature(self, sessions: List[Session]) -> Tuple[Tuple[Any, ...], ...]:
            return tuple(sorted(
                (
                    s.composer_id,
                    s.status,
                    s.hidden,
                    s.name,
                    s.content_keys,
                    s.message_count,
                    s.last_updated,
                    s.is_draft,
                )
                for s in sessions
            ))

        def refresh_data_quiet(self) -> None:
            self.refresh_data(quiet=True)

        def refresh_data(self, quiet: bool = False, force: bool = False) -> bool:
            current_id = self.row_to_id.get(self.current_row) if self.current_row else None
            old_signature = self._session_signature(self.sessions)
            try:
                new_sessions = scan()
                new_signature = self._session_signature(new_sessions)
                if not force and new_signature == old_signature:
                    self._last_scan_error = None
                    self.update_status()
                    self.set_filter_buttons()
                    return True
                self.sessions = new_sessions
                self.sess_classes = classify(self.sessions)
                # 清理不存在的勾选
                alive = {s.composer_id for s in self.sessions}
                self.selected &= alive
                self._last_scan_error = None
            except Exception as e:
                error_text = str(e)
                if not quiet or error_text != self._last_scan_error:
                    self.notify(f"扫描失败: {error_text}", severity="error", timeout=5)
                self._last_scan_error = error_text
                # 定时刷新遇到 Cursor 短暂锁库时保留旧列表，不要闪烁成空表。
                return False
            self.rebuild_table(preserve_id=current_id)
            self.set_filter_buttons()
            return True

        def visible_sessions(self) -> List[Session]:
            if self.filter_key == "all":
                return self.all_visible_sessions()
            return self.sess_classes.get(self.filter_key, [])

        def all_visible_sessions(self) -> List[Session]:
            """返回全部可见会话，不受当前筛选按钮影响。"""
            return [s for s in self.sessions if not s.hidden]

        def rebuild_table(self, preserve_id: Optional[str] = None) -> None:
            table = self.query_one(DataTable)
            table.clear()
            self.row_to_id = {}
            preserve_row: Optional[int] = None
            preserve_row_key: Optional[str] = None
            for row_index, s in enumerate(self.visible_sessions()):
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
                if preserve_id and s.composer_id == preserve_id:
                    preserve_row = row_index
                    preserve_row_key = row_key.value
            if preserve_row is not None and table.row_count:
                table.move_cursor(row=preserve_row, scroll=False)
                self.current_row = preserve_row_key
            elif table.row_count:
                table.move_cursor(row=0, scroll=False)
                self.current_row = table.ordered_rows[0].key.value
            elif not table.row_count:
                self.current_row = None
            self.update_status()

        def update_status(self) -> None:
            running = "● 运行中" if cursor_running() else "○ 未运行"
            total = len(self.visible_sessions())
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
                "[空格]勾选  [a]全选当前  [n]取消全选  [r]刷新  [v]查看聊天  [d]删除勾选  [b]备份  [q]退出"
            )

        def set_filter_buttons(self) -> None:
            for key, text in [("all", "全部"), (S_ARCHIVED, "归档"),
                              (S_MIRROR_ONLY, "残留"), (S_CONTENT_ONLY, "孤儿"), (S_ACTIVE, "未归档")]:
                btn = self.query_one(f"#f-{key}", Button)
                btn.classes = "filter-btn" + (" active-filter" if self.filter_key == key else "")
                # “全部”按钮的数量必须来自全量可见列表，不能使用
                # visible_sessions()，否则点到“归档/残留/孤儿”后会把
                # “全部(N)”错误地改成当前筛选项的数量。
                cnt = len(self.all_visible_sessions()) if key == "all" else len(self.sess_classes.get(key, []))
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

        def action_refresh(self) -> None:
            if self.refresh_data(force=True):
                self.notify("已刷新会话列表", timeout=2)

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
    global DB, SEARCH_INDEX, DB_DISCOVERY_ENABLED
    ap = argparse.ArgumentParser(description="Cursor 会话维护工具（TUI / CLI）")
    ap.add_argument("--op", choices=[n for n, _, _ in OPS], help="直接执行 CLI 操作，跳过 TUI")
    ap.add_argument("--yes", action="store_true", help="跳过确认提示（配合 --op）")
    ap.add_argument("--force", action="store_true", help="跳过 Cursor 运行检测")
    ap.add_argument("--db", help=r"指定 state.vscdb 路径（默认 %%APPDATA%%\Cursor\...；测试用）")
    args = ap.parse_args()

    if args.db:
        DB = os.path.abspath(args.db)
        SEARCH_INDEX = os.path.join(os.path.dirname(DB), "conversation-search.db")
        DB_DISCOVERY_ENABLED = False

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
