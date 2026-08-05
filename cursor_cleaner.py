# -*- coding: utf-8 -*-
"""
Cursor 会话清理工具（TUI + CLI）

TUI 用法:
    python cursor_cleaner.py

CLI 用法（自动化/测试）:
    python cursor_cleaner.py --op preview
    python cursor_cleaner.py --op delete-archived --yes
    python cursor_cleaner.py --op backup-sessions --ids <id1>,<id2>
    python cursor_cleaner.py --op restore-sessions --file <备份.json>

操作列表:
    preview         扫描并分类会话（归档/残留/孤儿/未归档）
    delete-archived 删除已归档会话 + 镜像残留 + 正文孤儿（核心功能）
    backup-sessions 备份指定会话为 JSON 存档（--ids 逗号分隔）
    restore-sessions 从 JSON 存档恢复会话（--file 指定，否则列表选择）
    wipe-all        清空全部会话（危险）
    purge-index     清理会话搜索索引 conversation-search.db
    repair-mirror   修复 composerHeaders 镜像

数据模型（state.vscdb）:
    - composerHeaders: 每行一个会话，isArchived=1 表示已归档
    - ItemTable['composer.composerHeaders']: 侧边栏镜像列表（含 isArchived）
    - cursorDiskKV: 会话正文（composerData:<id> / bubbleId:<id>:* / ...）

会话状态分类:
    ARCHIVED      表或镜像标记为已归档
    ACTIVE        表或镜像存在且未归档
    MIRROR_ONLY   仅在镜像中且无正文（空壳残留）
    CONTENT_ONLY  仅在正文中（表/镜像都已删，孤儿数据）
"""

import argparse
import base64
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
                # 新版 Cursor 的会话列表只写入 ItemTable 镜像，
                # composerHeaders 表已停止更新，因此“镜像有、表没有”
                # 并不代表表行被删过。只有没有任何正文键的空壳
                # （例如中断的删除过程遗留）才算残留，其余视为活跃。
                if self.content_keys:
                    return S_ACTIVE
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


_cursor_running_cache: Tuple[float, bool] = (0.0, False)


def cursor_running_cached(ttl: float = 5.0) -> bool:
    """带 TTL 缓存的 Cursor 运行检测，避免高频 UI 刷新时反复 spawn 子进程。

    TUI 状态栏/勾选等展示用途使用；删除、恢复等写操作前的安全检查
    仍调用无缓存的 cursor_running()，不做任何行为回退。
    """
    global _cursor_running_cache
    now = time.monotonic()
    if now - _cursor_running_cache[0] >= ttl:
        _cursor_running_cache = (now, cursor_running())
    return _cursor_running_cache[1]


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
    """扫描一个 DB 的正文键，返回 composerId -> 键数和 composerData。

    列表阶段只读键名（短字符串），不拉取 cursorDiskKV 的大 blob；
    composerData 的值单独按 key 精确读取，用于解析消息头。
    """
    if not _table_exists(con, "cursorDiskKV"):
        return {}, {}
    try:
        keys = [str(key) for (key,) in con.execute("SELECT key FROM cursorDiskKV")]
    except sqlite3.DatabaseError:
        return {}, {}

    counts: Dict[str, int] = {}
    composer_data_keys: Dict[str, List[str]] = {}
    unscoped_bubbles: Dict[str, int] = {}

    for key in keys:
        if key.startswith("composerData:"):
            cid = _key_identifier(key, "composerData:")
            if cid:
                counts[cid] = counts.get(cid, 0) + 1
                composer_data_keys.setdefault(cid, []).append(key)
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

    composer_data: Dict[str, List[Any]] = {}
    if composer_data_keys:
        try:
            for cid, key_list in composer_data_keys.items():
                placeholders = ",".join("?" * len(key_list))
                rows = con.execute(
                    f"SELECT key, value FROM cursorDiskKV WHERE key IN ({placeholders})",
                    key_list,
                ).fetchall()
                composer_data[cid] = [_decode_json(value) for _, value in rows]
        except sqlite3.DatabaseError:
            pass

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


def _bubble_key_index(con: sqlite3.Connection) -> Dict[str, str]:
    """读取一个 DB 的 bubble 键名（不取 value，避免拉大 blob）。

    返回 bid -> 完整键 的映射；兼容 bubbleId:<cid>:<bid>、bubbleId:<bid>
    及旧版 bubble:<cid>:<bid> / bubble:<bid> 写法。同一 bid 存在多种
    写法时，无 composerId 的（账号登录版本）优先；带版本/分片后缀的
    键（bid 在非末段）按独立段登记，供精确取值兜底。
    """
    if not _table_exists(con, "cursorDiskKV"):
        return {}
    try:
        rows = con.execute(
            "SELECT key FROM cursorDiskKV WHERE key LIKE 'bubbleId:%' OR key LIKE 'bubble:%'"
        ).fetchall()
    except sqlite3.DatabaseError:
        return {}

    result: Dict[str, str] = {}
    scoped_keys: List[str] = []
    for (key,) in rows:
        key = str(key)
        parts = key.split(":")
        if len(parts) >= 3:
            scoped_keys.append(key)
        elif len(parts) == 2:
            result[parts[1]] = key
    # 无 composerId 写法优先；带 composerId 的写法登记各独立段兜底。
    for key in scoped_keys:
        parts = key.split(":")
        for segment in parts[1:]:
            result.setdefault(segment, key)
    return result


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

    # 阶段一：每个库只读本会话的 composerData 值（精确键，不碰全表大
    # blob）和全部 bubble 键名，构造 bid -> 完整键 索引。
    loaded: List[Tuple[str, Dict[str, str], List[Any]]] = []
    for path in db_paths:
        try:
            con = open_db_ro(path)
        except sqlite3.Error:
            continue
        try:
            key_index = _bubble_key_index(con)
            payloads: List[Any] = []
            try:
                rows = con.execute(
                    "SELECT key, value FROM cursorDiskKV WHERE key = ? OR key LIKE ?",
                    (f"composerData:{cid}", f"composerData:{cid}:%"),
                ).fetchall()
                for key, value in rows:
                    payloads.append(_decode_json(value))
            except sqlite3.DatabaseError:
                pass
            if not payloads:
                payloads = _item_conversation_payloads(con, cid)
            loaded.append((path, key_index, payloads))
        finally:
            con.close()

    # 不同 Cursor 版本可能把 composerData 和 bubble 正文分散到不同
    # state.vscdb；跨库合并 bid 索引，先出现的库优先（与原全表行顺序
    # 语义一致）。
    all_keys: Dict[str, str] = {}
    for _, key_index, _ in loaded:
        for bid, key in key_index.items():
            all_keys.setdefault(bid, key)

    # 解析消息头，收集实际需要的 bubble 键集合。
    wanted_keys: Set[str] = set()
    for _, _, payloads in loaded:
        for cdata in payloads:
            headers = _conversation_headers(cdata)
            for header in headers:
                bid = _bubble_id(header)
                key = all_keys.get(bid) if bid else None
                if key:
                    wanted_keys.add(key)

    # 阶段二：按库用 WHERE key IN (...) 分批精确取值，替代原先对
    # 全表行集合的线性扫描（O(消息数 × 总键数) -> O(总键数 + 消息数)）。
    bubble_values: Dict[str, Any] = {}
    for path, key_index, _ in loaded:
        if not wanted_keys:
            break
        db_keys = sorted(key for key in key_index.values() if key in wanted_keys)
        if not db_keys:
            continue
        try:
            con = open_db_ro(path)
        except sqlite3.Error:
            continue
        try:
            for i in range(0, len(db_keys), 500):
                batch = db_keys[i:i + 500]
                placeholders = ",".join("?" * len(batch))
                rows = con.execute(
                    f"SELECT key, value FROM cursorDiskKV WHERE key IN ({placeholders})",
                    batch,
                ).fetchall()
                for key, value in rows:
                    bubble_values.setdefault(str(key), value)
        except sqlite3.DatabaseError:
            pass
        finally:
            con.close()

    for _, _, payloads in loaded:
        for cdata in payloads:
            headers = _conversation_headers(cdata)
            for index, header in enumerate(headers):
                bid = _bubble_id(header)
                bubble_key = all_keys.get(bid) if bid else None
                bubble_value = bubble_values.get(bubble_key) if bubble_key else None
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


def _user_message_summary(m: dict) -> str:
    """用户消息的简略摘要：取正文首行截断；无正文时退回工具名。"""
    text = (m.get("text") or "").strip()
    if text:
        first = text.splitlines()[0].strip()
        return first[:80] + "…" if len(first) > 80 else first
    for tool in m.get("tools") or []:
        return f"🔧 {tool['name']}"
    return "（无正文）"


# =====================================================================
# 聊天记录虚拟化布局（纯函数，供 ChatLog 组件使用）
#
# 思路：长会话不再渲染成几百个 widget 节点（MarkdownViewer 的 990 个
# 节点导致打开冻结 5.5s、滚动 9fps），而是把消息预排版为「固定宽度下
# 的文本行列表」，滚动视图只渲染视口内的几十行。
# =====================================================================

_INLINE_MD = None  # 惰性初始化的 MarkdownIt 实例（worker 线程内使用）


def _inline_md():
    """行内 markdown 解析器（惰性创建，TUI 与 CLI 共用）。"""
    global _INLINE_MD
    if _INLINE_MD is None:
        from markdown_it import MarkdownIt

        _INLINE_MD = MarkdownIt("commonmark", {"html": False, "breaks": False, "linkify": False})
    return _INLINE_MD


def _tokens_to_text(tokens) -> str:
    """把行内 markdown token 树转回纯文本，仅保留粗体/斜体/行内代码标记。

    返回的是带轻量样式的文本；[b]/[i]/[r] 标记由 Rich 的 markup 语法
    解释（与 TUI 状态栏等处的写法一致）。
    """
    parts: List[str] = []
    for tok in tokens:
        if tok.type == "text":
            parts.append(tok.content)
        elif tok.type == "softbreak" or tok.type == "hardbreak":
            parts.append("\n")
        elif tok.type == "code_inline":
            parts.append(f"[r]{tok.content}[/]")
        elif tok.type == "strong_open":
            parts.append("[b]")
        elif tok.type == "strong_close":
            parts.append("[/b]")
        elif tok.type == "em_open":
            parts.append("[i]")
        elif tok.type == "em_close":
            parts.append("[/i]")
        elif tok.type == "s_open":
            parts.append("[s]")
        elif tok.type == "s_close":
            parts.append("[/s]")
        elif tok.type == "link_open":
            # 链接只保留文字，丢弃 URL（终端里 URL 太长且不可点）
            pass
        elif tok.type == "link_close":
            pass
        elif tok.type == "image":
            parts.append(tok.content or "（图片）")
        elif tok.type == "html_inline":
            parts.append(tok.content)
        else:
            content = getattr(tok, "content", "") or ""
            if content:
                parts.append(content)
    return "".join(parts)


def _md_escape(text: str) -> str:
    """把普通文本转成 Rich markup 安全形式（避免 [ 被当作样式解析）。"""
    return text.replace("[", "[[")  # Rich 中 [[ 表示字面 [


@dataclass
class ChatMessageLayout:
    """一条消息排版后的行区间。

    start_line/end_line 为 [start, end) 左闭右开，end_line - start_line
    即该消息占用的总行数。
    """
    start_line: int = 0
    end_line: int = 0
    is_user: bool = False
    time: str = ""
    summary: str = ""


@dataclass
class ChatLayout:
    """整个会话的排版结果：固定宽度下的行列表 + 消息行区间索引。"""
    lines: List[str] = field(default_factory=list)     # 每行是带 Rich markup 的文本
    msg_layouts: List[ChatMessageLayout] = field(default_factory=list)
    total_lines: int = 0
    width: int = 0
    user_msg_indices: List[int] = field(default_factory=list)  # 用户消息在 msg_layouts 中的下标（缓存）

    def user_indices(self) -> List[int]:
        """用户消息在 msg_layouts 中的下标（与 UserIndexRail 一一对应）。"""
        return self.user_msg_indices

    def line_of_msg(self, index: int) -> int:
        """返回第 index 条消息的起始行号（圆点跳转目标）。"""
        if 0 <= index < len(self.msg_layouts):
            return self.msg_layouts[index].start_line
        return 0

    def dot_of_msg(self, msg_index: int) -> int:
        """消息下标 -> 圆点序号；msg_index 不是用户消息时返回 -1。"""
        if msg_index < 0:
            return -1
        for dot, idx in enumerate(self.user_msg_indices):
            if idx == msg_index:
                return dot
        return -1

    def user_index_at(self, top: int, height: int) -> int:
        """返回视口 [top, top+height) 内第一条用户消息的消息下标。

        视口内没有用户消息顶边时，回退到最近一条已滚过的用户消息；
        没有任何用户消息返回 -1。二分查找，O(log n)。
        """
        users = self.user_msg_indices
        if not users:
            return -1
        # 二分：第一个 start_line >= top 的用户消息
        lo, hi = 0, len(users)
        while lo < hi:
            mid = (lo + hi) // 2
            if self.msg_layouts[users[mid]].start_line < top:
                lo = mid + 1
            else:
                hi = mid
        if lo < len(users) and self.msg_layouts[users[lo]].start_line < top + height:
            return users[lo]
        # 回退：视口上方的最后一条用户消息
        return users[lo - 1] if lo > 0 else -1


def _wrap_line(text: str, width: int) -> List[str]:
    """按终端单元格宽度折行（中文字符占 2 格），返回物理行列表。

    使用 rich.cells.cell_len 计算宽度；不做单词断行保护（代码/长 token
    场景下逐字符硬折更符合终端预期）。文本内含 \\n 时按行拆分。
    """
    from rich.cells import cell_len

    if not text:
        return [""]
    rows: List[str] = []
    for segment in text.split("\n"):
        # 预先计算每个字符的宽度，避免 cell_len 在循环里反复查表
        chars = list(segment)
        widths = [cell_len(ch) for ch in chars]
        cur: List[str] = []
        cur_w = 0
        for ch, w in zip(chars, widths):
            if w == 0:
                # 零宽字符（组合字符等）直接并入当前行，不占格
                if cur:
                    cur.append(ch)
                continue
            if cur_w + w > width:
                if cur:
                    rows.append("".join(cur))
                cur = [ch]
                cur_w = w
            else:
                cur.append(ch)
                cur_w += w
        if cur:
            rows.append("".join(cur))
        elif not rows:
            rows.append("")
    return rows or [""]


def _parse_inline_md(line: str) -> str:
    """单行文本的行内 markdown 轻量渲染（返回带 markup 的文本）。

    只处理粗体/斜体/行内代码/删除线；fence、标题、列表等块级语法
    一律按普通文本对待（消息正文本来就是富文本字符串，不是 md 文档）。
    """
    md = _inline_md()
    tokens = md.parseInline(line)
    return _tokens_to_text(tokens)


def build_chat_layout(msgs: List[dict], width: int, title: str = "") -> ChatLayout:
    """把消息列表排版为固定宽度下的行列表。

    width 为内容区宽度（不含边框/内边距）。标题行、思考引用、工具调用
    均与旧 markdown 视图的样式保持一致，只是从"整段 markdown 解析"
    改为"逐行预排版 + 行内标记"。

    纯函数，可在 worker 线程执行。
    """
    content_width = max(8, width)
    layout = ChatLayout(width=content_width)
    lines: List[str] = []
    msg_layouts: List[ChatMessageLayout] = []

    def emit(lines_out: List[str], text: str, prefix: str = "") -> None:
        """按行写入 lines_out：文本折行（纯文本，无 markup）。"""
        text = text.replace("\r\n", "\n").replace("\r", "\n")
        for raw in text.split("\n"):
            for row in _wrap_line(raw, content_width):
                if prefix:
                    lines_out.append(prefix + row[: content_width - len(prefix)])
                else:
                    lines_out.append(row)

    def append_block(lines_out: List[str], text: str, prefix: str = "") -> None:
        """带行内 md 渲染的写入（正文/思考用）。

        先按纯文本折行、再逐物理行做行内 markdown 渲染，避免 markup
        标记（[b]/[i]/[r]）跨折行边界被截断。
        """
        text = text.replace("\r\n", "\n").replace("\r", "\n")
        for raw in text.split("\n"):
            for row in _wrap_line(raw, content_width):
                rendered = _parse_inline_md(row)
                if prefix:
                    lines_out.append(prefix + rendered[: content_width - len(prefix)])
                else:
                    lines_out.append(rendered)

    if not msgs:
        lines.append(f"[b]{_md_escape(title or '会话')}[/]")
        lines.append("")
        lines.append("_（无聊天记录或已无正文数据）_")
        layout.lines = lines
        layout.total_lines = len(lines)
        return layout

    lines.append(f"[b]{_md_escape(title or '会话')}[/]")
    lines.append(f"共 {len(msgs)} 条消息")
    lines.append("")

    for m in msgs:
        start = len(lines)
        is_user = m.get("type") == 1
        who = "用户" if is_user else "助手"
        ts = fmt_message_ts(m.get("time"))
        # 标题行带强调色，与旧视图的 ## 标题一致
        lines.append(f"[b]{_md_escape(who)}[/]  {_md_escape(ts)}")
        lines.append("")
        if m.get("text"):
            append_block(lines, m["text"])
            lines.append("")
        thinking = m.get("thinking") or ""
        if thinking:
            if len(thinking) > 500:
                thinking = thinking[:500] + "…（已截断）"
            lines.append("[dim]💭 思考:[/]")
            append_block(lines, thinking)
            lines.append("")
        for tool in m.get("tools") or []:
            st = f" [{tool['status']}]" if tool.get("status") else ""
            lines.append(f"[dim]🔧 {_md_escape(tool['name'])}{_md_escape(st)}[/]")
            detail = tool.get("detail") or ""
            if detail:
                if len(detail) > 220:
                    detail = detail[:220] + "…"
                append_block(lines, detail)
            lines.append("")

        end = len(lines)
        msg_layouts.append(
            ChatMessageLayout(
                start_line=start,
                end_line=end,
                is_user=is_user,
                time=ts,
                summary=_user_message_summary(m),
            )
        )

    layout.lines = lines
    layout.msg_layouts = msg_layouts
    layout.total_lines = len(lines)
    layout.user_msg_indices = [i for i, m in enumerate(msg_layouts) if m.is_user]
    return layout


def fmt_conversation_markdown(cid: str, name: str = "",
                              msgs: Optional[List[dict]] = None) -> str:
    """把会话聊天记录渲染为 Markdown 文本（供 TUI 展示）。

    msgs 可传入已拉取的消息列表，避免调用方重复读取数据库。
    """
    if msgs is None:
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
            # 镜像的 isArchived 是权威来源：新版 Cursor 归档只写 ItemTable
            # 镜像，composerHeaders 表的 isArchived 已停止更新、可能过时。
            # 只用表行在镜像缺失该字段时兜底，绝不能反过来覆盖镜像值，
            # 否则删除一个会话会把其余"表 0 + 镜像 1"的归档会话全部
            # 误改成未归档（反之亦然）。
            if meta is not None and "isArchived" not in normalized:
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

# ---- 会话级备份（JSON 存档） ----
#
# 一个会话的数据可能落在多个 state.vscdb，且每个库里分散在
# composerHeaders 表行、ItemTable 镜像 JSON、cursorDiskKV 正文键三处。
# 会话级备份按 composerId 收集这三处的原始数据，导出为 JSON 存档；
# 恢复时按原库路径写回，已存在的行/键/镜像条目一律跳过，绝不覆盖
# 现有数据（避免旧备份覆盖新数据）。

SESSION_BACKUP_VERSION = 1


def _encode_backup_value(value: Any) -> Any:
    """SQLite 值转 JSON 可序列化结构：str/数字/None 直存，bytes 标记类型。

    备份文件里必须保留 TEXT/BLOB 的区别，否则恢复时无法原样写回。
    """
    if isinstance(value, memoryview):
        value = value.tobytes()
    if isinstance(value, (bytes, bytearray)):
        return {"t": "b", "v": base64.b64encode(bytes(value)).decode("ascii")}
    return {"t": "s", "v": value}


def _decode_backup_value(obj: Any) -> Any:
    """_encode_backup_value 的逆操作。"""
    if isinstance(obj, dict) and obj.get("t") in ("s", "b"):
        if obj["t"] == "b":
            try:
                return base64.b64decode(obj["v"])
            except (TypeError, ValueError):
                return obj.get("v")
        return obj.get("v")
    return obj


def _collect_session_kv(con: sqlite3.Connection, cid: str) -> List[Tuple[str, Any]]:
    """收集一个库中某会话的全部 cursorDiskKV 键值。

    键匹配规则与 remove_keys_for() 对称：直接前缀匹配 + 从
    composerData 消息头反向关联无 composerId 的 bubbleId:<bid>。
    返回 [(key, value), ...]，value 为已解码的原始 SQLite 值。
    """
    if not _table_exists(con, "cursorDiskKV"):
        return []

    # 先解析 composerData 拿到无 composerId 的 bubbleId，与删除逻辑一致。
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

    keys: List[str] = []
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
                keys.append(key)
                continue
            if bubble_ids and key.startswith(("bubbleId:", "bubble:")):
                if any(bid in key.split(":")[1:] for bid in bubble_ids):
                    keys.append(key)
    except sqlite3.DatabaseError:
        return []

    if not keys:
        return []
    placeholders = ",".join("?" for _ in keys)
    try:
        rows = con.execute(
            f"SELECT key, value FROM cursorDiskKV WHERE key IN ({placeholders})",
            keys,
        ).fetchall()
    except sqlite3.DatabaseError:
        return []
    return [(str(key), value) for key, value in rows]


def _collect_mirror_headers(con: sqlite3.Connection, cid: str) -> List[dict]:
    """从 ItemTable 镜像中提取某会话的 header 对象。

    返回 [{item_key, header}]，header 为镜像 JSON 中该 cid 的原始
    字典对象，恢复时按原 item_key 原样放回。
    """
    if not _table_exists(con, "ItemTable"):
        return []
    try:
        rows = con.execute(
            "SELECT key, value FROM ItemTable WHERE key LIKE 'composer.%'"
        ).fetchall()
    except sqlite3.DatabaseError:
        return []

    found: List[dict] = []
    visited: Set[int] = set()

    def visit(node: Any, item_key: str, depth: int = 0) -> None:
        if depth > 8:
            return
        if isinstance(node, dict):
            marker = id(node)
            if marker in visited:
                return
            visited.add(marker)
            node_cid = _header_id(node.get("composerId") or node.get("composerID"))
            if node_cid == cid:
                found.append({"item_key": item_key, "header": node})
                return
            for key, child in node.items():
                if isinstance(child, (dict, list)):
                    visit(child, item_key, depth + 1)
        elif isinstance(node, list):
            for child in node:
                if isinstance(child, (dict, list)):
                    visit(child, item_key, depth + 1)

    for item_key, value in rows:
        item_key = str(item_key)
        if item_key not in COMPOSER_MIRROR_KEYS and "composer" not in item_key.lower():
            continue
        visited.clear()
        visit(_decode_json(value), item_key)
    return found


def collect_session_records(sessions: List[Session]) -> List[dict]:
    """跨库收集勾选会话的完整原始数据，供备份导出。

    返回 [{composer_id, db, table_row, mirror_headers, kv}]：
      db             键值所在的库文件路径（恢复时按此写回）
      table_row      该库 composerHeaders 中的原始行（列名 -> 编码后值），
                     无行时为 None
      mirror_headers 该库 ItemTable 镜像中的原始 header 列表
      kv             cursorDiskKV 中的 [(key, 编码后值), ...]
    """
    ids = {s.composer_id for s in sessions}
    if not ids:
        return []
    records: List[dict] = []
    for db_path in database_paths():
        try:
            con = open_db_ro(db_path)
        except sqlite3.Error:
            continue
        try:
            for cid in ids:
                record: dict = {
                    "composer_id": cid,
                    "db": db_path,
                    "table_row": None,
                    "mirror_headers": [],
                    "kv": [],
                }
                # composerHeaders 原始行
                if _table_exists(con, "composerHeaders"):
                    try:
                        columns = [item[1] for item in con.execute("PRAGMA table_info(composerHeaders)").fetchall()]
                        row = con.execute(
                            "SELECT * FROM composerHeaders WHERE composerId=?",
                            (cid,),
                        ).fetchone()
                        if row is not None:
                            record["table_row"] = {
                                col: _encode_backup_value(val)
                                for col, val in zip(columns, row)
                            }
                    except sqlite3.DatabaseError:
                        pass
                record["mirror_headers"] = _collect_mirror_headers(con, cid)
                record["kv"] = [
                    (key, _encode_backup_value(value))
                    for key, value in _collect_session_kv(con, cid)
                ]
                if record["table_row"] is None and not record["mirror_headers"] and not record["kv"]:
                    continue
                records.append(record)
        except sqlite3.DatabaseError:
            continue
        finally:
            con.close()
    return records


def backup_sessions(sessions: List[Session]) -> str:
    """把勾选会话导出为 JSON 存档文件，返回存档路径。

    存档放在全局库所在目录（默认 globalStorage），文件名
    sessions-backup-<时间戳>.json；只读操作，不需要 Cursor 退出。
    """
    data = {
        "version": SESSION_BACKUP_VERSION,
        "created_at": time.strftime("%Y-%m-%d %H:%M:%S"),
        "sessions": collect_session_records(sessions),
    }
    ts = time.strftime("%Y%m%d-%H%M%S")
    out_path = os.path.join(os.path.dirname(DB), f"sessions-backup-{ts}.json")
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False)
    return out_path


def list_session_backups() -> List[str]:
    """列出全部会话级备份存档（新到旧）。"""
    pattern = os.path.join(os.path.dirname(DB), "sessions-backup-*.json")
    return sorted(glob.glob(pattern), key=os.path.getmtime, reverse=True)


def _restore_table_row(con: sqlite3.Connection, cid: str, table_row: dict) -> bool:
    """写回 composerHeaders 行；已存在则跳过。返回是否新增。"""
    if not _table_exists(con, "composerHeaders"):
        return False
    existing = con.execute(
        "SELECT 1 FROM composerHeaders WHERE composerId=? LIMIT 1",
        (cid,),
    ).fetchone()
    if existing:
        return False
    columns = [item[1] for item in con.execute("PRAGMA table_info(composerHeaders)").fetchall()]
    if not columns:
        return False
    values = [_decode_backup_value(table_row.get(col)) for col in columns]
    placeholders = ",".join("?" * len(columns))
    con.execute(
        f"INSERT INTO composerHeaders ({','.join(columns)}) VALUES ({placeholders})",
        values,
    )
    return True


def _restore_mirror_headers(con: sqlite3.Connection, headers: List[dict]) -> int:
    """把备份的镜像 header 放回原 ItemTable 条目；已存在则跳过。

    只做"容器内 append 缺失的 cid"，不改动容器结构与排序。返回新增数。
    """
    if not headers or not _table_exists(con, "ItemTable"):
        return 0
    added = 0
    for item in headers:
        item_key = str(item.get("item_key") or "")
        header = item.get("header")
        if not item_key or not isinstance(header, dict):
            continue
        cid = _header_id(header.get("composerId") or header.get("composerID"))
        if not cid:
            continue
        row = con.execute(
            "SELECT value FROM ItemTable WHERE key=?", (item_key,)
        ).fetchone()
        if row is None:
            continue
        data = _decode_json(row[0])
        if not isinstance(data, dict):
            continue
        container_key = next(
            (
                key for key in ("allComposers", "composers", "composerHeaders", "items")
                if isinstance(data.get(key), list)
            ),
            None,
        )
        if container_key is None:
            continue
        exists = any(
            isinstance(h, dict) and _header_id(h.get("composerId") or h.get("composerID")) == cid
            for h in data[container_key]
        )
        if exists:
            continue
        data[container_key].append(dict(header))
        serialized = json.dumps(data, ensure_ascii=False, separators=(",", ":"))
        original_value = row[0]
        if isinstance(original_value, str):
            payload = serialized
        else:
            payload = serialized.encode("utf-8")
        con.execute(
            "UPDATE ItemTable SET value=? WHERE key=?",
            (payload, item_key),
        )
        added += 1
    return added


def restore_sessions(backup_path: str) -> dict:
    """从 JSON 存档恢复会话数据，返回统计信息。

    逐库写回 composerHeaders 行、ItemTable 镜像 header、cursorDiskKV
    键值；已存在的数据一律跳过（不覆盖）。调用方负责 Cursor 退出检查。
    """
    with open(backup_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    sessions_done: Set[str] = set()
    table_rows = 0
    mirror_added = 0
    kv_added = 0
    skipped = 0

    # 按库分组，一次连接写一个库，避免反复开关连接。
    by_db: Dict[str, List[dict]] = {}
    for record in data.get("sessions", []):
        by_db.setdefault(record.get("db", ""), []).append(record)

    for db_path, records in by_db.items():
        if not os.path.isfile(db_path):
            skipped += len(records)
            continue
        try:
            con = open_db_rw(db_path)
        except sqlite3.Error:
            skipped += len(records)
            continue
        try:
            for record in records:
                cid = str(record.get("composer_id") or "")
                if not cid:
                    skipped += 1
                    continue
                try:
                    table_row = record.get("table_row")
                    if isinstance(table_row, dict) and _restore_table_row(con, cid, table_row):
                        table_rows += 1
                    mirror_added += _restore_mirror_headers(con, record.get("mirror_headers") or [])
                    for key, encoded in record.get("kv") or []:
                        if not isinstance(key, str) or not isinstance(encoded, dict):
                            continue
                        existing = con.execute(
                            "SELECT 1 FROM cursorDiskKV WHERE key=? LIMIT 1",
                            (key,),
                        ).fetchone()
                        if existing:
                            skipped += 1
                            continue
                        con.execute(
                            "INSERT INTO cursorDiskKV (key, value) VALUES (?, ?)",
                            (key, _decode_backup_value(encoded)),
                        )
                        kv_added += 1
                    sessions_done.add(cid)
                except sqlite3.DatabaseError:
                    con.rollback()
                    skipped += 1
                    continue
            con.commit()
        finally:
            con.close()

    return {
        "sessions": len(sessions_done),
        "table_rows": table_rows,
        "mirror_added": mirror_added,
        "kv_added": kv_added,
        "skipped": skipped,
    }


def delete_sessions(sessions: List[Session]) -> dict:
    """
    删除指定会话：表行 + 镜像条目 + 正文键，单事务提交，之后 VACUUM。
    返回统计信息。本函数不自动备份；需要留档时请先调用
    backup_sessions() 备份勾选会话。
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
    for label, key in [("已归档", S_ARCHIVED), ("镜像残留(空壳)", S_MIRROR_ONLY),
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
    print("本操作不自动备份，如需留档请先用 backup-sessions 备份。")
    if not args.yes:
        if input("确认删除？[y/N] ").strip().lower() != "y":
            print("已取消。")
            return
    require_closed(args.force)
    stats = delete_sessions(targets)
    print(f"完成: 会话 {stats['sessions']}，表行 {stats['table_rows']}，"
          f"镜像 {stats['mirror'][0]} -> {stats['mirror'][1]}，正文键 {stats['keys']}")


def op_backup_sessions(args):
    if not args.ids:
        print("[err] 请用 --ids 指定要备份的会话 ID（逗号分隔），"
              "或用 TUI 勾选会话后按 b。")
        return
    ids = {cid.strip() for cid in args.ids.split(",") if cid.strip()}
    if not ids:
        print("[err] --ids 为空。")
        return
    all_sessions = scan(include_hidden=True)
    sessions = [s for s in all_sessions if s.composer_id in ids]
    if not sessions:
        print("[err] 没有找到与 --ids 匹配的会话。")
        return
    out = backup_sessions(sessions)
    kv = sum(s.content_keys for s in sessions)
    print(f"已备份 {len(sessions)} 个会话（正文键 {kv} 个）: {out}")


def op_restore_sessions(args):
    if args.file:
        baks = [os.path.abspath(args.file)]
        if not os.path.isfile(baks[0]):
            print(f"[err] 备份文件不存在: {baks[0]}")
            return
    else:
        baks = list_session_backups()
        if not baks:
            print("没有找到会话备份（sessions-backup-*.json）。")
            return
        print("可用备份（新到旧）:")
        for i, p in enumerate(baks, 1):
            ts = fmt_message_ts(os.path.getmtime(p) * 1000)
            print(f"  [{i}] {os.path.basename(p)}  ({ts})")
        try:
            idx = int(input("选择要恢复的备份编号: ")) - 1
        except (ValueError, EOFError):
            print("已取消。")
            return
        if not (0 <= idx < len(baks)):
            print("[err] 编号无效。")
            return
        baks = [baks[idx]]
    require_closed(args.force)
    if not args.yes and input("恢复只写回备份中缺失的数据，不覆盖现有会话，继续？[y/N] ").strip().lower() != "y":
        print("已取消。")
        return
    stats = restore_sessions(baks[0])
    print(f"完成: 恢复会话 {stats['sessions']}，表行 {stats['table_rows']}，"
          f"镜像 {stats['mirror_added']}，正文键 {stats['kv_added']}，跳过已存在 {stats['skipped']}")


def op_wipe_all(args):
    sessions = scan(include_hidden=True)
    if not sessions:
        print("数据库为空。")
        return
    print(f"将清空全部会话（共 {len(sessions)} 条，含未归档），此操作不可逆！")
    print("本操作不自动备份，如需留档请先运行 backup-sessions --ids 或 TUI 按 b 备份。")
    if not args.yes and input("确认清空全部会话？[y/N] ").strip().lower() != "y":
        print("已取消。")
        return
    require_closed(args.force)
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
    print(f"镜像已修复: {before} -> {after} 个条目（本操作不自动备份）")


def require_closed(force: bool):
    if cursor_running() and not force:
        print("[err] Cursor 正在运行，请完全退出后重试（或加 --force 强行执行）")
        sys.exit(1)


# =====================================================================
# 磁盘清理（TUI 共用；纯函数，不依赖 UI）
# =====================================================================

# 可清理类目的固定顺序
CLEANUP_TARGETS = ("backups", "search_index", "vacuum", "cache_dirs")
# 需要 Cursor 完全退出才能安全清理的类目
CLEANUP_NEEDS_CLOSED = frozenset({"vacuum", "cache_dirs"})

CACHE_DIR_NAMES = (
    "Cache", "Code Cache", "GPUCache", "DawnCache", "CachedData",
    "logs", "Crashpad",
)


def _fmt_bytes(n: int) -> str:
    """人类可读的字节数。"""
    if n >= 1 << 30:
        return f"{n / (1 << 30):.1f} GB"
    if n >= 1 << 20:
        return f"{n / (1 << 20):.1f} MB"
    if n >= 1 << 10:
        return f"{n / (1 << 10):.1f} KB"
    return f"{n} B"


def _dir_size(path: str) -> int:
    """递归统计目录占用字节数，跳过无法访问的条目。"""
    total = 0
    for root, _dirs, files in os.walk(path):
        for name in files:
            try:
                total += os.path.getsize(os.path.join(root, name))
            except OSError:
                continue
    return total


def _list_backup_files() -> List[str]:
    """收集会话级 JSON 备份与遗留的旧 .bak-* 三件套。

    sessions-backup-*.json 放在 globalStorage（会话级备份）；旧的
    state.vscdb*.bak-* 可能散落在各 workspace。用递归 glob 而非
    database_paths()，这样即使工作区目录已被整体清理、备份文件仍在
    时也能被发现。
    """
    user_dir = os.path.dirname(GLOBAL_STORAGE)  # %APPDATA%\Cursor\User
    found = [
        f for f in glob.glob(os.path.join(user_dir, "**", "sessions-backup-*.json"), recursive=True)
        if os.path.isfile(f)
    ]
    found += [
        f for f in glob.glob(os.path.join(user_dir, "**", "state.vscdb*.bak-*"), recursive=True)
        if os.path.isfile(f)
    ]
    return _unique_paths(found)


def _list_search_indexes() -> List[str]:
    """globalStorage 与各 workspace 下的会话搜索索引。"""
    paths = [SEARCH_INDEX] + [
        f for f in glob.glob(os.path.join(WORKSPACE_STORAGE, "*", "conversation-search.db"))
        if os.path.isfile(f)
    ]
    return _unique_paths(paths)


def scan_cleanup_targets() -> List[dict]:
    """扫描可清理类目并统计占用字节数。只读，可安全地在后台线程执行。

    返回每项 {key, label, size_bytes, files, note, default_on, requires_closed}。
    """
    targets: List[dict] = []

    backup_files = _list_backup_files()
    targets.append({
        "key": "backups",
        "label": "工具备份文件（会话级 .json / 旧 .bak-*）",
        "size_bytes": sum(os.path.getsize(f) for f in backup_files),
        "files": len(backup_files),
        "note": "删除后无法再恢复会话，请确认不再需要",
        "default_on": True,
        "requires_closed": False,
    })

    index_files = _list_search_indexes()
    targets.append({
        "key": "search_index",
        "label": "会话搜索索引 conversation-search.db",
        "size_bytes": sum(os.path.getsize(f) for f in index_files),
        "files": len(index_files),
        "note": "Cursor 下次启动会自动重建",
        "default_on": True,
        "requires_closed": False,
    })

    vacuum_bytes = 0
    vacuum_dbs = 0
    for db_path in database_paths():
        try:
            con = open_db_ro(db_path)
        except sqlite3.Error:
            continue
        try:
            page_size = con.execute("PRAGMA page_size").fetchone()[0]
            freelist = con.execute("PRAGMA freelist_count").fetchone()[0]
            vacuum_bytes += page_size * freelist
            vacuum_dbs += 1
        except sqlite3.DatabaseError:
            continue
        finally:
            con.close()
    targets.append({
        "key": "vacuum",
        "label": "压缩会话数据库 (VACUUM)",
        "size_bytes": vacuum_bytes,
        "files": vacuum_dbs,
        "note": "回收删除会话后留下的空洞，需要 Cursor 退出",
        "default_on": True,
        "requires_closed": True,
    })

    cache_files: List[str] = []
    cache_bytes = 0
    cursor_root = os.path.dirname(GLOBAL_STORAGE)  # %APPDATA%\Cursor
    for name in CACHE_DIR_NAMES:
        p = os.path.join(cursor_root, name)
        if os.path.isdir(p):
            cache_files.append(p)
            cache_bytes += _dir_size(p)
    targets.append({
        "key": "cache_dirs",
        "label": "Cursor 缓存/日志目录",
        "size_bytes": cache_bytes,
        "files": len(cache_files),
        "note": "缓存会自动重建；logs 删除后旧日志不可追溯，需要 Cursor 退出",
        "default_on": False,
        "requires_closed": True,
    })

    return targets


def run_cleanup(selected: Dict[str, bool]) -> Dict[str, int]:
    """按勾选执行清理，返回 {key: 实际释放字节数}。

    调用方负责前置 Cursor 退出检查；单个类目失败不阻断其余类目。
    """
    chosen = {k for k, v in selected.items() if v}
    freed: Dict[str, int] = {}

    if "backups" in chosen:
        removed = 0
        for f in _list_backup_files():
            try:
                removed += os.path.getsize(f)
                os.remove(f)
            except OSError:
                continue
        freed["backups"] = removed

    if "search_index" in chosen:
        removed = 0
        for f in _list_search_indexes():
            try:
                removed += os.path.getsize(f)
                os.remove(f)
            except OSError:
                continue
        freed["search_index"] = removed

    if "vacuum" in chosen:
        freed["vacuum"] = 0
        for db_path in database_paths():
            try:
                con = open_db_ro(db_path)
                try:
                    page_size = con.execute("PRAGMA page_size").fetchone()[0]
                    reclaimable = page_size * con.execute("PRAGMA freelist_count").fetchone()[0]
                finally:
                    con.close()
            except sqlite3.Error:
                continue
            if not reclaimable:
                continue
            try:
                con = open_db_rw(db_path)
                try:
                    con.execute("VACUUM")
                    con.commit()
                finally:
                    con.close()
                freed["vacuum"] += reclaimable
            except sqlite3.DatabaseError:
                continue

    if "cache_dirs" in chosen:
        removed = 0
        cursor_root = os.path.dirname(GLOBAL_STORAGE)
        for name in CACHE_DIR_NAMES:
            p = os.path.join(cursor_root, name)
            if not os.path.isdir(p):
                continue
            size = _dir_size(p)
            shutil.rmtree(p, ignore_errors=True)
            removed += size if not os.path.exists(p) else 0
        freed["cache_dirs"] = removed

    return freed


def _copy_to_clipboard(text: str) -> bool:
    """Windows 下用 Win32 API 写入剪贴板（CF_UNICODETEXT），无外部依赖。

    OpenClipboard/EmptyClipboard/SetClipboardData/CloseClipboard 位于
    user32.dll，内存分配在 kernel32.dll；两者不能混用同一个句柄。
    """
    if not sys.platform.startswith("win"):
        return False
    try:
        import ctypes
        from ctypes import wintypes

        user32 = ctypes.windll.user32
        kernel32 = ctypes.windll.kernel32
        # 64 位进程下 HGLOBAL/HANDLE/LPVOID 是指针，必须声明 restype/argtypes，
        # 否则 ctypes 按 32 位 c_int 截断返回值或转换参数，句柄会失效。
        kernel32.GlobalAlloc.restype = wintypes.HGLOBAL
        kernel32.GlobalLock.argtypes = [wintypes.HGLOBAL]
        kernel32.GlobalLock.restype = wintypes.LPVOID
        kernel32.GlobalUnlock.argtypes = [wintypes.HGLOBAL]
        kernel32.GlobalFree.argtypes = [wintypes.HGLOBAL]
        user32.SetClipboardData.argtypes = [wintypes.UINT, wintypes.HGLOBAL]
        user32.SetClipboardData.restype = wintypes.HANDLE
        data = (text + "\0").encode("utf-16-le")
        if not user32.OpenClipboard(0):
            return False
        try:
            user32.EmptyClipboard()
            buf = kernel32.GlobalAlloc(0x0042, len(data))  # GHND = GMEM_MOVEABLE | GMEM_ZEROINIT
            if not buf:
                return False
            locked = kernel32.GlobalLock(buf)
            if locked:
                ctypes.memmove(locked, data, len(data))
                kernel32.GlobalUnlock(buf)
            if not user32.SetClipboardData(13, buf):  # 13 = CF_UNICODETEXT
                kernel32.GlobalFree(buf)
                return False
            return True
        finally:
            user32.CloseClipboard()
    except (AttributeError, OSError, ValueError):
        return False


OPS = [
    ("preview", "扫描并分类会话", op_preview),
    ("delete-archived", "删除归档会话+残留+孤儿", op_delete_archived),
    ("backup-sessions", "备份指定会话为 JSON（--ids id1,id2）", op_backup_sessions),
    ("restore-sessions", "从 JSON 备份恢复会话（--file 指定，否则列表选择）", op_restore_sessions),
    ("wipe-all", "清空全部会话（危险）", op_wipe_all),
    ("purge-index", "清理会话搜索索引", op_purge_index),
    ("repair-mirror", "修复 composerHeaders 镜像", op_repair_mirror),
]


# =====================================================================
# TUI（textual）
# =====================================================================

def _db_fingerprint() -> Tuple[Tuple[str, float, int], ...]:
    """数据库文件指纹：stat (mtime, size)，用于定时刷新前的快速短路。

    只做文件系统调用（O(库数量)），避免每 2~5 秒全量 scan() 拉取所有
    cursorDiskKV 大 blob。任一 -wal/-shm 变化也计入，避免漏检未
    checkpoint 的写入。
    """
    fingerprint: List[Tuple[str, float, int]] = []
    for db_path in database_paths():
        for p in (db_path, db_path + "-wal", db_path + "-shm"):
            try:
                st = os.stat(p)
            except OSError:
                continue
            fingerprint.append((p, st.st_mtime, st.st_size))
    return tuple(sorted(fingerprint))


def _tui_imports():
    try:
        from textual import on
        from textual.app import App, ComposeResult, RenderResult
        from textual.binding import Binding
        from textual.containers import Horizontal, Vertical, VerticalScroll
        from textual.events import Click, Leave, MouseMove
        from textual.geometry import Offset, Size
        from textual.message import Message
        from textual.screen import ModalScreen, Screen
        from textual.scroll_view import ScrollView
        from textual.widgets import Button, Checkbox, DataTable, Footer, Header, Label, Static
        from textual.worker import Worker, WorkerState
        return (App, ComposeResult, RenderResult, Binding, Horizontal, Vertical, VerticalScroll,
                Click, Leave, MouseMove, Offset, Size, Message, ModalScreen, Screen, ScrollView,
                Button, Checkbox, DataTable, Footer, Header, Label, Static,
                on, Worker, WorkerState)
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
    (App, ComposeResult, RenderResult, Binding, Horizontal, Vertical, VerticalScroll,
     Click, Leave, MouseMove, Offset, Size, Message, ModalScreen, Screen, ScrollView,
     Button, Checkbox, DataTable, Footer, Header, Label, Static,
     on, Worker, WorkerState) = mods

    STATUS_LABEL = {
        S_ARCHIVED: "归档",
        S_ACTIVE: "未归档",
        S_MIRROR_ONLY: "残留",
        S_CONTENT_ONLY: "孤儿",
    }

    class UserDotHovered(Message):
        """鼠标悬停到索引圆点上。index=-1 表示离开轨道。"""

        def __init__(self, index: int, screen_x: float = 0, screen_y: float = 0) -> None:
            super().__init__()
            self.index = index
            self.screen_x = screen_x
            self.screen_y = screen_y

    class UserDotSelected(Message):
        """点击索引圆点，请求跳到对应用户消息。"""

        def __init__(self, index: int) -> None:
            super().__init__()
            self.index = index

    class UserIndexChanged(Message):
        """正文滚动后视口内高亮圆点变化（ChatLog 发出，ChatScreen 转发给轨道）。"""

        def __init__(self, index: int) -> None:
            super().__init__()
            self.index = index

    class ConfirmModal(ModalScreen[bool]):
        """确认弹窗（删除/清理等破坏性操作前使用）。"""

        def __init__(self, message: str, confirm_label: str = "确认删除"):
            super().__init__()
            self._message = message
            self._confirm_label = confirm_label

        def compose(self) -> ComposeResult:
            with Vertical(id="confirm-box"):
                yield Label(self._message, id="confirm-text")
                with Horizontal(id="confirm-btns"):
                    yield Button(self._confirm_label, variant="error", id="btn-yes")
                    yield Button("取消", variant="primary", id="btn-no")

        def on_button_pressed(self, event: Button.Pressed) -> None:
            self.dismiss(event.button.id == "btn-yes")

    class CleanupScreen(Screen[None]):
        """磁盘清理面板：列出可清理类目与占用，勾选后执行。

        占用扫描走后台 worker（缓存目录递归统计可能较慢）；执行清理
        也在 worker 线程完成，完成后自动返回主界面。
        """

        BINDINGS = [
            Binding("escape", "close", "返回"),
            Binding("q", "close", "返回"),
        ]

        CSS = """
        #cleanup-status { height: 1; background: $panel; color: $text; padding: 0 1; }
        #cleanup-list { height: 1fr; padding: 0 1; }
        /* Checkbox 默认 tall 边框 + padding 占 3 行高，4 个类目在小终端
           里会被挤出可视区；压缩成单行，标签/大小/说明同一行显示。 */
        #cleanup-list Checkbox {
            border: none;
            padding: 0 1;
            min-height: 1;
            height: 1;
        }
        .cleanup-size { width: 12; text-align: right; color: $text-muted; }
        .cleanup-note { width: 1fr; color: $text-muted; }
        #cleanup-actions { height: 3; padding: 0 1; align: left middle; }
        #cleanup-actions Button { margin-right: 1; }
        #cleanup-foot { height: 1; background: $panel; color: $text-muted; padding: 0 1; }
        """

        def __init__(self) -> None:
            super().__init__()
            self._targets: List[dict] = []
            self._selected: Set[str] = set()
            self._cursor_running = cursor_running()
            self._scan_worker: Optional[Worker] = None
            self._run_worker: Optional[Worker] = None

        def compose(self) -> ComposeResult:
            yield Header(show_clock=True)
            yield Static("正在扫描磁盘占用…", id="cleanup-status")
            with VerticalScroll(id="cleanup-list"):
                yield Label("扫描中…", id="cleanup-empty")
            with Horizontal(id="cleanup-actions"):
                yield Button("全部勾选", id="ck-all")
                yield Button("取消勾选", id="ck-none")
                yield Button("执行清理", variant="error", id="ck-run")
            yield Static(
                "[空格]勾选/取消  [q/esc]返回 — 需退出 Cursor 的类目在运行中会禁用",
                id="cleanup-foot",
            )

        def on_mount(self) -> None:
            self._scan_worker = self.run_worker(
                scan_cleanup_targets, thread=True, exit_on_error=False
            )

        def _render_targets(self) -> None:
            """按扫描结果重建类目列表（单行：勾选框 + 大小 + 说明）。

            嵌套结构通过容器构造器传 children 一次性构建，再由
            mount_all 递归挂载；不能对未挂载的节点单独调用 mount。
            """
            container = self.query_one("#cleanup-list", VerticalScroll)
            container.remove_children()
            items: List[Any] = []
            for t in self._targets:
                key = t["key"]
                disabled = bool(t["requires_closed"] and self._cursor_running)
                if disabled:
                    self._selected.discard(key)
                # 禁用的类目在 label 上直接标注原因，避免依赖深色样式区分。
                label = t["label"] + ("（需退出 Cursor）" if disabled else "")
                cb = Checkbox(
                    label,
                    id=f"ck-{key}",
                    value=key in self._selected,
                    disabled=disabled,
                )
                size = Static(_fmt_bytes(t["size_bytes"]), classes="cleanup-size")
                note = Label(t["note"], classes="cleanup-note")
                items.append(Horizontal(cb, size, note))
            container.mount_all(items)

        def _update_status(self) -> None:
            total = sum(t["size_bytes"] for t in self._targets if t["key"] in self._selected)
            self.query_one("#cleanup-status", Static).update(
                f"可清理 {len(self._targets)} 类  |  已勾选 {len(self._selected)} 项，"
                f"预计释放 {_fmt_bytes(total)}"
            )

        def _sync_checkboxes(self) -> None:
            """把全选/取消全选的结果同步到各 Checkbox（禁用的跳过）。"""
            for t in self._targets:
                key = t["key"]
                cb = self.query_one(f"#ck-{key}", Checkbox)
                if not cb.disabled and cb.value != (key in self._selected):
                    cb.value = key in self._selected
            self._update_status()

        def on_checkbox_changed(self, event: Checkbox.Changed) -> None:
            key = (event.checkbox.id or "").removeprefix("ck-")
            if key in {t["key"] for t in self._targets}:
                if event.value:
                    self._selected.add(key)
                else:
                    self._selected.discard(key)
            self._update_status()

        def on_button_pressed(self, event: Button.Pressed) -> None:
            bid = event.button.id or ""
            if bid == "ck-all":
                self._selected = {t["key"] for t in self._targets}
                self._sync_checkboxes()
            elif bid == "ck-none":
                self._selected.clear()
                self._sync_checkboxes()
            elif bid == "ck-run":
                self._ask_run()

        def _ask_run(self) -> None:
            if self._run_worker is not None:
                self.notify("清理正在进行中…", timeout=2)
                return
            if not self._selected:
                self.notify("没有勾选任何类目", timeout=3)
                return
            total = sum(t["size_bytes"] for t in self._targets if t["key"] in self._selected)
            lines = [f"将清理 {len(self._selected)} 项，预计释放 {_fmt_bytes(total)}："]
            lines += [
                f"  - {t['label']}: {_fmt_bytes(t['size_bytes'])}"
                + (f"（{t['note']}）" if t.get("note") else "")
                for t in self._targets
                if t["key"] in self._selected
            ]
            self.app.push_screen(
                ConfirmModal("\n".join(lines), confirm_label="确认清理"), self._run
            )

        def _run(self, result: bool) -> None:
            if not result:
                self.notify("已取消", timeout=2)
                return
            if self._selected & CLEANUP_NEEDS_CLOSED and cursor_running():
                self.notify("所选类目需要 Cursor 完全退出才能清理", severity="error", timeout=5)
                return
            chosen = {key: True for key in self._selected}
            self.query_one("#cleanup-status", Static).update("正在清理…")
            self._run_worker = self.run_worker(
                lambda: run_cleanup(chosen), thread=True, exit_on_error=False
            )

        @on(Worker.StateChanged)
        def on_worker_state_changed(self, event: "Worker.StateChanged") -> None:
            # 注意：PENDING/RUNNING 状态也会派发本事件，worker 引用只在
            # SUCCESS/ERROR 终止状态时清空，否则后续事件会因 is 判断失败
            # 而丢失（参照 ChatScreen 的 worker 回调结构）。
            if event.worker is self._scan_worker:
                if event.state == WorkerState.SUCCESS:
                    self._targets = event.worker.result or []
                    self._selected = {t["key"] for t in self._targets if t["default_on"]}
                    self._render_targets()
                    self._update_status()
                    self._scan_worker = None
                elif event.state == WorkerState.ERROR:
                    self.notify(f"扫描失败: {event.worker.error}", severity="error", timeout=5)
                    self._scan_worker = None
            elif event.worker is self._run_worker:
                if event.state == WorkerState.SUCCESS:
                    freed = event.worker.result or {}
                    total = sum(freed.values())
                    detail = "，".join(
                        f"{t['label']} {_fmt_bytes(freed[t['key']])}"
                        for t in self._targets
                        if freed.get(t["key"])
                    )
                    self.notify(f"清理完成: 释放 {_fmt_bytes(total)}（{detail}）", timeout=6)
                    self.dismiss(None)
                    self._run_worker = None
                elif event.state == WorkerState.ERROR:
                    self.notify(f"清理失败: {event.worker.error}", severity="error", timeout=5)
                    self._run_worker = None

        def action_close(self) -> None:
            self.dismiss(None)

    class ChatLog(ScrollView):
        """虚拟化聊天记录滚动视图。

        聊天内容在 worker 线程预排版为 ChatLayout（固定宽度下的行列表，
        每行是带 Rich markup 的文本），本组件渲染时只取视口内的几十行，
        打开/滚动成本与消息总数无关（原 MarkdownViewer 需要把整段
        markdown 转成几百个 widget 节点并全量布局，666 条消息打开要
        冻结约 5.5 秒）。
        """

        DEFAULT_CSS = """
        ChatLog {
            height: 1fr;
            scrollbar-gutter: stable;
            background: $surface;
        }
        """

        def __init__(self, layout: "ChatLayout", msgs: List[dict], title: str, **kwargs):
            super().__init__(**kwargs)
            self._layout = layout
            self._msgs = msgs          # 宽度变化时重建排版用
            self._title = title
            self._current_user = -1    # 视口内高亮圆点下标
            self._rebuild_worker: Optional[Worker] = None
            self.virtual_size = Size(layout.width, layout.total_lines)

        # ---- 布局 ----

        def set_layout(self, layout: "ChatLayout") -> None:
            """替换排版结果（resize 后由 worker 重建调用）。"""
            self._layout = layout
            self.virtual_size = Size(layout.width, layout.total_lines)
            self.refresh()

        def on_mount(self) -> None:
            # 首次布局完成后按真实宽度校正（初始宽度是估算值）
            self.call_after_refresh(self._check_width)

        def on_resize(self) -> None:
            self._check_width()

        def _available_width(self) -> int:
            try:
                region = self.scrollable_content_region
                if region.width > 0:
                    return max(8, region.width)
            except Exception:
                pass
            return max(8, self.size.width - 2)

        def _check_width(self) -> None:
            width = self._available_width()
            if width != self._layout.width:
                self._schedule_rebuild(width)

        def _schedule_rebuild(self, width: int) -> None:
            """宽度变化时后台重建排版（0.1~0.2s，不阻塞事件循环）。"""
            if self._rebuild_worker is not None:
                return
            msgs, title = self._msgs, self._title
            self._rebuild_worker = self.run_worker(
                lambda: build_chat_layout(msgs, width, title),
                thread=True, exit_on_error=False,
            )

        def on_worker_state_changed(self, event: "Worker.StateChanged") -> None:
            if event.worker is not self._rebuild_worker:
                return
            if event.state == WorkerState.SUCCESS:
                layout = event.worker.result
                if layout is not None and layout.width != self._layout.width:
                    self.set_layout(layout)
                self._rebuild_worker = None
            elif event.state in (WorkerState.ERROR, WorkerState.CANCELLED):
                self._rebuild_worker = None

        # ---- 渲染（虚拟化：只生成视口内的行） ----

        def render(self) -> RenderResult:
            lines = self._layout.lines
            total = self._layout.total_lines
            if not lines:
                return ""
            sy = min(int(self.scroll_y), max(0, total - 1))
            height = max(1, self.size.height)
            # 每行用 Rich Text 编译（支持 [b]/[i]/[dim] 等标记），
            # 只编译视口内的行，成本与消息总数无关。
            from rich.console import Group
            from rich.text import Text

            return Group(*(Text.from_markup(line) for line in lines[sy : sy + height]))

        # ---- 圆点高亮（滚动事件驱动，替代原先 0.15s 轮询） ----

        def watch_scroll_y(self, old_value: float, new_value: float) -> None:
            super().watch_scroll_y(old_value, new_value)
            self._update_current_user()

        def _update_current_user(self) -> None:
            msg_idx = self._layout.user_index_at(int(self.scroll_y), max(1, self.size.height))
            # 转成圆点序号（UserIndexRail 的 index 语义）
            dot = self._layout.dot_of_msg(msg_idx)
            if dot != self._current_user:
                self._current_user = dot
                self.post_message(UserIndexChanged(dot))

        def jump_to_user(self, dot_index: int) -> None:
            """圆点跳转：dot_index 是圆点序号，直接滚动到该用户消息起始行。

            行号在排版时预计算，跳转精确直达（原实现依赖文档树里
            widget 的 region，且用户标题会被嵌套结构吞掉导致跳转失效）。
            """
            users = self._layout.user_msg_indices
            if 0 <= dot_index < len(users):
                self.scroll_to(y=self._layout.line_of_msg(users[dot_index]), animate=False)

    class UserIndexRail(ScrollView):
        """右侧用户输入索引轨道：圆点 + 虚线连接，可独立滚动。

        每条用户消息一个圆点，消息之间以竖虚线相连；悬停圆点弹出
        信息面板（ChatScreen 负责），点击圆点跳转到正文对应位置。

        继承 ScrollView 做线渲染（render() 返回文本），不挂子组件，
        避免容器布局/子组件更新在渲染管线中触发竞态。
        """

        DEFAULT_CSS = """
        UserIndexRail {
            width: 3;
            scrollbar-size-vertical: 0;
            background: $surface;
        }
        """

        def __init__(self, user_msgs: List[Tuple[str, str]], **kwargs):
            super().__init__(**kwargs)
            self._user_msgs = user_msgs
            self._hovered = -1   # 悬停圆点行号
            self._current = -1   # 视口内高亮圆点行号
            n = len(user_msgs)
            # 每行 1 个圆点或虚线，内容宽 3（无边框，圆点右对齐贴滚动条）
            self.virtual_size = Size(3, n * 2 - 1 if n else 1)

        def render(self) -> RenderResult:
            lines: List[str] = []
            n = len(self._user_msgs)
            for i in range(n):
                if i == self._hovered:
                    dot = "[b $accent reverse]●[/]"
                elif i == self._current:
                    dot = "[b $primary]●[/]"
                else:
                    dot = "●"
                lines.append(f"  {dot}")
                if i < n - 1:
                    lines.append("  ┆")
            return "\n".join(lines)

        def set_current(self, index: int) -> None:
            """由 ChatScreen 在正文滚动时调用，高亮视口内圆点。"""
            if index != self._current:
                self._current = index
                self.refresh()

        def _dot_index_at(self, y: float) -> int:
            # 鼠标坐标相对本组件 region（无边框，内容区从 y=0 开始）；
            # 圆点行占 1 行、虚线行占 1 行，轮流排列。
            line = int(y) + int(self.scroll_y)
            if line < 0 or line % 2 != 0:
                return -1
            i = line // 2
            return i if 0 <= i < len(self._user_msgs) else -1

        def _set_hover(self, index: int, screen_x: float = 0, screen_y: float = 0) -> None:
            if index != self._hovered:
                self._hovered = index
                self.refresh()
                self.post_message(UserDotHovered(index, screen_x, screen_y))

        def on_mouse_move(self, event: MouseMove) -> None:
            self._set_hover(self._dot_index_at(event.y), event.screen_x, event.screen_y)

        def on_leave(self, event: Leave) -> None:
            self._set_hover(-1)

        def on_click(self, event: Click) -> None:
            i = self._dot_index_at(event.y)
            if i >= 0:
                self.post_message(UserDotSelected(i))

    class ChatScreen(Screen[None]):
        """会话聊天记录查看器（全屏，可滚动）。

        正文用 ChatLog 虚拟化滚动视图（只渲染视口内几十行），右侧带
        用户输入索引轨道：圆点索引每条用户消息，悬停查看简略信息，
        点击跳转到正文对应行；正文滚动时高亮当前圆点。
        """

        BINDINGS = [
            # ChatScreen 打开后，v 不应再冒泡到 CleanerApp，避免重复
            # push_screen 导致同一条记录被层层打开、返回次数增加。
            Binding("v", "ignore_view_chat", "", show=False),
            Binding("escape", "close", "返回"),
            Binding("q", "close", "返回"),
        ]

        CSS = """
        #chat-body { height: 1fr; }
        #user-rail {
            position: absolute;
            layer: overlay;
            width: 3;
            scrollbar-size-vertical: 0;
            background: $surface;
        }
        #user-preview {
            display: none;
            position: absolute;
            layer: overlay;
            border: round $primary;
            background: $panel;
            padding: 0 1;
            height: auto;
            color: $text;
        }
        #chat-footer { height: 1; background: $panel; color: $text-muted; padding: 0 1; }
        """

        def __init__(self, title: str, msgs: List[dict], layout: "ChatLayout"):
            super().__init__()
            self._title = title
            self._msgs = msgs
            self._layout = layout
            # 圆点轨道数据来自排版结果，与正文一一对应，不再依赖
            # 解析文档树收集标题块（旧实现会漏掉被嵌套吞掉的标题）。
            self._user_msgs = [
                (ml.time, ml.summary) for ml in layout.msg_layouts if ml.is_user
            ]
            self._last_preview_index = -1

        def compose(self) -> ComposeResult:
            yield Header(show_clock=True)
            with Horizontal(id="chat-body"):
                yield ChatLog(self._layout, self._msgs, self._title, id="chat-log")
            # 索引轨道与信息面板都是绝对定位覆盖层，位置由
            # absolute_offset 动态指定（_place_rail / _update_preview）。
            if self._user_msgs:
                yield UserIndexRail(self._user_msgs, id="user-rail")
            yield Static(id="user-preview")
            yield Static(
                f"[q/esc]返回  ↑↓/PgUp/PgDn滚动  悬停圆点查看/滚轮滚动圆点链 — {self._title}",
                id="chat-footer",
            )

        def on_mount(self) -> None:
            self._chat_log = self.query_one("#chat-log", ChatLog)
            self._chat_log.focus()
            self._rail = self.query_one("#user-rail", UserIndexRail)
            self._preview = self.query_one("#user-preview", Static)
            # 轨道绝对定位在滚动条左侧；首次布局完成后再定位，
            # resize 时重算（此时 chat_log 的 scrollable_content_region
            # 才有效）。
            self.call_after_refresh(self._place_rail)

        def on_resize(self) -> None:
            self._place_rail()

        def _place_rail(self) -> None:
            """把索引轨道放在正文滚动条左侧，高度封顶 3/4 屏。"""
            if not hasattr(self, "_chat_log") or not hasattr(self, "_rail"):
                return  # resize 事件可能先于 on_mount 的组件查询到达
            content = self._chat_log.scrollable_content_region
            if content.height <= 0:
                return
            self._rail.absolute_offset = Offset(content.right - 3, content.y)
            rail_h = min(self._rail.virtual_size.height, max(3, content.height * 3 // 4))
            if self._rail.styles.height != rail_h:
                self._rail.styles.height = rail_h
            self._rail.styles.refresh(layout=True)

        def on_user_dot_hovered(self, event: UserDotHovered) -> None:
            event.stop()
            self._update_preview(event.index, event.screen_x, event.screen_y)

        def on_user_dot_selected(self, event: UserDotSelected) -> None:
            event.stop()
            self._chat_log.jump_to_user(event.index)

        def on_user_index_changed(self, event: UserIndexChanged) -> None:
            event.stop()
            self._rail.set_current(event.index)

        def _update_preview(self, index: int, screen_x: float = 0, screen_y: float = 0) -> None:
            """在鼠标附近显示简略信息面板，防止超出屏幕。

            只在圆点切换时更新并重布局，避免鼠标在圆点间滑动时
            每次移动都触发全屏重布局（旧实现每帧 styles.refresh）。
            """
            if index < 0 or index >= len(self._user_msgs):
                if self._last_preview_index >= 0:
                    self._preview.display = False
                    self._last_preview_index = -1
                return
            if index == self._last_preview_index:
                return
            self._last_preview_index = index
            ts, summary = self._user_msgs[index]
            self._preview.update(f"[b]{ts}[/b]\n{summary}")
            panel_w = min(int(self.size.width * 0.45), 70)
            est_h = 3 + (len(summary) + 59) // 60
            # 默认在鼠标右侧；放不下则移到左侧。
            left = int(screen_x) + 1
            if left + panel_w >= self.size.width:
                left = int(screen_x) - panel_w - 1
            # 默认在鼠标下方；超出下缘则移到上方。
            top = int(screen_y) + 1
            if top + est_h >= self.size.height:
                top = int(screen_y) - est_h - 1
            left = max(1, left)
            top = max(1, top)
            self._preview.styles.width = panel_w
            self._preview.absolute_offset = Offset(left, top)
            self._preview.styles.refresh(layout=True)
            self._preview.display = True

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
            Binding("c", "copy_id", "复制ID"),
            Binding("d", "delete_selected", "删除勾选"),
            Binding("b", "do_backup", "备份"),
            Binding("x", "open_cleanup", "磁盘清理"),
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
            self._last_fingerprint: Optional[Tuple[Tuple[str, float, int], ...]] = None
            self._chat_worker: Optional[Worker] = None
            self._chat_title: Optional[str] = None
            self._backup_worker: Optional[Worker] = None

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
            # 指纹无变化时跳过 scan()，避免周期性全量扫描阻塞事件循环。
            self._refresh_timer = self.set_interval(5.0, self.refresh_data_quiet)

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
            # 数据库文件没有变化时直接跳过，避免每 5 秒全量 scan()。
            try:
                fp = _db_fingerprint()
            except OSError:
                fp = None
            if fp is not None and fp == self._last_fingerprint:
                return
            self.refresh_data(quiet=True)

        def refresh_data(self, quiet: bool = False, force: bool = False) -> bool:
            current_id = self.row_to_id.get(self.current_row) if self.current_row else None
            old_signature = self._session_signature(self.sessions)
            try:
                new_sessions = scan()
                new_signature = self._session_signature(new_sessions)
                # scan 成功后同步指纹，供定时刷新短路；失败路径不更新，
                # 保留旧指纹以便下次定时重试。
                try:
                    self._last_fingerprint = _db_fingerprint()
                except OSError:
                    self._last_fingerprint = None
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
            # 状态栏展示用缓存版检测，避免每次勾选/切选项卡都 spawn
            # tasklist 子进程阻塞事件循环（Windows 下 0.5~2s）。
            running = "● 运行中" if cursor_running_cached() else "○ 未运行"
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
                "[空格]勾选 [a]全选 [n]取消 [v]聊天 [c]复制ID [d]删除 [b]备份 [x]磁盘清理 [r]刷新 [q]退出"
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
            if not self.selected:
                self.notify("没有勾选任何会话（先空格勾选，再按 b 备份）", timeout=3)
                return
            if self._backup_worker is not None:
                self.notify("备份正在进行中…", timeout=2)
                return
            selected_sessions = [s for s in self.sessions if s.composer_id in self.selected]
            self._backup_worker = self.run_worker(
                lambda: backup_sessions(selected_sessions), thread=True, exit_on_error=False
            )
            self.notify(f"正在备份 {len(selected_sessions)} 个会话…", timeout=3)

        def action_refresh(self) -> None:
            if self.refresh_data(force=True):
                self.notify("已刷新会话列表", timeout=2)

        def action_copy_id(self) -> None:
            """把当前行的完整会话 ID 复制到剪贴板（表格 ID 列只显示前 13 位）。"""
            if not self.current_row:
                self.notify("先选中一行再复制（↑↓ 移动高亮）", timeout=3)
                return
            cid = self.row_to_id.get(self.current_row)
            if cid is None:
                return
            if _copy_to_clipboard(cid):
                self.notify(f"已复制会话 ID: {cid}", timeout=4)
            else:
                self.notify("复制失败：剪贴板不可用", severity="error", timeout=3)

        def action_open_cleanup(self) -> None:
            # 聊天查看器打开时 x 不应冒泡进来；面板已在台上时也不重复压入。
            if isinstance(self.screen, (ChatScreen, CleanupScreen)):
                return
            self.push_screen(CleanupScreen(), callback=self._on_cleanup_closed)

        def _on_cleanup_closed(self, _result: None) -> None:
            # VACUUM/缓存清理会改变库文件大小，返回后刷新列表与状态栏。
            self.refresh_data()

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
            if self._chat_worker is not None:
                self.notify("聊天记录正在加载中…", timeout=2)
                return
            sess = next((s for s in self.sessions if s.composer_id == cid), None)
            title = sess.display_name if sess else cid

            def _generate() -> "Tuple[List[dict], ChatLayout]":
                # 数据层在 worker 线程内新建 sqlite3 连接，主线程不参与，
                # 线程安全；生成过程不再阻塞事件循环。宽度先用屏幕宽度
                # 估算，ChatLog 挂载后按真实内容宽度自动校正重建。
                msgs = fetch_conversation(cid)
                width_hint = max(8, self.screen.size.width - 2)
                layout = build_chat_layout(msgs, width_hint, title)
                return msgs, layout

            self._chat_worker = self.run_worker(_generate, thread=True, exit_on_error=False)
            self._chat_title = title
            self.notify("正在加载聊天记录…", timeout=2)

        @on(Worker.StateChanged)
        def on_worker_state_changed(self, event: "Worker.StateChanged") -> None:
            """聊天详情/备份 worker 结束时在主线程处理结果。"""
            if event.worker is self._backup_worker:
                if event.state == WorkerState.SUCCESS:
                    worker = self._backup_worker
                    self._backup_worker = None
                    out_path = worker.result if isinstance(worker.result, str) else ""
                    self.notify(
                        f"备份完成: {os.path.basename(out_path) if out_path else '未知文件'}",
                        timeout=5,
                    )
                elif event.state == WorkerState.ERROR:
                    worker = self._backup_worker
                    self._backup_worker = None
                    self.notify(f"备份失败: {worker.error}", severity="error", timeout=5)
                elif event.state == WorkerState.CANCELLED:
                    self._backup_worker = None
                return
            if event.worker is not self._chat_worker:
                return
            if event.state == WorkerState.SUCCESS:
                worker = self._chat_worker
                title = self._chat_title or ""
                self._chat_worker = None
                self._chat_title = None
                if self.screen is not None:
                    msgs, layout = worker.result or ([], None)
                    if layout is not None:
                        self.push_screen(ChatScreen(title, msgs, layout))
            elif event.state == WorkerState.ERROR:
                worker = self._chat_worker
                self._chat_worker = None
                self._chat_title = None
                self.notify(f"读取聊天记录失败: {worker.error}", severity="error", timeout=5)
            elif event.state == WorkerState.CANCELLED:
                self._chat_worker = None
                self._chat_title = None

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
                   f"删除前不会自动备份，建议先按 b 备份勾选会话；且要求 Cursor 已退出。")

            def _ask(result: bool) -> None:
                if not result:
                    self.notify("已取消", timeout=2)
                    return
                if cursor_running():
                    self.notify("Cursor 正在运行！请先完全退出再删除。", severity="error", timeout=5)
                    return
                try:
                    stats = delete_sessions(selected_sessions)
                except Exception as e:
                    self.notify(f"删除失败: {e}", severity="error", timeout=5)
                    return
                self.selected.clear()
                self.refresh_data()
                self.notify(
                    f"完成: 会话 {stats['sessions']}，镜像 {stats['mirror'][0]}→{stats['mirror'][1]}，正文键 {stats['keys']}",
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
    ap.add_argument("--ids", help="备份/操作指定的会话 ID，逗号分隔（配合 --op backup-sessions）")
    ap.add_argument("--file", help="指定备份文件路径（配合 --op restore-sessions）")
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
