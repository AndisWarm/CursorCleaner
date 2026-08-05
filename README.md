# Cursor Cleaner

一个面向 Windows 的 Cursor 会话维护工具，提供 TUI（终端界面）和 CLI（命令行）两种使用方式。

它会读取 Cursor 的本地会话数据库，帮助你查看、备份和清理无用会话数据。聊天记录和数据库内容只在本机处理，不会上传到网络。

工具会同时检查 `globalStorage` 和 `workspaceStorage` 下的 `state.vscdb`。这样使用 Cursor 账号登录时，存放在工作区数据库中的会话也能被发现和查看；旧版/API key 场景使用的全局数据库仍保持兼容。

## 功能

- 扫描并分类 Cursor 会话
- 查看聊天记录
- 识别归档会话、镜像残留和正文孤儿数据
- 删除已归档会话、镜像残留和正文孤儿数据
- 删除前自动备份数据库
- 备份和恢复 Cursor 会话数据库
- 清理会话搜索索引 `conversation-search.db`
- 磁盘清理面板：一键清理工具备份文件、搜索索引、压缩会话数据库、删除 Cursor 缓存/日志目录
- 复制会话 ID 到剪贴板
- 支持通过 `--db` 指定测试数据库
- 自动隐藏同一账号会话产生的空草稿/占位副本，优先保留有正文的记录
- 自动检测 Cursor 外部归档/取消归档状态，默认每 5 秒检查数据库文件变化、有变化才刷新；也可按 `R` 手动刷新
- 查看聊天记录在后台线程加载，长会话不会卡住界面
- 聊天记录中的 UTC 时间会转换为当前机器本地时间显示
- 防止打开聊天记录后重复按 `v` 导致页面重复堆叠

## 会话状态

工具会根据 `state.vscdb` 中的多处数据判断会话状态：

| 状态 | 含义 |
| --- | --- |
| `ARCHIVED` | 会话表或侧边栏镜像标记为已归档 |
| `ACTIVE` | 会话仍存在且未归档（含镜像列表中存在、有正文但无表行的新版本会话） |
| `MIRROR_ONLY` | 仅剩侧边栏镜像数据且无任何正文键的空壳残留 |
| `CONTENT_ONLY` | 仅剩聊天正文数据，属于孤儿数据 |

> 注意：新版 Cursor 只把会话列表写入 `ItemTable` 镜像（`composer.composerData`），
> `composerHeaders` 表已停止更新。因此只要会话在镜像中且有正文，即使表中没有
> 对应行也视为正常活跃会话，不会被误判为残留。

## 环境要求

- Windows 10/11
- Python 3.10 或更高版本
- 已安装 Cursor

## 安装

在项目目录中打开 PowerShell 或 CMD：

```bash
python -m venv .venv
```

PowerShell 激活虚拟环境：

```powershell
.\.venv\Scripts\Activate.ps1
```

CMD 激活虚拟环境：

```bat
.venv\Scripts\activate.bat
```

安装依赖：

```bash
python -m pip install -r requirements.txt
```

## 使用方式

### 启动 TUI

```bash
python cursor_cleaner.py
```

也可以双击运行：

```text
start_cursor_cleaner.bat
```

启动窗口会显示实际执行的 `cursor_cleaner.py` 路径和修改时间，避免误运行其它目录中的旧副本。

### TUI 快捷键

| 按键 | 功能 |
| --- | --- |
| `Space` | 勾选/取消当前会话 |
| `A` | 全选当前筛选结果 |
| `N` | 取消全选 |
| `V` | 查看当前会话聊天记录 |
| `C` | 复制当前会话完整 ID 到剪贴板 |
| `D` | 删除勾选的会话 |
| `B` | 备份数据库 |
| `X` | 打开磁盘清理面板（备份/索引/VACUUM/缓存） |
| `R` | 立即刷新会话列表 |
| `Q` | 退出工具 |
| `Q` / `Esc` | 在聊天记录页面返回 |

### CLI

预览会话分类：

```bash
python cursor_cleaner.py --op preview
```

删除已归档会话、镜像残留和正文孤儿数据：

```bash
python cursor_cleaner.py --op delete-archived --yes
```

其他可用操作：

```text
backup          备份当前数据库
restore         从备份恢复
wipe-all        清空全部会话（危险）
purge-index     清理 conversation-search.db
```

跳过 Cursor 运行检测：

```bash
python cursor_cleaner.py --op delete-archived --yes --force
```

指定数据库路径进行测试（指定后只读取该文件，不再自动扫描 workspaceStorage）：

```bash
python cursor_cleaner.py --db "D:\\test\\state.vscdb" --op preview
```

## 数据位置

默认读取：

```text
%APPDATA%\Cursor\User\globalStorage\state.vscdb
```

另外会自动扫描：

```text
%APPDATA%\Cursor\User\workspaceStorage\*\state.vscdb
```

会话搜索索引默认位于：

```text
%APPDATA%\Cursor\User\globalStorage\conversation-search.db
```

## 项目结构

```text
cursor-cleaner/
├── cursor_cleaner.py          # 主程序
├── start_cursor_cleaner.bat   # Windows 启动脚本
├── requirements.txt           # Python 依赖
├── README.md                  # 项目说明
├── LICENSE                    # 开源许可证
└── .gitignore                 # Git 忽略规则
```

## 开源协议

本项目使用 MIT License，详见 [LICENSE](LICENSE)。
