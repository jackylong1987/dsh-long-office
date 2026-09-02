# dsh-office-reader

DSH Web 插件：**Office 读取 + 生成**。读取 Word/Excel/PowerPoint（含旧版格式）为文本/表格，内存解析、不落中间文件；并从 Markdown 生成带页码的 Word（原 `dsh-long-plugins` 的 `md2docx` 迁入本插件）。

---

## 功能

### 1. 读取 Office → 文本 / 表格（内存解析，模型按需用）
| 格式 | 依赖 | 说明 |
| --- | --- | --- |
| `.docx` | `mammoth` | Word 2007+ → 纯文本 |
| `.doc` | `word-extractor` | 旧版 Word → 纯文本 |
| `.xlsx` | `exceljs` | Excel 2007+ → 每工作表表格 |
| `.xls` | `xlsx` (SheetJS) | 旧版 Excel → 每工作表 CSV |
| `.pptx` | `jszip` | PPT 2007+ → 每页文本（zip 内 slide XML） |
| `.ppt` | — | 暂不支持，提示转换为 `.pptx` |
| `.txt / .md` | — | 直接按 UTF-8 读取 |

读取在内存中完成，**不生成任何中间文件**，模型按需调用 `office_read` 工具查看内容。

### 2. 生成 Word（`md2docx`）
把 Markdown 转为带页码的 Word，需 `python3` + `python-docx`。脚本路径可通过配置 `md2docxScript` 覆盖（默认：插件内 `lib/md2docx.py`）。

---

## 安装

在 DSH profile（如 `web`）下执行：

```bash
# 通过 npm 安装（未发布时用本地 file: 链接）
dsh plugin --profile web add dsh-office-reader
# 源码调试：
dsh plugin --profile web add file:./dsh-office-reader
```

`cordis.patch.yml` 会把插件 `dsh-office-reader` 插入 profile 的层栈（`inject: [webRuntime]`）。

---

## 提供的服务端能力

- 工具 `office_read`：读取 Office 文件 → 文本/表格。
- 工具 `md2docx`：Markdown → Word。
- 路由 `GET /api/dsh-office-reader/content?path=<绝对路径>`：返回 Office 文件的内存文本/表格。

> 说明：本插件主要提供服务端工具与路由，`client/client.js` 仅为一个占位模块（无前端 UI）。

---

## 配置

`cordis.patch.yml` 默认复用 DSH Web 内建的 `trustedHosts`，可选覆盖 `md2docxScript`：

```yaml
- insert:
    - id: dsh-office-reader
      name: dsh-office-reader
      inject: [webRuntime]
      config:
        # 复用内建 Web API 的 Host/Origin 白名单
        trustedHosts: !!js ctx.webRuntime.trustedHosts
        # md2docx 脚本覆盖（默认 <plugin>/lib/md2docx.py）
        # md2docxScript: /your/path/md2docx.py
```

---

## 验证

安装、重启 DSH 后，验证插件已加载并可用：

```bash
# 1) 服务已启动
curl -s -o /dev/null -w "%{http_code}\n" http://127.0.0.1:3080
# 期望：200

# 2) 读取一个 docx
curl -s "http://127.0.0.1:3080/api/dsh-office-reader/content?path=/absolute/path/to/example.docx"
# 期望：{"ok":true,"content":"..."}

# 3) 工具可用（通过会话问模型）
# "用 office_read 读一下 /path/to/example.xlsx"
# "用 md2docx 把 /path/to/report.md 生成成 Word"
```

---

## License

MIT
