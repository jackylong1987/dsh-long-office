# dsh-long-office

DSH Web 插件：**Office 读取 + 生成**。读取 Word/Excel/PowerPoint（含旧版格式）为文本/表格，内存解析、不落中间文件；并从 Markdown 生成 Word / Excel / PowerPoint（`md2docx` / `md2xlsx` / `md2pptx`，原 `dsh-long-plugins` 的 `md2docx` 迁入本插件）。

**工具一览**：`office_read`（读 Office）、`md2docx`（Markdown→Word）、`md2xlsx`（Markdown 表格→Excel）、`md2pptx`（Markdown 大纲→PPT，16:9/备注/图片/坐标回显）。

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

### 2. 生成 Office 文档（Markdown → Word / Excel / PowerPoint）

| 工具 | 产物 | 依赖 | 要点 |
| --- | --- | --- | --- |
| `md2docx` | `.docx` | `python3` + `python-docx` | 带页码；中文公文排版（宋体＋Times New Roman／1.5 倍行距／首行缩进 2 字符／表格居中／无页眉与装饰线／正文不加粗） |
| `md2xlsx` | `.xlsx` | `python3` + `openpyxl` | 每个 Markdown 表格一个 sheet（名取上方标题）；表头加粗、细边框、内容居中、**无底纹**、列宽自适应、冻结首行；`=` 开头写成**真公式** |
| `md2pptx` | `.pptx` | `python3` + `python-pptx` + `Pillow` | 16:9；`# `封面、`## `分页、`- `要点、`> 备注：`演讲者备注、`![说明](图 "宽x高")` 插图；生成后回显每页元素坐标 |

三者默认都**不覆盖**已存在文件（工具的 `overwrite: true` 才覆盖）。脚本路径可分别用配置 `md2docxScript` / `md2xlsxScript` / `md2pptxScript` 覆盖（默认：插件内 `lib/md2docx.py`、`lib/md2xlsx.py`、`lib/md2pptx.py`）。

---

## 安装

在 DSH profile（如 `web`）下执行：

```bash
# 通过 npm 安装（未发布时用本地 file: 链接）
dsh plugin --profile web add dsh-long-office
# 源码调试：
dsh plugin --profile web add file:./dsh-long-office
```

`cordis.patch.yml` 会把插件 `dsh-long-office` 插入 profile 的层栈（`inject: [webRuntime]`）。

---

## 提供的服务端能力

- 工具 `office_read`：读取 Office 文件 → 文本/表格。
- 工具 `md2docx`：Markdown → Word。
- 路由 `GET /api/dsh-long-office/content?path=<绝对路径>`：返回 Office 文件的内存文本/表格。

> 说明：本插件主要提供服务端工具与路由，`client/client.js` 仅为一个占位模块（无前端 UI）。

---

## 配置

`cordis.patch.yml` 默认复用 DSH Web 内建的 `trustedHosts`，可选覆盖 `md2docxScript`：

```yaml
- insert:
    - id: dsh-long-office
      name: dsh-long-office
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
curl -s "http://127.0.0.1:3080/api/dsh-long-office/content?path=/absolute/path/to/example.docx"
# 期望：{"ok":true,"content":"..."}

# 3) 工具可用（通过会话问模型）
# "用 office_read 读一下 /path/to/example.xlsx"
# "用 md2docx 把 /path/to/report.md 生成成 Word"
```

---

## License

MIT
