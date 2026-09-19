import { fileURLToPath } from "node:url";
import { dirname, extname, join, resolve } from "node:path";
import { readFile } from "node:fs/promises";
import { spawn } from "node:child_process";
import mammoth from "mammoth";
import ExcelJS from "exceljs";
import XLSX from "xlsx";
import WordExtractor from "word-extractor";
import { defineTool } from "@deepseek-ai/dsh-tools";

const PACKAGE_DIR = dirname(fileURLToPath(import.meta.url));
const MD2DOCX_SCRIPT = resolve(process.env.MD2DOCX_SCRIPT || join(PACKAGE_DIR, "md2docx.py"));
  const MD2XLSX_SCRIPT = resolve(process.env.MD2XLSX_SCRIPT || join(PACKAGE_DIR, "md2xlsx.py"));
  const MD2PPTX_SCRIPT = resolve(process.env.MD2PPTX_SCRIPT || join(PACKAGE_DIR, "md2pptx.py"));

function sendJson(res, status, obj) {
  res.writeHead(status, { "content-type": "application/json; charset=utf-8", "cache-control": "no-store" });
  res.end(JSON.stringify(obj));
}

/** 内存读取 docx → 纯文本（mammoth）。 */
async function docxToText(buf) {
  const { value } = await mammoth.extractRawText({ buffer: buf });
  return value;
}

/** 内存读取 doc → 纯文本（word-extractor）。 */
async function docToText(buf) {
  const w = new WordExtractor();
  const d = await w.extract(Buffer.from(buf));
  return d.getBody();
}

/** 内存读取 xlsx → 表格文本（exceljs）。 */
async function xlsxToText(buf) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buf);
  const rows = [];
  wb.eachSheet((ws) => {
    rows.push(`## 工作表：${ws.name}`);
    ws.eachRow((row) => {
      rows.push((row.values || []).slice(1).map((c) => c == null ? "" : String(c)).join("\t"));
    });
    rows.push("");
  });
  return rows.join("\n");
}

/** 内存读取 xls → 表格文本（SheetJS/xlsx，支持 .xls 与 .xlsx）。 */
function xlsToText(buf) {
  const wb = XLSX.read(buf, { type: "buffer" });
  const out = [];
  for (const name of wb.SheetNames) {
    out.push(`## 工作表：${name}`);
    const csv = XLSX.utils.sheet_to_csv(wb.Sheets[name]);
    out.push(csv);
    out.push("");
  }
  return out.join("\n");
}

/** 内存读取 pptx → 每页文本（pptx 是 zip + slide XML）。 */
async function pptxToText(buf) {
  const jszip = (await import("jszip")).default;
  const zip = await jszip.loadAsync(buf);
  const entryNames = Object.keys(zip.files).filter((n) => /^ppt\/slides\/slide\d+\.xml$/i.test(n)).sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
  const out = [];
  for (const n of entryNames) {
    const xml = await zip.file(n).async("string");
    const texts = [...xml.matchAll(/<a:t>([^<]*)<\/a:t>/g)].map((m) => m[1]).join("");
    out.push(texts);
  }
  return out.join("\n\n");
}

async function officeToText(path) {
  const ext = extname(path).toLowerCase();
  const buf = await readFile(path);
  if (ext === ".docx") return docxToText(buf);
  if (ext === ".doc") return docToText(buf);
  if (ext === ".xls") return xlsToText(buf);
  if (ext === ".xlsx") return xlsxToText(buf);
  if (ext === ".pptx") return pptxToText(buf);
  if (ext === ".ppt") return `暂不支持读取旧版 .ppt（请转换为 .pptx）`;
  if (/\.(txt|md|markdown|log)$/i.test(ext)) return buf.toString("utf8");
  return `不支持读取该类型：${ext || "未知"}`;
}

// 声明 apply 需要的 DSH 服务（否则 ctx.tools/ctx.webServer 会报 "without inject"）
export const inject = ["webServer", "tools"];

async function apply(ctx, config = {}) {
  // 容错：任何一步失败都只记录、不拖垮 DSH 启动
  try {
    await applyInner(ctx, config);
  } catch (error) {
    if (ctx && ctx.logger) ctx.logger.error(error instanceof Error ? error : new Error(String(error)));
    else console.error("[dsh-long-office] apply failed:", error);
  }
}

async function applyInner(ctx, config = {}) {
  const onError = (error) => ctx.logger.error(error instanceof Error ? error : new Error(String(error)));
  const runScript = (script, args) => new Promise((resolvePromise) => {
    const child = spawn("python3", [script, ...args], { stdio: ["ignore", "pipe", "pipe"] });
    let stdout = "", stderr = "";
    child.stdout.on("data", (d) => { stdout += d.toString(); });
    child.stderr.on("data", (d) => { stderr += d.toString(); });
    child.on("error", (error) => resolvePromise({ ok: false, error: String(error), stdout, stderr }));
    child.on("close", (code) => resolvePromise({ ok: code === 0, code, stdout, stderr }));
  });

  // 工具：md2docx（从 dsh-long-plugins 搬来）
  ctx.tools.register(defineTool({
    name: "md2docx",
    description: "Convert a Markdown file to a styled Word (.docx) document with a page-number footer (python-docx). Requires python3 + python-docx. Override the script path via config.md2docxScript.",
    parameters: {
      input: { type: "string", required: true, description: "Absolute path to the input .md file." },
      output: { type: "string", description: "Optional absolute path for the output .docx. Defaults to input path with .docx extension." }
    },
    output: {
      schema: { type: "object", additionalProperties: false, properties: { ok: { type: "boolean", required: true }, docxPath: { type: "string" }, error: { type: "string" } } },
      render: (_args, value) => [{ type: "text", text: value && value.ok === true ? `已生成 Word 文档：${value.docxPath}` : `md2docx 失败：${value?.error ?? "未知错误"}` }]
    },
    presentCall: (args) => {
      const inPath = resolve(String(args.input ?? ""));
      const outPath = args.output ? resolve(String(args.output)) : inPath.replace(/\.md$/i, ".docx");
      return { card: "generic", title: "md2docx", kind: "edit", locations: [{ path: outPath }] };
    },
    async execute(args) {
      const inPath = resolve(String(args.input ?? ""));
      const outPath = args.output ? resolve(String(args.output)) : inPath.replace(/\.md$/i, ".docx");
      const result = await runScript(config.md2docxScript ?? MD2DOCX_SCRIPT, [inPath, outPath]);
      if (!result.ok) return { ok: false, error: (result.stderr || result.stdout || String(result.error)).trim() || `md2docx failed (exit ${result.code})` };
      return { ok: true, docxPath: outPath };
    }
  }));


  // 工具：md2xlsx（Markdown 表格 → Excel）
  ctx.tools.register(defineTool({
    name: "md2xlsx",
    description: "Convert Markdown tables to an Excel (.xlsx) workbook: one sheet per table (named from the nearest heading), bold header, thin borders, centered, no shading, auto column width, frozen header row; a cell starting with '=' becomes a real formula. Requires python3 + openpyxl. Override the script path via config.md2xlsxScript.",
    parameters: {
      input: { type: "string", required: true, description: "Absolute path to the input .md file." },
      output: { type: "string", description: "Optional absolute path for the output .xlsx. Defaults to input path with .xlsx extension." },
      overwrite: { type: "boolean", description: "Overwrite an existing output file (default false)." }
    },
    output: {
      schema: { type: "object", additionalProperties: false, properties: { ok: { type: "boolean", required: true }, xlsxPath: { type: "string" }, error: { type: "string" } } },
      render: (_args, value) => [{ type: "text", text: value && value.ok === true ? `已生成 Excel：${value.xlsxPath}` : `md2xlsx 失败：${value?.error ?? "未知错误"}` }]
    },
    presentCall: (args) => {
      const inPath = resolve(String(args.input ?? ""));
      const outPath = args.output ? resolve(String(args.output)) : inPath.replace(/\.md$/i, ".xlsx");
      return { card: "generic", title: "md2xlsx", kind: "edit", locations: [{ path: outPath }] };
    },
    async execute(args) {
      const inPath = resolve(String(args.input ?? ""));
      const outPath = args.output ? resolve(String(args.output)) : inPath.replace(/\.md$/i, ".xlsx");
      const scriptArgs = [inPath, outPath];
      if (args.overwrite === true) scriptArgs.push("--force");
      const result = await runScript(config.md2xlsxScript ?? MD2XLSX_SCRIPT, scriptArgs);
      if (!result.ok) return { ok: false, error: (result.stderr || result.stdout || String(result.error)).trim() || `md2xlsx failed (exit ${result.code})` };
      return { ok: true, xlsxPath: outPath };
    }
  }));

  // 工具：md2pptx（Markdown 大纲 → PowerPoint，16:9，带备注/图片/坐标回显）
  ctx.tools.register(defineTool({
    name: "md2pptx",
    description: "Convert a Markdown outline to a 16:9 PowerPoint (.pptx): '# '=cover, '## '=new slide, '- '=bullets, '> 备注：'=speaker notes, '![alt](path \"WxH\")'=image. Requires python3 + python-pptx + Pillow. Override the script path via config.md2pptxScript.",
    parameters: {
      input: { type: "string", required: true, description: "Absolute path to the input .md file." },
      output: { type: "string", description: "Optional absolute path for the output .pptx. Defaults to input path with .pptx extension." },
      overwrite: { type: "boolean", description: "Overwrite an existing output file (default false)." }
    },
    output: {
      schema: { type: "object", additionalProperties: false, properties: { ok: { type: "boolean", required: true }, pptxPath: { type: "string" }, error: { type: "string" } } },
      render: (_args, value) => [{ type: "text", text: value && value.ok === true ? `已生成 PPT：${value.pptxPath}` : `md2pptx 失败：${value?.error ?? "未知错误"}` }]
    },
    presentCall: (args) => {
      const inPath = resolve(String(args.input ?? ""));
      const outPath = args.output ? resolve(String(args.output)) : inPath.replace(/\.md$/i, ".pptx");
      return { card: "generic", title: "md2pptx", kind: "edit", locations: [{ path: outPath }] };
    },
    async execute(args) {
      const inPath = resolve(String(args.input ?? ""));
      const outPath = args.output ? resolve(String(args.output)) : inPath.replace(/\.md$/i, ".pptx");
      const scriptArgs = [inPath, outPath];
      if (args.overwrite === true) scriptArgs.push("--force");
      const result = await runScript(config.md2pptxScript ?? MD2PPTX_SCRIPT, scriptArgs);
      if (!result.ok) return { ok: false, error: (result.stderr || result.stdout || String(result.error)).trim() || `md2pptx failed (exit ${result.code})` };
      return { ok: true, pptxPath: outPath };
    }
  }));

  // 工具：读 Office → 文本/表格（内存、无中间文件、模型按需用）
  ctx.tools.register(defineTool({
    name: "office_read",
    description: "Read an Office/Word/Excel/PowerPoint file (.docx/.doc/.xlsx/.xls/.pptx/.ppt) and return its text/table content in memory (no intermediate files). Use to let the model see/analyze office document content on demand.",
    parameters: {
      path: { type: "string", required: true, description: "Absolute path to the office file." }
    },
    output: {
      schema: { type: "object", additionalProperties: false, properties: { ok: { type: "boolean", required: true }, content: { type: "string" }, error: { type: "string" } } },
      render: (_args, value) => [{ type: "text", text: value && value.ok === true ? (value.content || "(空)") : `office_read 失败：${value?.error ?? "未知错误"}` }]
    },
    async execute(args) {
      try {
        const path = resolve(String(args.path ?? ""));
        const content = await officeToText(path);
        return { ok: true, content: String(content || "").slice(0, 1_000_000) };
      } catch (error) {
        return { ok: false, error: String(error && error.message || error) };
      }
    }
  }));

  // 路由：返回 office 文件的内存文本/表格（供前端/工具读取）
  ctx.effect(() => ctx.webServer.register({
    kind: "exact",
    path: "/api/dsh-long-office/content",
    handler: async (req, res) => {
      try {
        if (req.method !== "GET" && req.method !== "HEAD") { res.writeHead(405); res.end(); return; }
        const rel = (() => { try { return decodeURIComponent(new URL(req.url || "/", "http://dsh.internal").searchParams.get("path") || ""); } catch { return ""; } })();
        const content = await officeToText(rel);
        sendJson(res, 200, { ok: true, content });
      } catch (error) { onError(error); sendJson(res, 500, { ok: false, error: String(error && error.message || error) }); }
    },
  }), "dsh-long-office: content route");
}

export { apply };
export const name = "dsh-long-office";
