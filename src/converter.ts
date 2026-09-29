/**
 * export2word — Markdown → Word 核心转换引擎
 * v1.0.0
 *
 * 使用 markdown-it 解析 Markdown，docx 库生成 .docx 文件
 */

import MarkdownIt from "markdown-it";
import {
    Document, Packer, Paragraph, TextRun, HeadingLevel,
    Table, TableRow, TableCell, WidthType,
    AlignmentType, BorderStyle, ExternalHyperlink,
    ImageRun, convertInchesToTwip,
    ShadingType, TableLayoutType,
    FootnoteReferenceRun
} from "docx";
import hljs, { type HighlightResult } from "highlight.js";
import { wikilinkPlugin } from "./wikilink-plugin";
// @ts-ignore - markdown-it-footnote 无类型声明
import footnotePlugin from "markdown-it-footnote";
import type { App, Vault, TFile } from "obsidian";
import { normalizePath, Platform, requestUrl } from "obsidian";
import type { WikilinkMode } from "./settings";

// ============================================================
// 类型定义
// ============================================================

interface Token {
    type: string;
    tag: string;
    content: string;
    markup?: string;
    info?: string;
    attrs?: [string, string][];
    children?: Token[];
    level?: number;
    map?: [number, number];
    block?: boolean;
    hidden?: boolean;
    nesting?: number;
    ordered?: boolean;
    start?: number;
    /** markdown-it Token 运行时方法（读取/设置属性） */
    attrGet?(name: string): string | null;
    attrSet?(name: string, value: string): void;
}

interface ConverterContext {
    vault: Vault;
    app: App;
    sourceFile: TFile;
    wikilinkMode: WikilinkMode;
    /** 递归深度（zip 模式） */
    recursionDepth: number;
    /** 收集到的所有 wikilink 目标笔记名 */
    wikilinkTargets: string[];
    /** 当前处理到第几个 token */
    idx: number;
    /** 所有 tokens */
    tokens: Token[];
    /** 文档子元素列表 */
    children: any[];
    /**
     * 图片字节缓存：key = 图片引用原值（image token 的 src / 嵌入 wikilink 的 target）
     * 由 prefetchImages 在解析前异步填充，renderInline 同步读取
     */
    imageCache: Map<string, { data: Uint8Array; format: string; width: number; height: number }>;
    /** 脚注收集：key = 脚注标签，value = 脚注内容 */
    footnotes: Map<string, string>;
    /** 脚注计数器（用于生成序号） */
    footnoteCounter: number;
}

// ============================================================
// 工厂函数
// ============================================================

function createMarkdownParser(): MarkdownIt {
    const md = new MarkdownIt({
        html: false,
        breaks: true,
        linkify: true,
        typographer: true,
        xhtmlOut: false,
    });
    md.use(wikilinkPlugin);
    md.use(footnotePlugin);

    // 关键：覆盖 markdown-it 默认 normalizeLink。
    // 默认实现会用 new URL() 对 src 里的中文做百分号编码（如 "封面" → "%E5%B0%81%E9%9D%A2"），
    // 导致 ![](images/中文名.png) 的 src 变成编码串，vault.getFileByPath() 找不到文件 → 图片降级为占位。
    // 这里原样返回，保留 Obsidian 原始相对路径（vault 内文件路径是明文中文）。
    md.normalizeLink = (url: string) => url;
    return md;
}

// ============================================================
// 图片解析
// ============================================================

function resolveImagePath(
    src: string,
    sourceFile: TFile
): { resolvedPath: string; isLocal: boolean } {
    // 兜底：src 若带百分号编码（如 %E5%B0%81），先解码回明文中文，
    // 与 vault 内文件路径（明文）对齐；解码失败（如孤立 %）则保持原样
    if (src.includes("%")) {
        try {
            const decoded = decodeURIComponent(src);
            if (decoded !== src) src = decoded;
        } catch {
            // 忽略解码错误，沿用原值
        }
    }
    // 网络图片
    if (src.startsWith("http://") || src.startsWith("https://")) {
        return { resolvedPath: src, isLocal: false };
    }
    // 绝对路径
    if (src.startsWith("/") || src.match(/^[A-Za-z]:/)) {
        return { resolvedPath: src, isLocal: true };
    }
    // 相对路径：基于源文件所在目录（支持 ../ 相对路径，用 normalizePath 归一化）
    const sourceDir = sourceFile.parent?.path ?? "/";
    const resolved = sourceDir.endsWith("/")
        ? `${sourceDir}${src}`
        : `${sourceDir}/${src}`;
    return { resolvedPath: normalizePath(resolved), isLocal: true };
}

async function readImageAsBase64(
    vault: Vault,
    path: string
): Promise<{ data: Uint8Array; format: string } | null> {
    try {
        const file = vault.getFileByPath(path);
        if (!file) {
            // 尝试去除 Obsidian 最短路径前缀
            const altFile = vault.getFileByPath(path.replace(/^\.\//, ""));
            if (!altFile) return null;
            const buffer = await vault.readBinary(altFile);
            const ext = altFile.extension.toLowerCase();
            return {
                data: new Uint8Array(buffer),
                format: mapExtension(ext),
            };
        }
        const buffer = await vault.readBinary(file);
        const ext = file.extension.toLowerCase();
        return {
            data: new Uint8Array(buffer),
            format: mapExtension(ext),
        };
    } catch {
        return null;
    }
}

function mapExtension(ext: string): string {
    const map: Record<string, string> = {
        png: "png",
        jpg: "jpeg",
        jpeg: "jpeg",
        gif: "gif",
        svg: "svg",
        webp: "webp",
        bmp: "bmp",
    };
    return map[ext] || "png";
}

/** 正文可嵌入图片的最大显示宽度（像素，按页面可用宽度留余量） */
const IMAGE_MAX_WIDTH = 520;

/** 判断一个引用目标是否是图片文件（按扩展名） */
function isImageFile(name: string): boolean {
    return /\.(png|jpe?g|gif|bmp|webp|svg)(\?.*)?$/i.test(name.trim());
}

/**
 * 从图片字节流解析原始像素尺寸（支持 png / gif / jpeg）
 * 解析失败时回退到默认尺寸，避免 docx 报错
 */
function getImageDimensions(
    data: Uint8Array,
    format: string
): { width: number; height: number } {
    try {
        if (format === "png" && data.length >= 24) {
            // PNG：IHDR 块，字节 16-19 宽，20-23 高（大端）
            const width = (data[16] << 24) | (data[17] << 16) | (data[18] << 8) | data[19];
            const height = (data[20] << 24) | (data[21] << 16) | (data[22] << 8) | data[23];
            if (width > 0 && height > 0) return { width, height };
        } else if (format === "gif" && data.length >= 10) {
            // GIF：字节 6-7 宽，8-9 高（小端）
            const width = data[6] | (data[7] << 8);
            const height = data[8] | (data[9] << 8);
            if (width > 0 && height > 0) return { width, height };
        } else if (format === "jpeg") {
            // JPEG：扫描 SOF 标记，读取宽高
            let i = 2;
            while (i < data.length - 8) {
                if (data[i] !== 0xff) { i++; continue; }
                const marker = data[i + 1];
                // SOF0..SOF15，排除 DHT(C4)/JPG(C8)/DAC(CC)
                if (
                    marker >= 0xc0 && marker <= 0xcf &&
                    marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc
                ) {
                    const height = (data[i + 5] << 8) | data[i + 6];
                    const width = (data[i + 7] << 8) | data[i + 8];
                    if (width > 0 && height > 0) return { width, height };
                    break;
                }
                const len = (data[i + 2] << 8) | data[i + 3];
                if (len <= 0) break;
                i += 2 + len;
            }
        }
    } catch {
        /* 忽略，走默认 */
    }
    return { width: 400, height: 300 };
}

/** 把原始尺寸按最大宽度等比缩放，得到 docx 显示尺寸 */
function scaleToMaxWidth(
    dim: { width: number; height: number }
): { width: number; height: number } {
    if (dim.width <= IMAGE_MAX_WIDTH) return dim;
    const ratio = IMAGE_MAX_WIDTH / dim.width;
    return {
        width: IMAGE_MAX_WIDTH,
        height: Math.max(1, Math.round(dim.height * ratio)),
    };
}

/**
 * 读取网络图片字节（http/https），返回字节与格式
 * 失败返回 null（调用方回退为占位文字）
 */
async function loadNetworkImage(
    url: string
): Promise<{ data: Uint8Array; format: string } | null> {
    try {
        const resp = await requestUrl({ url });
        const data = new Uint8Array(resp.arrayBuffer);
        // 优先按 URL 扩展名，其次按 Content-Type
        const extMatch = url.split("?")[0].match(/\.([a-z0-9]+)$/i);
        let format = extMatch ? mapExtension(extMatch[1].toLowerCase()) : "";
        if (!format) {
            const ct = (resp.headers?.["content-type"] || resp.headers?.["Content-Type"] || "").toLowerCase();
            if (ct.includes("png")) format = "png";
            else if (ct.includes("jpeg") || ct.includes("jpg")) format = "jpeg";
            else if (ct.includes("gif")) format = "gif";
            else if (ct.includes("bmp")) format = "bmp";
            else format = "png";
        }
        return { data, format };
    } catch {
        return null;
    }
}

/**
 * 按 Obsidian 附件语义查找嵌入图片文件（![[图.png]]）
 * Obsidian 使用最短路径，target 可能是 "图.png"、"子目录/图.png"，
 * 也可能是相对当前笔记的 "../附件/图.png"（相对路径需基于源文件目录解析）
 * 优先精确路径匹配，其次按文件名匹配，最后按结尾匹配
 */
async function loadEmbedImage(
    vault: Vault,
    target: string,
    sourceFile?: TFile
): Promise<{ data: Uint8Array; format: string } | null> {
    try {
        const clean = target.trim();

        // 0) Obsidian 相对路径（![[../附件/图.png]]）：以源笔记所在目录为基准解析
        if (sourceFile && (clean.startsWith("./") || clean.startsWith("../"))) {
            const sourceDir = sourceFile.parent?.path ?? "";
            const absPath = normalizePath(`${sourceDir}/${clean}`);
            const relFile = vault.getFileByPath(absPath);
            if (relFile) {
                const buffer = await vault.readBinary(relFile);
                return {
                    data: new Uint8Array(buffer),
                    format: mapExtension(relFile.extension.toLowerCase()),
                };
            }
            return null;
        }

        const files = vault.getFiles();
        // 1) 精确路径匹配
        let file = files.find((f) => f.path === clean);
        // 2) 文件名（含扩展名）匹配
        if (!file) file = files.find((f) => f.name === clean);
        // 3) 去掉可能的路径前缀后按结尾匹配
        if (!file) file = files.find((f) => f.path.endsWith("/" + clean));
        if (!file) return null;

        const buffer = await vault.readBinary(file);
        return {
            data: new Uint8Array(buffer),
            format: mapExtension(file.extension.toLowerCase()),
        };
    } catch {
        return null;
    }
}

/**
 * 解析前预扫所有 token，异步读取图片字节写入 ctx.imageCache。
 * 覆盖：标准 image token（src）、嵌入 wikilink（data-embed 且 target 为图片）。
 * renderInline 为同步函数，故图片必须在此提前读好。
 */
async function prefetchImages(ctx: ConverterContext, tokens: Token[]): Promise<void> {
    for (const token of tokens) {
        // 递归进入 inline 的 children
        if (token.children && token.children.length > 0) {
            await prefetchImages(ctx, token.children);
        }

        const anyTok = token as any;

        // 标准 Markdown 图片 ![](src)
        if (token.type === "image") {
            const src: string = anyTok.attrGet?.("src") || "";
            if (!src || ctx.imageCache.has(src)) continue;
            let loaded: { data: Uint8Array; format: string } | null = null;
            if (src.startsWith("http://") || src.startsWith("https://")) {
                loaded = await loadNetworkImage(src);
            } else {
                const { resolvedPath } = resolveImagePath(src, ctx.sourceFile);
                loaded = await readImageAsBase64(ctx.vault, resolvedPath);
            }
            if (loaded) {
                ctx.imageCache.set(src, await normalizeImageCache(loaded));
            }
            continue;
        }

        // 嵌入图 ![[图.png]]
        if (token.type === "wikilink" && anyTok.attrGet?.("data-embed") === "true") {
            const target: string = anyTok.attrGet?.("data-target") || token.content || "";
            if (!target || !isImageFile(target) || ctx.imageCache.has(target)) continue;
            let loaded: { data: Uint8Array; format: string } | null = null;
            if (target.startsWith("http://") || target.startsWith("https://")) {
                loaded = await loadNetworkImage(target);
            } else {
                loaded = await loadEmbedImage(ctx.vault, target, ctx.sourceFile);
            }
            if (loaded) {
                ctx.imageCache.set(target, await normalizeImageCache(loaded));
            }
        }
    }
}

/**
 * 浏览器/Electron 环境下，将 svg / webp 位图转换为 PNG（base64 → Image → canvas → PNG）。
 * 桌面端（Electron）支持完整 DOM Canvas，可无损转 PNG 再交给 docx；
 * 移动端 Obsidian 无 DOM Canvas，返回 null 由调用方降级为占位文字（P0-05 收口）。
 */
async function convertImageToPng(
    data: Uint8Array,
    format: string,
    fallbackW: number,
    fallbackH: number
): Promise<{ data: Uint8Array; width: number; height: number } | null> {
    if (!Platform.isDesktop) return null;
    try {
        const mime = format === "svg" ? "image/svg+xml" : `image/${format}`;
        const dataUrl = `data:${mime};base64,${uint8ToBase64(data)}`;
        const img = new Image();
        await new Promise<void>((resolve, reject) => {
            img.onload = () => resolve();
            img.onerror = () => reject(new Error("image load failed"));
            img.src = dataUrl;
        });
        const srcW = img.naturalWidth || fallbackW || 400;
        const srcH = img.naturalHeight || fallbackH || 300;
        const ratio = srcW > IMAGE_MAX_WIDTH ? IMAGE_MAX_WIDTH / srcW : 1;
        const drawW = Math.max(1, Math.round(srcW * ratio));
        const drawH = Math.max(1, Math.round(srcH * ratio));
        const canvas = document.createElement("canvas");
        canvas.width = drawW;
        canvas.height = drawH;
        const c2d = canvas.getContext("2d");
        if (!c2d) return null;
        c2d.drawImage(img, 0, 0, drawW, drawH);
        const pngDataUrl = canvas.toDataURL("image/png");
        const pngBase64 = pngDataUrl.split(",")[1] || "";
        if (!pngBase64) return null;
        return { data: base64ToUint8Array(pngBase64), width: drawW, height: drawH };
    } catch {
        return null;
    }
}

/** Uint8Array → base64 字符串（分块避免大数组栈溢出） */
function uint8ToBase64(bytes: Uint8Array): string {
    let binary = "";
    const chunk = 0x8000;
    for (let i = 0; i < bytes.length; i += chunk) {
        binary += String.fromCharCode.apply(null, Array.from(bytes.subarray(i, i + chunk)) as any);
    }
    return btoa(binary);
}

/** base64 字符串 → Uint8Array */
function base64ToUint8Array(b64: string): Uint8Array {
    const binary = atob(b64);
    const out = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) out[i] = binary.charCodeAt(i);
    return out;
}

/**
 * 把预取的图片规整为缓存条目：svg / webp 在桌面端尝试转 PNG，
 * 转换失败则保留原格式（最终由 makeImageRun 降级为占位）。
 */
async function normalizeImageCache(
    loaded: { data: Uint8Array; format: string }
): Promise<{ data: Uint8Array; format: string; width: number; height: number }> {
    const dim = scaleToMaxWidth(getImageDimensions(loaded.data, loaded.format));
    if (loaded.format === "webp" || loaded.format === "svg") {
        const converted = await convertImageToPng(loaded.data, loaded.format, dim.width, dim.height);
        if (converted) {
            return { data: converted.data, format: "png", width: converted.width, height: converted.height };
        }
    }
    return { data: loaded.data, format: loaded.format, width: dim.width, height: dim.height };
}

/**
 * 从缓存构造 docx ImageRun；缓存未命中或格式不支持时返回 null。
 * P0-05：svg / webp 在桌面端已由 convertImageToPng 转成 PNG 预存，
 * 因此此处只需认 png/jpeg/gif/bmp；移动端未转换则降级占位。
 */
function makeImageRun(
    cache: { data: Uint8Array; format: string; width: number; height: number } | undefined
): ImageRun | null {
    if (!cache) return null;
    const supported: Record<string, true> = { png: true, jpeg: true, gif: true, bmp: true };
    if (!supported[cache.format]) return null;
    return new ImageRun({
        type: cache.format as any,
        data: cache.data,
        transformation: { width: cache.width, height: cache.height },
    });
}

// ============================================================
// 代码高亮 → docx TextRun 数组
// ============================================================

// HLJS 调色板（按 scope 名 → 颜色）
// 2026-07-16 配色加深：改用 vscode light+ 主题色（设计给白底，对比度 AA 级），
// 替换原 github-dark 浅底配色（在 F5F5F5 灰底上太浅看不清）。
const HLJS_COLORS: Record<string, string> = {
    keyword:   "0000FF", // 关键字（深蓝）
    built_in:  "267F99", // 内建（深青）
    type:      "267F99", // 类型名（深青）
    string:    "A31515", // 字符串（深红）
    number:    "098658", // 数字（深绿）
    comment:   "008000", // 注释（中深绿）
    function:  "795E26", // 函数名（棕黄）
    title:     "795E26", // 函数/类名（棕黄）
    class:     "267F99", // 类名（深青）
    attr:      "001080", // 属性名（深蓝）
    attribute: "001080",
    variable:  "001080",
    tag:       "0000FF", // 标签（深蓝）
    name:      "001080",
    meta:      "808080", // 元信息（灰，保留）
    selector:  "795E26", // 选择器（棕黄）
    params:    "001080",
    literal:   "0000FF",
    regexp:    "D16969", // 正则（深红）
};

// 把 hljs 输出的 HTML 切成 [{text, color}] 段
// 跳过所有标签字符（<, >），只保留可见文本；用 stack 跟踪当前 scope class
function parseHighlightedHtml(html: string): { text: string; color?: string }[] {
    const out: { text: string; color?: string }[] = [];
    const stack: string[] = [];
    let i = 0;
    const len = html.length;

    while (i < len) {
        const ch = html[i];

        if (ch === "<") {
            // 找结束 '>'
            const end = html.indexOf(">", i);
            if (end === -1) {
                // 残缺，剩余当文本
                pushText(out, html.slice(i), currentColor(stack));
                break;
            }
            const tag = html.slice(i + 1, end);
            // 是不是 </span> 闭合？
            if (tag.startsWith("/")) {
                // 闭合标签永远是最内层 span 先关——直接 pop 栈顶
                // （不能用 tag 名匹配，因为栈里存的是 scope 名如 "keyword"，不是 "span"）
                stack.pop();
            } else if (tag.endsWith("/")) {
                // 自闭合 <span/>，忽略
            } else {
                // 开标签 <span class="hljs-keyword">
                const m = tag.match(/class\s*=\s*"([^"]*)"/);
                if (m) {
                    const classes = m[1].split(/\s+/);
                    // 取第一个 hljs- 开头的 scope
                    const scope = classes.find(c => c.startsWith("hljs-"));
                    if (scope) {
                        const key = scope.slice(5); // 去掉 "hljs-"
                        stack.push(key);
                    }
                }
            }
            i = end + 1;
            continue;
        }

        // 处理 HTML 实体（&lt; &amp; &quot; 等）→ 还原为字面字符
        if (ch === "&") {
            const semi = html.indexOf(";", i);
            if (semi !== -1 && semi - i <= 6) {
                const entity = html.slice(i, semi + 1);
                const decoded = decodeHtmlEntity(entity);
                if (decoded !== null) {
                    pushText(out, decoded, currentColor(stack));
                    i = semi + 1;
                    continue;
                }
            }
        }

        // 普通字符：累积直到下一个 '<'，但中间可能含 HTML 实体（如 &quot;）→ 按 '&' 切分逐个解码
        const next = html.indexOf("<", i);
        const slice = next === -1 ? html.slice(i) : html.slice(i, next);
        if (slice) {
            // 用 entity 分隔切，逐段处理
            let p = 0;
            while (p < slice.length) {
                const ampIdx = slice.indexOf("&", p);
                if (ampIdx === -1) {
                    // 无 entity，整段推入
                    pushText(out, slice.slice(p), currentColor(stack));
                    break;
                }
                // 实体前文本
                if (ampIdx > p) {
                    pushText(out, slice.slice(p, ampIdx), currentColor(stack));
                }
                // 找 ';'
                const semi = slice.indexOf(";", ampIdx);
                if (semi !== -1 && semi - ampIdx <= 6) {
                    const entity = slice.slice(ampIdx, semi + 1);
                    const decoded = decodeHtmlEntity(entity);
                    if (decoded !== null) {
                        pushText(out, decoded, currentColor(stack));
                        p = semi + 1;
                        continue;
                    }
                }
                // 不是合法 entity，当作普通 '&' 字符
                pushText(out, "&", currentColor(stack));
                p = ampIdx + 1;
            }
        }
        i = next === -1 ? len : next;
    }

    return out;
}

function currentColor(stack: string[]): string | undefined {
    // 从栈顶往下找第一个能在 HLJS_COLORS 命中颜色的 scope
    for (let k = stack.length - 1; k >= 0; k--) {
        const c = HLJS_COLORS[stack[k]];
        if (c) return c;
    }
    return undefined;
}

function pushText(out: { text: string; color?: string }[], text: string, color?: string) {
    if (!text) return;
    const last = out[out.length - 1];
    if (last && last.color === color) {
        last.text += text;
    } else {
        out.push({ text, color });
    }
}

function decodeHtmlEntity(s: string): string | null {
    const map: Record<string, string> = {
        "&lt;": "<",
        "&gt;": ">",
        "&amp;": "&",
        "&quot;": '"',
        "&apos;": "'",
        "&#39;": "'",
        "&nbsp;": " ",
    };
    if (map[s]) return map[s];
    const numMatch = s.match(/^&#(\d+);$/);
    if (numMatch) return String.fromCharCode(parseInt(numMatch[1], 10));
    const hexMatch = s.match(/^&#x([0-9a-fA-F]+);$/);
    if (hexMatch) return String.fromCharCode(parseInt(hexMatch[1], 16));
    return null;
}

// 把代码一行构造成 [{text, color?}, ...] runs
function buildLineRuns(lineSegments: { text: string; color?: string }[]): TextRun[] {
    return lineSegments.map(seg =>
        new TextRun({
            text: seg.text,
            font: { name: "Courier New" },
            size: 20,
            color: seg.color,
            bold: seg.color !== undefined,
        })
    );
}

function highlightCodeToRuns(
    code: string,
    language: string
): { runs: TextRun[]; isFirst: boolean; isLast: boolean }[] {
    let result: HighlightResult;

    if (language && hljs.getLanguage(language)) {
        result = hljs.highlight(code, { language });
    } else {
        result = hljs.highlightAuto(code);
    }

    const html = result.value;
    // 按 '\n' 拆 html：hljs 输出里 '<' 一般与 '\n' 同处，用 split 不够准
    // 用更稳健的方法：扫描 html 输出，每遇到换行符 '\n' 落一行
    const segments: { text: string; color?: string }[][] = [[]];
    const allSegs = parseHighlightedHtml(html);

    for (const seg of allSegs) {
        // 把 seg.text 按 '\n' 拆开
        const parts = seg.text.split("\n");
        for (let k = 0; k < parts.length; k++) {
            if (parts[k]) {
                segments[segments.length - 1].push({ text: parts[k], color: seg.color });
            }
            if (k < parts.length - 1) {
                segments.push([]); // 下一行
            }
        }
    }

    const lines = segments.map((lineSegments) => buildLineRuns(lineSegments));
    return lines.map((runs, idx) => ({
        runs,
        isFirst: idx === 0,
        isLast: idx === lines.length - 1,
    }));
}

// ============================================================
// 主解析器：将 markdown-it token 流转为 docx 元素
// ============================================================

async function dispatchToken(ctx: ConverterContext, targetArray?: any[]): Promise<void> {
    const { tokens } = ctx;
    const token = tokens[ctx.idx];
    const destination = targetArray || ctx.children;

    switch (token.type) {
        case "heading_open":
            destination.push(await parseHeading(ctx, token));
            break;

        case "paragraph_open":
            destination.push(...await parseParagraph(ctx));
            break;

        case "bullet_list_open":
            destination.push(...(await parseList(ctx, false)));
            break;

        case "ordered_list_open":
            destination.push(...(await parseList(ctx, true)));
            break;

        case "blockquote_open":
            destination.push(await parseBlockquote(ctx));
            break;

        case "fence":
            destination.push(...(await parseFence(ctx, token)));
            break;

        case "table_open":
            destination.push(await parseTable(ctx));
            break;

        case "hr":
            destination.push(
                new Paragraph({
                    alignment: AlignmentType.LEFT,
                    children: [new TextRun({ text: "" })],
                    border: {
                        bottom: {
                            style: BorderStyle.SINGLE,
                            size: 6,
                            color: "CCCCCC",
                            space: 1,
                        },
                    },
                })
            );
            break;

        case "paragraph_close":
        case "heading_close":
        case "bullet_list_close":
        case "ordered_list_close":
        case "blockquote_close":
        case "table_close":
        case "list_item_close":
        case "footnote_close":
            ctx.idx++;
            break;

        case "footnote_open":
            ctx.idx++;
            break;

        case "footnote_anchor":
        case "footnote_block_open":
        case "footnote_block_close":
            ctx.idx++;
            break;

        default:
            ctx.idx++;
            break;
    }
}

async function parseTokens(ctx: ConverterContext): Promise<any[]> {
    const { tokens } = ctx;

    for (ctx.idx = 0; ctx.idx < tokens.length; ctx.idx++) {
        await dispatchToken(ctx, ctx.children);
    }

    return ctx.children;
}

// ============================================================
// 各元素解析器
// ============================================================

async function parseHeading(
    ctx: ConverterContext,
    token: Token
): Promise<Paragraph> {
    const level = parseInt(token.tag.replace("h", ""), 10);
    const heading = (() => {
        switch (level) {
            case 1: return HeadingLevel.HEADING_1;
            case 2: return HeadingLevel.HEADING_2;
            case 3: return HeadingLevel.HEADING_3;
            case 4: return HeadingLevel.HEADING_4;
            case 5: return HeadingLevel.HEADING_5;
            case 6: return HeadingLevel.HEADING_6;
            default: return HeadingLevel.HEADING_1;
        }
    })();
    ctx.idx++; // skip heading_open

    const inlineToken = ctx.tokens[ctx.idx];
    const runs = inlineToken.children
        ? renderInline(inlineToken.children, ctx)
        : [new TextRun({ text: inlineToken.content || "" })];

    ctx.idx++; // skip inline
    // heading_close handled by loop

    return new Paragraph({
        children: runs.length > 0 ? runs : [new TextRun({ text: " " })],
        heading: heading,
        spacing: { before: 240, after: 120 },
    });
}

async function parseParagraph(ctx: ConverterContext): Promise<Paragraph[]> {
    ctx.idx++;
    const inlineToken = ctx.tokens[ctx.idx];

    if (!inlineToken || inlineToken.type !== "inline") {
        ctx.idx++;
        return [];
    }

    const runs = inlineToken.children
        ? renderInline(inlineToken.children, ctx)
        : [new TextRun({ text: inlineToken.content || " " })];

    ctx.idx++;

    return splitParagraphs(runs);
}

/** 将 runs 按 break: 1 拆成多个 Paragraph，每个段落独立 */
function splitParagraphs(runs: any[]): Paragraph[] {
    if (runs.length === 0) return [];

    const paragraphs: Paragraph[] = [];
    let group: TextRun[] = [];

    for (const run of runs) {
        if ((run as any)._isBreak) {
            if (group.length > 0) {
                paragraphs.push(new Paragraph({
                    children: group,
                    alignment: AlignmentType.LEFT,
                    spacing: { after: 40 },
                }));
                group = [];
            }
        } else {
            group.push(run);
        }
    }
    if (group.length > 0) {
        paragraphs.push(new Paragraph({
            children: group,
            alignment: AlignmentType.LEFT,
            spacing: { after: 40 },
        }));
    }

    return paragraphs;
}

/**
 * 判断是否为可点击的外部链接（http/https/mailto/obsidian 协议）
 */
function isExternalLink(href: string): boolean {
    return (
        /^https?:\/\//i.test(href) ||
        /^mailto:/i.test(href) ||
        /^obsidian:/i.test(href)
    );
}

/**
 * 生成一个可点击的 Word 超链接（ExternalHyperlink 包裹 TextRun）
 */
function makeHyperlink(
    text: string,
    href: string,
    fmt: { bold?: boolean; italics?: boolean; strike?: boolean }
): any {
    return new ExternalHyperlink({
        children: [
            new TextRun({
                text,
                bold: fmt.bold,
                italics: fmt.italics,
                strike: fmt.strike,
                style: "Hyperlink",
            }),
        ],
        link: href,
    });
}

/**
 * 将 wikilink 目标解析为 obsidian:// 深度链接 URI
 * 指向 vault 内的源笔记，点击即在 Obsidian 中打开
 * 找不到对应笔记时返回 null（调用方回退为纯文本）
 */
/**
 * 将 Wikilink 目标解析为可点击的超链接 URI
 * 桌面端：file:// 绝对路径指向源 .md（Word/WPS 均能打开本地文件）
 * 移动端：obsidian:// 深度链接（移动端无文件系统路径，由 Obsidian app 处理）
 * 找不到对应笔记时返回 null（调用方回退为纯文本）
 */
function resolveHyperlinkUri(app: App, target: string): string | null {
    const file = app.metadataCache.getFirstLinkpathDest(target, "");
    if (!file) return null;

    const vault = app.vault;

    if (Platform.isMobile) {
        const vaultName = vault.getName();
        return `obsidian://open?vault=${encodeURIComponent(vaultName)}&file=${encodeURIComponent(file.path)}`;
    }

    const adapter = vault.adapter as any;
    const basePath: string = adapter.basePath;
    if (!basePath) return null;
    const absPath = `${basePath}/${file.path}`.replace(/\\/g, "/").replace(/\/{2,}/g, "/");
    return `file:///${encodeURI(absPath.replace(/^\/+/, ""))}`;
}

/**
 * 将处理过内联标记的 children 转成段落子元素数组
 * 使用格式栈追踪当前生效的格式，创建 TextRun 时一次性设置
 * 超链接（wikilink hyperlink 模式 / 外部 markdown 链接）返回 ExternalHyperlink
 */
function renderInline(
    children: Token[],
    ctx: ConverterContext
): any[] {
    const runs: any[] = [];

    // 格式栈：追踪当前嵌套的格式状态
    type Format = { bold?: boolean; italics?: boolean; strike?: boolean };
    const formatStack: Format[] = [{}];

    function currentFormat(): Format {
        return formatStack[formatStack.length - 1];
    }

    function applyFormat(type: string): void {
        const fmt = { ...currentFormat() };
        switch (type) {
            case "strong": fmt.bold = true; break;
            case "em": fmt.italics = true; break;
            case "s": fmt.strike = true; break;
        }
        formatStack.push(fmt);
    }

    function popFormat(): void {
        if (formatStack.length > 1) formatStack.pop();
    }

    // 当前生效的外部链接 href（markdown 链接 [text](url) 用）
    let currentLinkHref: string | null = null;

    for (let i = 0; i < children.length; i++) {
        const child = children[i];
        const tag = child.type;
        const text = child.content || "";

        if (tag === "link_open") {
            currentLinkHref = child.attrGet?.("href") || null;
            continue;
        }

        if (tag === "link_close") {
            currentLinkHref = null;
            continue;
        }

        const isOpening = child.nesting === 1 || tag?.endsWith("_open");
        const isClosing = child.nesting === -1 || tag?.endsWith("_close");

        if (isOpening) {
            applyFormat(tag.replace(/_open$/, ""));
            continue;
        }

        if (isClosing) {
            popFormat();
            continue;
        }

        const fmt = currentFormat();

        switch (tag) {
            case "text":
                if (text) {
                    if (currentLinkHref && isExternalLink(currentLinkHref)) {
                        runs.push(makeHyperlink(text, currentLinkHref, fmt));
                    } else {
                        runs.push(new TextRun({ text, bold: fmt.bold, italics: fmt.italics, strike: fmt.strike }));
                    }
                }
                break;

            case "code_inline":
                runs.push(new TextRun({
                    text,
                    bold: fmt.bold, italics: fmt.italics, strike: fmt.strike,
                    font: { name: "Courier New" },
                    shading: { type: ShadingType.SOLID, color: "F0F0F0", fill: "F0F0F0" },
                }));
                break;

            case "hardbreak":
            case "softbreak": {
                const br = new TextRun({ break: 1 });
                (br as any)._isBreak = true;
                runs.push(br);
                break;
            }

            case "wikilink": {
                const target = child.attrGet?.("data-target") || text;
                const alias = child.attrGet?.("data-alias");
                const displayText = alias || target;
                const isEmbed = child.attrGet?.("data-embed") === "true";

                if (isEmbed && isImageFile(target)) {
                    const imgRun = makeImageRun(ctx.imageCache.get(target));
                    if (imgRun) {
                        runs.push(imgRun);
                    } else {
                        runs.push(new TextRun({ text: `[图片: ${target}]`, italics: true, color: "999999" }));
                    }
                    break;
                }

                if (!ctx.wikilinkTargets) ctx.wikilinkTargets = [];
                if (!ctx.wikilinkTargets.includes(target)) {
                    ctx.wikilinkTargets.push(target);
                }

                if (ctx.wikilinkMode === "hyperlink") {
                    const uri = resolveHyperlinkUri(ctx.app, target);
                    if (uri) {
                        runs.push(makeHyperlink(displayText, uri, fmt));
                    } else {
                        runs.push(new TextRun({ text: displayText, bold: fmt.bold, italics: fmt.italics, strike: fmt.strike }));
                    }
                } else {
                    runs.push(new TextRun({ text: displayText, bold: fmt.bold, italics: fmt.italics, strike: fmt.strike }));
                }
                break;
            }

            case "image": {
                const imgSrc = child.attrGet?.("src") || "";
                const alt = child.attrGet?.("alt") || imgSrc || "图片";
                const imgRun = makeImageRun(ctx.imageCache.get(imgSrc));
                if (imgRun) {
                    runs.push(imgRun);
                } else {
                    runs.push(new TextRun({ text: `[图片: ${alt}]`, italics: true, color: "999999" }));
                }
                break;
            }

            case "footnote_ref": {
                const label = (child as any).meta?.label || "";
                const content = ctx.footnotes.get(label) || "";
                if (label && content) {
                    ctx.footnoteCounter++;
                    runs.push(new FootnoteReferenceRun(ctx.footnoteCounter));
                }
                break;
            }

            default:
                if (text) {
                    runs.push(new TextRun({ text, bold: fmt.bold, italics: fmt.italics, strike: fmt.strike }));
                }
                break;
        }
    }

    return runs;
}

async function parseCallout(
    ctx: ConverterContext,
    inlineToken: Token
): Promise<Paragraph> {
    // 解析 > [!TYPE] Title 格式
    const content = inlineToken.content || "";
    const calloutMatch = content.match(/^\[!(\w+)\]\s*(.*)/);
    if (!calloutMatch) {
        // 不是 callout，按普通段落处理
        const runs = inlineToken.children
            ? renderInline(inlineToken.children, ctx)
            : [new TextRun({ text: content })];
        ctx.idx++;
        return new Paragraph({
            alignment: AlignmentType.LEFT,
            children: runs,
            spacing: { after: 120 },
        });
    }

    const calloutType = calloutMatch[1].toUpperCase();
    const calloutTitle = calloutMatch[2] || calloutType;

    // Callout 颜色映射
    const colors: Record<string, { border: string; bg: string; icon: string }> = {
        NOTE: { border: "448AFF", bg: "F3F6FC", icon: "📝" },
        INFO: { border: "00BCD4", bg: "F0FAFB", icon: "ℹ️" },
        WARNING: { border: "FF9800", bg: "FFF8F0", icon: "⚠️" },
        DANGER: { border: "F44336", bg: "FFF5F5", icon: "🚨" },
        TIP: { border: "4CAF50", bg: "F4FBF4", icon: "💡" },
        IMPORTANT: { border: "9C27B0", bg: "FAF5FC", icon: "📌" },
        EXAMPLE: { border: "607D8B", bg: "F5F7F8", icon: "📋" },
        QUOTE: { border: "9E9E9E", bg: "FAFAFA", icon: "💬" },
        ABSTRACT: { border: "00ACC1", bg: "F0F9FA", icon: "📄" },
        TODO: { border: "2196F3", bg: "F3F8FD", icon: "☑️" },
        SUCCESS: { border: "4CAF50", bg: "F4FBF4", icon: "✅" },
        QUESTION: { border: "FF5722", bg: "FFF6F3", icon: "❓" },
        FAILURE: { border: "F44336", bg: "FFF5F5", icon: "❌" },
    };

    const style = colors[calloutType] || colors.NOTE;

    ctx.idx++; // skip inline
    return new Paragraph({
        alignment: AlignmentType.LEFT,
        children: [
            new TextRun({
                text: `${style.icon} ${calloutTitle}`,
                bold: true,
                color: style.border,
            }),
            new TextRun({
                text: `\n${content.replace(/^\[!\w+\]\s*/, "")}`,
            }),
        ],
        border: {
            left: { style: BorderStyle.SINGLE, size: 24, color: style.border, space: 8 },
        },
        shading: { type: ShadingType.SOLID, color: style.bg, fill: style.bg },
        indent: { left: convertInchesToTwip(0.25) },
        spacing: { after: 160, before: 80 },
    });
}

async function parseList(
    ctx: ConverterContext,
    ordered: boolean,
    depth: number = 1
): Promise<Paragraph[]> {
    const paragraphs: Paragraph[] = [];
    const listToken = ctx.tokens[ctx.idx];
    const startNumber = ordered ? (listToken as any).start || 1 : 1;
    ctx.idx++; // skip list_open

    const indentLeft = convertInchesToTwip(0.5 * depth);
    const indentHanging = convertInchesToTwip(0.25);

    let itemIndex = 0;
    while (ctx.idx < ctx.tokens.length) {
        const token = ctx.tokens[ctx.idx];

        if (token.type === "list_item_open") {
            itemIndex++;
            ctx.idx++; // skip list_item_open

            // 收集 item 内所有段落的 runs
            let itemRuns: any[] = [];
            let isFirstPara = true;
            let hasTaskPrefix = false;
            let taskChecked = false;

            while (ctx.idx < ctx.tokens.length) {
                const innerToken = ctx.tokens[ctx.idx];
                if (innerToken.type === "list_item_close") {
                    ctx.idx++; // skip list_item_close
                    break;
                }

                if (innerToken.type === "paragraph_open") {
                    ctx.idx++; // skip paragraph_open
                    const inlineToken = ctx.tokens[ctx.idx];
                    if (inlineToken?.type === "inline") {
                        const content = inlineToken.content || "";
                        const taskMatch = content.match(/^\[( |x)\]\s*/i);
                        if (taskMatch) {
                            hasTaskPrefix = true;
                            taskChecked = taskMatch[1].toLowerCase() === "x";
                        }

                        if (!isFirstPara) {
                            itemRuns.push(new TextRun({ break: 1 }));
                        }
                        // 任务列表：在源码层（inline token children）剥离 [ ] /[x] 前缀再渲染。
                        // 不要事后剥离 TextRun —— docx 运行时不暴露可读取的 .text 属性
                        // （._options/.options 均为 undefined），否则会出现「复选框 + [ ] 文本」的重复。
                        const bodyChildren = taskMatch
                            ? stripTaskPrefixFromChildren(inlineToken.children ?? [], taskMatch[0].length)
                            : inlineToken.children;
                        const runs = bodyChildren
                            ? renderInline(bodyChildren, ctx)
                            : [new TextRun({ text: content.replace(/^\[( |x)\]\s*/i, "") })];
                        itemRuns.push(...runs);
                        isFirstPara = false;
                    }
                    ctx.idx++; // skip inline
                } else if (innerToken.type === "bullet_list_open" || innerToken.type === "ordered_list_open") {
                    // 嵌套列表：先完成当前 item 的主段落
                    if (itemRuns.length > 0) {
                        paragraphs.push(makeListParagraph(itemRuns, ordered, startNumber, itemIndex, hasTaskPrefix, taskChecked, indentLeft, indentHanging));
                        itemRuns = [];
                    }
                    // 递归解析嵌套列表
                    const nestedOrdered = innerToken.type === "ordered_list_open";
                    const nestedParagraphs = await parseList(ctx, nestedOrdered, depth + 1);
                    paragraphs.push(...nestedParagraphs);
                } else {
                    ctx.idx++;
                }
            }

            // 输出 item 的主段落（如果有 runs 未输出）
            if (itemRuns.length > 0) {
                paragraphs.push(makeListParagraph(itemRuns, ordered, startNumber, itemIndex, hasTaskPrefix, taskChecked, indentLeft, indentHanging));
            }

        } else if (token.type === "list_item_close") {
            ctx.idx++;
        } else {
            break; // bullet_list_close / ordered_list_close
        }
    }

    return paragraphs;
}

/**
 * 从 inline token 的 children 中剥离任务列表前缀 `[ ]` / `[x]`（含其后空白）。
 * 之所以在源码层处理而不是事后剥离 TextRun，是因为 docx 的 TextRun 运行时不暴露
 * 可读取的 .text 属性（._options/.options 均为 undefined），事后剥离一定会失效，
 * 导致「复选框 + [ ] 文本」的重复。
 */
function stripTaskPrefixFromChildren(children: any[], prefixLen: number): any[] {
    if (!children || children.length === 0) return children;
    let toRemove = prefixLen;
    const out = children.map((c) => ({ ...c }));
    for (const c of out) {
        if (toRemove <= 0) break;
        if (typeof c.content === "string" && c.content.length > 0) {
            const take = Math.min(toRemove, c.content.length);
            c.content = c.content.slice(take);
            toRemove -= take;
        }
    }
    return out.filter((c) => !(typeof c.content === "string" && c.content.length === 0));
}

function makeListParagraph(
    runs: any[],
    ordered: boolean,
    startNumber: number,
    itemIndex: number,
    hasTaskPrefix: boolean,
    taskChecked: boolean,
    indentLeft: number,
    indentHanging: number
): Paragraph {
    if (hasTaskPrefix) {
        // 前缀已在 parseList 源码层剥离，这里直接拼接复选框与正文 run。
        return new Paragraph({
            alignment: AlignmentType.LEFT,
            children: [
                new TextRun({
                    text: taskChecked ? "☑ " : "☐ ",
                    color: taskChecked ? "4CAF50" : "9E9E9E",
                }),
                ...runs,
            ],
            spacing: { after: 60 },
            indent: { left: indentLeft, hanging: indentHanging },
        });
    }

    const number = ordered
        ? `${startNumber + itemIndex - 1}.`
        : "•";

    return new Paragraph({
        alignment: AlignmentType.LEFT,
        children: [
            new TextRun({ text: `${number}\t` }),
            ...runs,
        ],
        spacing: { after: 60 },
        indent: { left: indentLeft, hanging: indentHanging },
    });
}

async function parseBlockquote(ctx: ConverterContext): Promise<Paragraph> {
    ctx.idx++; // skip blockquote_open

    const firstToken = ctx.tokens[ctx.idx];
    if (firstToken?.type === "inline" && firstToken.content?.startsWith("[!")) {
        const calloutPara = await parseCallout(ctx, firstToken);
        return calloutPara || new Paragraph({
            alignment: AlignmentType.LEFT,
            children: [new TextRun({ text: " " })],
            border: {
                left: { style: BorderStyle.SINGLE, size: 12, color: "CCCCCC", space: 8 },
            },
            indent: { left: convertInchesToTwip(0.25) },
            spacing: { after: 120 },
        });
    }

    const allRuns: any[] = [];
    let firstContent = true;

    while (ctx.idx < ctx.tokens.length) {
        const token = ctx.tokens[ctx.idx];
      // 注意：不要在这里 ctx.idx++。parseTokens 的外层 for 循环会在 dispatchToken
      // 返回后再 idx++ 一次。若此处多跳一格，会把 blockquote 之后的那个块（标题/表格/列表）
      // 一并吞掉 —— 这是「引用格式后内容丢失」的根因。
      if (token.type === "blockquote_close") {
        break;
      }

        if (token.type === "inline") {
            if (!firstContent) {
                allRuns.push(new TextRun({ break: 1 }));
            }
            if (token.children) {
                allRuns.push(...renderInline(token.children, ctx));
            } else {
                allRuns.push(new TextRun({ text: token.content || "", italics: true, color: "666666" }));
            }
            firstContent = false;
            ctx.idx++;
        } else if (token.type === "paragraph_open") {
            ctx.idx++;
            const inlineToken = ctx.tokens[ctx.idx];
            if (inlineToken) {
                if (!firstContent) {
                    allRuns.push(new TextRun({ break: 1 }));
                }
                if (inlineToken.children) {
                    allRuns.push(...renderInline(inlineToken.children, ctx));
                } else if (inlineToken.content) {
                    allRuns.push(new TextRun({ text: inlineToken.content, italics: true, color: "666666" }));
                }
                firstContent = false;
            }
            ctx.idx++;
        } else {
            ctx.idx++;
        }
    }

    return new Paragraph({
        alignment: AlignmentType.LEFT,
        children: allRuns.length > 0 ? allRuns : [new TextRun({ text: " " })],
        border: {
            left: { style: BorderStyle.SINGLE, size: 12, color: "CCCCCC", space: 8 },
        },
        indent: { left: convertInchesToTwip(0.25) },
        spacing: { after: 120 },
    });
}

async function parseFence(
    ctx: ConverterContext,
    token: Token
): Promise<Paragraph[]> {
    const code = token.content || "";
    const language = token.info || "";

    // 防御性兜底：如果高亮链路任何一步异常，回退到原"整段灰底无着色"模式
    // 保证 .docx 至少能正常导出（即使代码块没有彩色）
    let lines: { runs: TextRun[]; isFirst: boolean; isLast: boolean }[];
    try {
        lines = highlightCodeToRuns(code, language);
    } catch (err) {
        console.warn("[export2word] highlight failed, fallback to plain code block:", err);
        const fallbackPara = new Paragraph({
            alignment: AlignmentType.LEFT,
            children: [
                new TextRun({
                    text: code,
                    font: { name: "Courier New" },
                    size: 20,
                    shading: { type: ShadingType.SOLID, color: "F5F5F5", fill: "F5F5F5" },
                }),
            ],
            shading: { type: ShadingType.SOLID, color: "F5F5F5", fill: "F5F5F5" },
            spacing: { after: 160 },
            indent: { left: convertInchesToTwip(0.25) },
            border: {
                left: { style: BorderStyle.SINGLE, size: 6, color: "E0E0E0", space: 8 },
            },
        });
        return [fallbackPara];
    }

    if (lines.length === 0) {
        const fallbackPara = new Paragraph({
            alignment: AlignmentType.LEFT,
            children: [
                new TextRun({
                    text: code,
                    font: { name: "Courier New" },
                    size: 20,
                }),
            ],
            shading: { type: ShadingType.SOLID, color: "F5F5F5", fill: "F5F5F5" },
            spacing: { after: 160 },
            indent: { left: convertInchesToTwip(0.25) },
            border: {
                left: { style: BorderStyle.SINGLE, size: 6, color: "E0E0E0", space: 8 },
            },
        });
        return [fallbackPara];
    }

    const indent = convertInchesToTwip(0.25);

    const paragraphs: Paragraph[] = lines.map(line => {
        const runs = line.runs.length > 0
            ? line.runs
            : [new TextRun({ text: " ", font: { name: "Courier New" }, size: 20 })];

        // 边框：所有段都加 left；仅首段加 top、仅末段加 bottom；构造时主动避免 undefined
        const border: any = {
            left: { style: BorderStyle.SINGLE, size: 6, color: "E0E0E0", space: 8 },
        };
        if (line.isFirst) {
            border.top = { style: BorderStyle.SINGLE, size: 6, color: "E0E0E0", space: 1 };
        }
        if (line.isLast) {
            border.bottom = { style: BorderStyle.SINGLE, size: 6, color: "E0E0E0", space: 1 };
        }

        return new Paragraph({
            alignment: AlignmentType.LEFT,
            children: runs,
            shading: { type: ShadingType.SOLID, color: "F5F5F5", fill: "F5F5F5" },
            spacing: { after: 0, before: 0, line: 280 },
            indent: { left: indent },
            border,
        });
    });

    // 代码块前后加一个空段作为视觉留白（每个空段 children 都给一个空 TextRun）
    paragraphs.unshift(
        new Paragraph({
            children: [new TextRun({ text: "" })],
            spacing: { after: 0, before: 80 },
        })
    );
    paragraphs.push(
        new Paragraph({
            children: [new TextRun({ text: "" })],
            spacing: { after: 80, before: 0 },
        })
    );

    return paragraphs;
}

async function parseTable(ctx: ConverterContext): Promise<Table> {
    const rows: TableRow[] = [];
    let isHeader = false;
    let dataRowIndex = 0; // 仅数据行计数，用于斑马纹

    ctx.idx++; // skip table_open
    while (ctx.idx < ctx.tokens.length) {
        const token = ctx.tokens[ctx.idx];

        if (token.type === "table_close") break;

        if (token.type === "thead_open") {
            isHeader = true;
            ctx.idx++;
            continue;
        }

        if (token.type === "thead_close") {
            isHeader = false;
            ctx.idx++;
            continue;
        }

        if (token.type === "tbody_open" || token.type === "tbody_close") {
            isHeader = false;
            ctx.idx++;
            continue;
        }

        if (token.type === "tr_open") {
            ctx.idx++;
            const cells: TableCell[] = [];
            // 斑马纹：仅数据行的奇数行（第 2、4、6... 行）加浅底色
            const isZebra = !isHeader && dataRowIndex % 2 === 1;

            while (ctx.idx < ctx.tokens.length) {
                const cellToken = ctx.tokens[ctx.idx];
                if (cellToken.type === "tr_close") break;

                if (cellToken.type === "th_open" || cellToken.type === "td_open") {
                    // P0-06：读取列对齐（markdown-it 以 style="text-align:center|right" 标注单元格）
                    const styleAttr = typeof cellToken.attrGet === "function"
                        ? (cellToken.attrGet("style") || "")
                        : "";
                    const cellAlign = styleAttr.includes("text-align:center")
                        ? AlignmentType.CENTER
                        : styleAttr.includes("text-align:right")
                        ? AlignmentType.RIGHT
                        : AlignmentType.LEFT;

                    ctx.idx++;
                    const inlineToken = ctx.tokens[ctx.idx];

                    // 使用 renderInline 保留单元格内的格式（加粗/斜体等）
                    let cellRuns: TextRun[];
                    if (inlineToken?.children) {
                        cellRuns = renderInline(inlineToken.children, ctx);
                    } else {
                        cellRuns = [new TextRun({ text: inlineToken?.content || "", bold: isHeader, size: 20 })];
                    }

                    const cellParagraph = new Paragraph({
                        children: cellRuns.length > 0 ? cellRuns : [new TextRun({ text: " " })],
                        alignment: cellAlign,
                    });

                    cells.push(
                        new TableCell({
                            children: [cellParagraph],
                            margins: { top: 100, bottom: 100, left: 160, right: 160 },
                            shading: isHeader
                                ? { type: ShadingType.SOLID, color: "E8EAF6", fill: "E8EAF6" }
                                : isZebra
                                ? { type: ShadingType.SOLID, color: "F4F6F8", fill: "F4F6F8" }
                                : undefined,
                        })
                    );

                    ctx.idx++; // skip inline
                    ctx.idx++; // skip th_close/td_close
                } else {
                    break;
                }
            }

            if (cells.length > 0) {
                rows.push(new TableRow({ children: cells }));
            }
            if (!isHeader) dataRowIndex++; // 仅数据行计入斑马纹
            ctx.idx++; // skip tr_close
            continue;
        }

        ctx.idx++;
    }

    return new Table({
        rows,
        width: { size: 100, type: WidthType.PERCENTAGE },
        layout: TableLayoutType.AUTOFIT,
        borders: {
            top: { style: BorderStyle.SINGLE, size: 4, color: "BFBFBF" },
            bottom: { style: BorderStyle.SINGLE, size: 4, color: "BFBFBF" },
            left: { style: BorderStyle.SINGLE, size: 4, color: "BFBFBF" },
            right: { style: BorderStyle.SINGLE, size: 4, color: "BFBFBF" },
            insideHorizontal: { style: BorderStyle.SINGLE, size: 4, color: "BFBFBF" },
            insideVertical: { style: BorderStyle.SINGLE, size: 4, color: "BFBFBF" },
        },
    });
}

// ============================================================
// 公开 API：导出
// ============================================================

export interface ExportOptions {
    app: App;
    vault: Vault;
    sourceFile: TFile;
    wikilinkMode: WikilinkMode;
    /** 递归深度（仅 zip 模式生效，默认 1） */
    recursionDepth?: number;
}

export async function exportToDocx(
    markdown: string,
    options: ExportOptions
): Promise<{ buffer: Uint8Array; wikilinkTargets: string[] }> {
    const md = createMarkdownParser();
    const tokens = md.parse(markdown, {}) as Token[];

    const ctx: ConverterContext = {
        app: options.app,
        vault: options.vault,
        sourceFile: options.sourceFile,
        wikilinkMode: options.wikilinkMode,
        recursionDepth: options.recursionDepth ?? 1,
        wikilinkTargets: [],
        idx: 0,
        tokens,
        children: [],
        imageCache: new Map(),
        footnotes: new Map(),
        footnoteCounter: 0,
    };

    // 预扫描：收集脚注定义（markdown-it-footnote 结构：footnote_open → paragraph_open → inline → paragraph_close → footnote_close）
    for (let i = 0; i < tokens.length; i++) {
        const token = tokens[i];
        if (token.type === "footnote_open") {
            const label = (token as any).meta?.label || "";
            let inlineContent = "";
            for (let j = i + 1; j < tokens.length; j++) {
                const nextToken = tokens[j];
                if (nextToken.type === "inline") {
                    inlineContent = nextToken.content || "";
                    break;
                }
                if (nextToken.type === "footnote_close") {
                    break;
                }
            }
            if (label && inlineContent) {
                ctx.footnotes.set(label, inlineContent);
            }
        }
    }

    // 解析前预取所有图片字节（renderInline 为同步函数，图片须提前读好）
    await prefetchImages(ctx, tokens);

    const children = await parseTokens(ctx);

    // 构建脚注定义（docx 要求 Record<id, children> 格式）
    const footnotesArray = Array.from(ctx.footnotes.entries());
    const footnotesRecord: Record<string, { children: Paragraph[] }> = {};
    footnotesArray.forEach(([label, content], index) => {
        footnotesRecord[String(index + 1)] = {
            children: [new Paragraph({
                children: [
                    new TextRun({ text: `${index + 1}. `, superScript: true }),
                    new TextRun({ text: content }),
                ],
            })],
        };
    });

    const doc = new Document({
        footnotes: Object.keys(footnotesRecord).length > 0 ? footnotesRecord : undefined,
        sections: [
            {
                properties: {
                    page: {
                        margin: {
                            top: convertInchesToTwip(1),
                            bottom: convertInchesToTwip(1),
                            left: convertInchesToTwip(1.25),
                            right: convertInchesToTwip(1.25),
                        },
                    },
                },
                children: children as any,
            },
        ],
        styles: {
            default: {
                document: {
                    paragraph: {
                        alignment: AlignmentType.LEFT,
                    },
                },
            },
        },
    });

    const buffer = await Packer.toBuffer(doc);
    return { buffer: new Uint8Array(buffer), wikilinkTargets: ctx.wikilinkTargets };
}

/**
 * 轻量级链接提取：仅解析 markdown 并收集 wikilink 目标，不执行完整 docx 转换
 * 用于 collectAttachmentTree 性能优化
 */
export function extractWikilinkTargets(markdown: string): string[] {
    const md = createMarkdownParser();
    const tokens = md.parse(markdown, {}) as Token[];
    const targets: string[] = [];

    function walkTokens(tokens: Token[]): void {
        for (const token of tokens) {
            if (token.type === "wikilink") {
                const target = (token as any).attrGet?.("data-target") || token.content || "";
                if (target && !targets.includes(target)) {
                    targets.push(target);
                }
            }
            if (token.children) {
                walkTokens(token.children);
            }
        }
    }

    walkTokens(tokens);
    return targets;
}
