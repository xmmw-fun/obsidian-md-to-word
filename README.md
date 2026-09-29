<!-- 发版维护提醒：每次发版必须同步三处版本号 —— manifest.json 的 version、package.json 的 version、versions.json 追加一行 "新版本": "minAppVersion"；然后打同名 tag 的 GitHub Release 并上传 main.js / manifest.json / styles.css 三件套。README 只有功能/用法变化时才需要改。 -->

# MD to Word

**Export your notes to Word (.docx) with one click.** Full Markdown rendering — tables, embedded images, syntax-highlighted code blocks, footnotes, task lists, and wikilink attachment bundles. Zero external dependencies. Works on desktop and mobile.

一键将笔记导出为 Word（.docx）。完整 Markdown 渲染，零外部依赖，桌面端与移动端均可用。

## Features

- **Complete Markdown rendering**: heading hierarchy, bold/italic/inline code, ordered/unordered lists, blockquotes, horizontal rules
- **Tables**: content-based autofit column widths, shaded header, zebra striping, thin gray borders, per-column alignment
- **Images**: `![[image.png]]` (including subfolders and `../` relative paths), standard Markdown `![](path)` (including non-ASCII file names), and remote images over HTTPS; auto-scaled to page width
- **Syntax highlighting**: token-level colors for code blocks (keywords, strings, comments, numbers) in a continuous shaded code block
- **Footnotes**: `[^label]` exported as native Word footnotes
- **Task lists**: `- [ ]` / `- [x]` rendered as colored check symbols
- **Wikilinks**: plain-text mode, or attachment-bundle mode — recursively export linked notes and pack everything into a numbered `.zip` with an attachment index
- **Batch export**: right-click a folder to export all its notes (optionally merged into a single `.zip`)
- **Export location**: `_exports` folder, source-note folder, or a custom path
- **Cross-platform**: no Node/Electron APIs — desktop and mobile

## Installation

**Community plugin directory** (once approved): Settings → Community plugins → Browse → search "MD to Word".

**Manual**:

1. Download `main.js`, `manifest.json`, and `styles.css` from the [latest release](https://github.com/xmmw-fun/obsidian-md-to-word/releases/latest)
2. Copy them into `<YourVault>/.obsidian/plugins/md-to-word/`
3. Enable "MD to Word" under Settings → Community plugins

**Beta via BRAT**: add `xmmw-fun/obsidian-md-to-word` to the BRAT plugin.

## Usage

- Command palette: `MD to Word: 导出为 Word (.docx)`
- Ribbon icon (file-down arrow on the left sidebar)
- File menu: right-click a note → `导出为 Word (.docx)`
- Folder menu: right-click a folder → `批量导出文件夹为 Word (.docx)`
- Hotkey: none by default — bind one in Settings → Hotkeys by searching "导出为 Word"

## Settings

| Setting | Description |
|---------|-------------|
| Wikilink mode | Plain text / Attachment bundle (.zip, recursion depth 0–3) |
| Export path | `_exports` folder / source-note folder / custom path |
| Batch export as single .zip | Merge batch results into one archive |
| Attachment index | List linked documents hierarchically in zip mode |

## Network requests

This plugin makes network requests in exactly one case: when a note references a **remote image** (`![](https://...)`), the image is downloaded via Obsidian's official `requestUrl` API during export so it can be embedded in the .docx. There is no other network activity, no telemetry, and no remote code execution.

## Tech stack

TypeScript · [docx.js](https://docx.js.org/) · markdown-it · highlight.js · jszip

## Roadmap

- Custom .docx templates
- Pre-export health check (broken links / missing images)
- LaTeX equation export
- Clickable wikilink mode (under investigation)
- UI translations

## 中文说明

**MD to Word** 是一键把笔记导出成 Word 文档的插件：

- 表格自适应列宽、斑马纹；代码块语法高亮；脚注转为 Word 原生脚注；`- [ ]` 待办渲染为勾选符号
- 图片全格式支持：Obsidian 附件 `![[图.png]]`（含 `../` 相对路径）、Markdown 图片（含中文文件名）、网络图片
- Wikilink 两种模式：纯文本，或「附件打包」——递归导出关联笔记打成带编号和目录的 .zip
- 右键文件夹可批量导出；导出路径可选 `_exports` / 源文件目录 / 自定义
- 无默认热键，可在「设置 → 快捷键」中搜索「导出为 Word」自行绑定

## License

[MIT](LICENSE)
