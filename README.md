<!-- 发版维护提醒：每次发版必须同步三处版本号 —— manifest.json 的 version、package.json 的 version、versions.json 追加一行 "新版本": "minAppVersion"；然后打同名 tag 的 GitHub Release 并上传 main.js / manifest.json / styles.css 三件套。README 只有功能/用法变化时才需要改。 -->

# Export to Word

一键将 Obsidian Markdown 笔记导出为 Word（.docx）文档，零外部依赖，桌面端与移动端均可使用。

Export Obsidian notes to Word (.docx) with one click. Zero external dependencies. Works on desktop and mobile.

## Features / 功能特性

- **完整的 Markdown 渲染**：标题层级、加粗/斜体/行内代码、有序/无序列表、引用块、分隔线
- **表格**：按内容自适应列宽，表头底色 + 斑马纹 + 细灰边框 + 单元格对齐
- **图片嵌入**：支持 `![[image.png]]`（含子目录与 `../` 相对路径）、标准 Markdown `![](path)`（含中文文件名）、网络图片；超出页宽自动等比缩放
- **代码块语法高亮**：token 级着色（关键字/字符串/注释/数字等），连续灰底代码块样式
- **脚注**：`[^label]` 语法导出为 Word 原生脚注
- **待办事项**：`- [ ]` / `- [x]` 渲染为带颜色的勾选符号
- **Wikilink 处理**：纯文本模式，或「附件打包」模式——将关联笔记递归导出并连同主文档一起打包为 .zip，自动编号并生成附件目录
- **批量导出**：右键文件夹即可批量导出整个文件夹为 .docx（可选合并为一个 .zip）
- **导出路径可选**：Vault 根目录 `_exports`、源文件所在目录，或自定义路径
- **全平台**：不依赖 Node/Electron API，桌面端与移动端均可导出

## Installation / 安装

**社区市场（审核通过后）**：在 Obsidian 设置 → 第三方插件 → 社区插件市场中搜索 "Export to Word" 安装。

**手动安装**：

1. 从 [Releases](https://github.com/xmmw-fun/obsidian-export-to-word/releases) 下载 `main.js`、`manifest.json`、`styles.css`
2. 放入 `<你的Vault>/.obsidian/plugins/export-to-word/` 目录
3. 在 Obsidian 设置 → 第三方插件中启用

**BRAT 内测**：将 `xmmw-fun/obsidian-export-to-word` 添加到 BRAT 插件。

## Usage / 使用

- 命令面板：`Export to Word: 导出为 Word (.docx)`
- 左侧功能区图标（向下箭头文件图标）
- 右键文件菜单：`导出为 Word (.docx)`
- 右键文件夹菜单：`批量导出文件夹为 Word (.docx)`
- 默认快捷键：未绑定，可在「设置 → 快捷键」中搜索 "Export to Word" 自行绑定（也可在插件设置中一键启用 Ctrl/Cmd+Shift+E）

## Settings / 设置

| 设置项 | 说明 |
|--------|------|
| Wikilink 处理方式 | 纯文本 / 附件打包（.zip，含递归深度 0–3 层） |
| 导出路径 | `_exports` 目录 / 源文件所在目录 / 自定义路径 |
| 批量导出为单个 .zip | 批量导出时合并打包 |
| 附件目录 | 在 zip 模式下列出关联文档层级清单 |
| 启用导出快捷键 | 开启后按 Ctrl/Cmd+Shift+E 导出当前文件 |

## Network Requests / 网络请求说明

本插件唯一的网络行为：当笔记引用**网络图片**（`![](https://...)`）时，导出过程中会通过 Obsidian 官方 `requestUrl` API 下载该图片以嵌入 .docx。除此之外没有任何网络请求、没有遥测、没有远程代码。

## Tech Stack / 技术路线

TypeScript · [docx.js](https://docx.js.org/) · markdown-it · highlight.js · jszip

## Roadmap / 路线图

- [ ] 自定义 .docx 模板
- [ ] 导出前智能检查（损坏链接 / 缺失图片提示）
- [ ] LaTeX 公式导出
- [ ] 超链接模式（Wikilink 导出为可点击链接，研究中）
- [ ] UI 国际化

## License

[MIT](LICENSE)
