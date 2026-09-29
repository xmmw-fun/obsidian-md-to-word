/**
 * md-to-word — Obsidian 插件主入口
 * v1.0.0
 *
 * 在 Obsidian 中一键将 .md 文件导出为 .docx
 */

import {
    App,
    Notice,
    Plugin,
    PluginSettingTab,
    Setting,
    TFile,
    TFolder,
    SuggestModal,
    normalizePath,
    Platform,
} from "obsidian";
import { exportToDocx, extractWikilinkTargets } from "./converter";
import {
    DEFAULT_SETTINGS,
    type Export2WordSettings,
    type WikilinkMode,
} from "./settings";

/** 附件目录树节点（用于层级编号 1 / 1.1 / 1.2） */
interface AttachmentNode {
    file: TFile;
    number: string;
    children: AttachmentNode[];
}

export default class Export2WordPlugin extends Plugin {
    declare settings: Export2WordSettings;

    async onload(): Promise<void> {
        await this.loadSettings();

        // ==========================================================
        // 命令：导出当前文件为 Word
        // ==========================================================
        this.addCommand({
            id: "export-to-word",
            name: "导出为 Word (.docx)",
            icon: "file-text",
            callback: () => this.exportCurrentFile(),
        });

        // ==========================================================
        // 命令：导出当前文件为 Word（使用附件打包模式）
        // ==========================================================
        this.addCommand({
            id: "export-to-word-zip",
            name: "导出为 Word 压缩包（含附件）",
            icon: "package",
            callback: () => this.exportCurrentFile("zip"),
        });

        // ==========================================================
        // 命令：批量导出整个文件夹为 Word
        // ==========================================================
        this.addCommand({
            id: "export-folder-to-word",
            name: "批量导出文件夹为 Word (.docx)",
            icon: "folder-down",
            callback: () => this.openFolderPicker(),
        });

        // ==========================================================
        // 功能区图标
        // ==========================================================
        this.addRibbonIcon("file-down", "导出为 Word", () => {
            void this.exportCurrentFile();
        });

        // ==========================================================
        // 文件菜单
        // ==========================================================
        this.registerEvent(
            this.app.workspace.on("file-menu", (menu, file) => {
                // 文件夹：批量导出
                if (file instanceof TFolder) {
                    menu.addItem((item) => {
                        item
                            .setTitle("批量导出文件夹为 Word (.docx)")
                            .setIcon("folder-down")
                            .onClick(() => this.exportFolder(file));
                    });
                    return;
                }

                if (!(file instanceof TFile) || file.extension !== "md") return;

                menu.addItem((item) => {
                    item
                        .setTitle("导出为 Word (.docx)")
                        .setIcon("file-text")
                        .onClick(() => this.exportFile(file));
                });

                menu.addSeparator();
            })
        );

        // ==========================================================
        // 设置页面
        // ==========================================================
        this.addSettingTab(new Export2WordSettingTab(this.app, this));
    }

    onunload(): void {
        // 清理工作
    }

    // ==========================================================
    // 核心导出逻辑
    // ==========================================================

    async exportCurrentFile(forceMode?: WikilinkMode): Promise<void> {
        const file = this.app.workspace.getActiveFile();
        if (!file) {
            new Notice("⚠️ 没有打开的文件");
            return;
        }

        if (file.extension !== "md") {
            new Notice("⚠️ 当前文件不是 Markdown 文件");
            return;
        }

        await this.exportFile(file, forceMode);
    }

    async exportFile(file: TFile, forceMode?: WikilinkMode): Promise<void> {
        const notice = new Notice("正在导出...", 0);

        try {
            const rawMode = forceMode || this.settings.wikilinkMode;
            const wikilinkMode: WikilinkMode = rawMode === "zip" ? "zip" : "plain";
            // openAfter=true：单文件导出后自动打开
            await this.exportOneFile(file, wikilinkMode, { openAfter: true });
            new Notice(`✅ 导出成功`);
        } catch (error) {
            console.error("export2word error:", error);
            new Notice(`❌ 导出失败：${error instanceof Error ? error.message : "未知错误"}`);
        } finally {
            notice.hide();
        }
    }

    /**
     * 核心：读取一个 .md 并转换 + 保存（单文件与批量模式共用）。
     * @param mode "zip" 走附件打包；其余统一视为 "plain" 普通 .docx
     * @param opts.openAfter 保存后是否在 Obsidian 中自动打开（批量模式传 false，避免卡死 UI）
     * @param opts.exportDirOverride 自定义导出目录（批量模式用它保持原文件夹层级）；不传则按设置决定
     */
    async exportOneFile(
        file: TFile,
        mode: WikilinkMode,
        opts: { openAfter: boolean; exportDirOverride?: string; includeAttachmentIndex?: boolean; exportNameOverride?: string },
    ): Promise<void> {
        const content = await this.app.vault.read(file);

        if (mode === "zip") {
            // Zip 模式：打包主文档 + 所有关联文档
            await this.exportZip(file, content, {
                openAfter: opts.openAfter,
                exportDirOverride: opts.exportDirOverride,
            });
            return;
        }

        // 普通模式：单个 .docx（超链接模式暂未实现，兜底纯文本）
        const wantIndex =
            opts.includeAttachmentIndex !== false && this.settings.attachmentIndexEnabled;

        // 附件目录：先收集关联文档树（仅一次），末尾追加 / 单独文档共用；
        // 不递归（recursionDepth<=0）或没有任何 link 时返回空树
        const attachmentTree =
            wantIndex ? await this.collectAttachmentTree(file, this.settings.recursionDepth) : [];

        // 附件目录（末尾追加）：把目录拼进正文，一次性转换
        let docxContent = content;
        if (attachmentTree.length > 0 && this.settings.attachmentIndexLocation === "end") {
            const appendix =
                "\n\n---\n\n## 附件目录\n\n" + this.attachmentTreeToMarkdown(attachmentTree);
            docxContent = content.endsWith("\n")
                ? content + appendix
                : content + "\n" + appendix;
        }

        const result = await exportToDocx(docxContent, {
            app: this.app,
            vault: this.app.vault,
            sourceFile: file,
            wikilinkMode: "plain",
        });

        const exportName = opts.exportNameOverride ?? `${file.basename}.docx`;

        if (Platform.isMobile) {
            await this.saveOnMobile(result.buffer, exportName);
        } else {
            await this.saveOnDesktop(result.buffer, exportName, file, {
                openAfter: opts.openAfter,
                exportDirOverride: opts.exportDirOverride,
            });
        }

        // 附件目录（单独文档）：额外生成一个「附件目录.docx」，与正文同导出位置
        if (attachmentTree.length > 0 && this.settings.attachmentIndexLocation === "separate") {
            const md = `# ${file.basename} · 附件目录\n\n` + this.attachmentTreeToMarkdown(attachmentTree);
            const idxResult = await exportToDocx(md, {
                app: this.app,
                vault: this.app.vault,
                sourceFile: file,
                wikilinkMode: "plain",
            });
            const idxName = `${file.basename}-附件目录.docx`;
            const exportDir = opts.exportDirOverride ?? this.resolveExportBaseDir(file);
            await this.saveOnDesktop(idxResult.buffer, idxName, file, {
                openAfter: false,
                exportDirOverride: exportDir,
            });
        }

        // 附件目录开启且存在关联文档时：除文本清单外，把每个被 link 的子文档
        // 实际导出为独立的 .docx 文件，归入「<主文档名>附件」子文件夹（如 A附件/甲.docx）。
        // 子文档自身递归关闭（includeAttachmentIndex:false），避免无限展开。
        if (attachmentTree.length > 0) {
            const exportDir = opts.exportDirOverride ?? this.resolveExportBaseDir(file);
            const attachmentSubDir = normalizePath(`${exportDir}/${file.basename}附件`);
            const flatNodes: AttachmentNode[] = [];
            const collectFlat = (nodes: AttachmentNode[]): void => {
                for (const n of nodes) {
                    flatNodes.push(n);
                    if (n.children.length) collectFlat(n.children);
                }
            };
            collectFlat(attachmentTree);
            let okCount = 0;
            for (const node of flatNodes) {
                try {
                    // 文件名带上与清单一致的层级编号（附件1-甲.docx / 附件1.1-乙.docx），方便对照
                    const attName = `附件${node.number}-${node.file.basename}.docx`;
                    await this.exportOneFile(node.file, "plain", {
                        openAfter: false,
                        exportDirOverride: attachmentSubDir,
                        includeAttachmentIndex: false,
                        exportNameOverride: attName,
                    });
                    okCount++;
                } catch {
                    // 单个附件导出失败不影响主流程
                    continue;
                }
            }
            if (okCount > 0) {
                new Notice(`📎 已导出 ${okCount} 个关联文档到「${file.basename}附件」`);
            }
        }
    }

    /**
     * 批量导出整个文件夹（含所有子文件夹）的 .md。
     * - 递归收集所有 .md
     * - 路径走「默认导出路径」统一设置（与单文件共用）：在其下再建以文件夹名命名的子目录
     * - batchZipMode 关闭（默认）：每个 .md 独立 .docx，按原层级写入
     * - batchZipMode 开启：所有 .md 转 .docx 后汇总压成「一个」zip（保持原目录层级）
     * - 统一进度 Notice；单个文件失败跳过并汇总
     */
    async exportFolder(folder: TFolder): Promise<void> {
        // 递归收集所有 .md 文件
        const files: TFile[] = [];
        const collect = (f: TFolder): void => {
            for (const child of f.children) {
                if (child instanceof TFolder) {
                    collect(child);
                } else if (child instanceof TFile && child.extension === "md") {
                    files.push(child);
                }
            }
        };
        collect(folder);

        if (files.length === 0) {
            new Notice("⚠️ 该文件夹内没有可导出的 .md 文件");
            return;
        }

        const folderName = folder.name || this.app.vault.getName();
        // 统一默认路径（与单文件导出共用同一设置）
        const baseExportDir = this.resolveExportBaseDir(folder);
        const isMobile = Platform.isMobile;
        const batchZip = isMobile ? true : this.settings.batchZipMode;
        // $sourceDir$ 时 baseExportDir 已是被导出文件夹自身，非 zip 模式不再拼 folderName（避免嵌套）；
        // vault 根目录（无 path）仍走拼接归组
        const isSourceDir = this.settings.exportPath === "$sourceDir$" && !!folder.path;

        // 非 zip 模式：在导出基础目录下建以文件夹名命名的子目录（已存在则加时间戳）
        let baseDir = "";
        if (!batchZip) {
            if (isSourceDir) {
                // $sourceDir$：baseDir 即源文件夹自身（docx 按层级生成在其内，与源 md 并列）
                baseDir = baseExportDir;
            } else {
                baseDir = normalizePath(`${baseExportDir}/${folderName}`);
                if (await this.app.vault.adapter.exists(baseDir)) {
                    const ts = new Date().toISOString().replace(/[:.]/g, "-").slice(0, 19);
                    baseDir = normalizePath(`${baseExportDir}/${folderName}_${ts}`);
                }
            }
        }

        const failed: { name: string; error: string }[] = [];
        let done = 0;
        const notice = new Notice(
            batchZip
                ? `正在批量导出 0 / ${files.length}（汇总为单个 zip）...`
                : `正在批量导出 0 / ${files.length}...`,
            0,
        );

        // zip 模式：收集所有 docx 字节，最后压成一个 zip
        const docxEntries: { relPath: string; buffer: Uint8Array }[] = [];

        for (const file of files) {
            try {
                // 相对路径（保持原目录层级）
                const base = folder.path ? folder.path + "/" : "";
                const rel = base ? file.path.slice(base.length) : file.path;

                if (batchZip) {
                    // 转 docx 字节，收集进 zip 列表（不落盘单文件）
                    const content = await this.app.vault.read(file);
                    const result = await exportToDocx(content, {
                        app: this.app,
                        vault: this.app.vault,
                        sourceFile: file,
                        wikilinkMode: "plain",
                    });
                    const relDocx = rel.replace(/\.md$/i, ".docx");
                    docxEntries.push({ relPath: relDocx, buffer: result.buffer });
                } else {
                    // 普通模式：每个 .md 独立 .docx，按层级写入 baseDir
                    const relDir = rel.includes("/")
                        ? rel.slice(0, rel.lastIndexOf("/"))
                        : "";
                    const targetDir = relDir
                        ? normalizePath(`${baseDir}/${relDir}`)
                        : baseDir;
                    await this.exportOneFile(file, "plain", {
                        openAfter: false,
                        exportDirOverride: targetDir,
                        includeAttachmentIndex: false,
                    });
                }
            } catch (error) {
                failed.push({
                    name: file.basename,
                    error: error instanceof Error ? error.message : "未知错误",
                });
            }
            done++;
            notice.setMessage(`正在批量导出 ${done} / ${files.length}...`);
        }

        // zip 模式：把所有 docx 压成一个 zip 落盘
        if (batchZip) {
            try {
                const JSZip = (await import("jszip")).default;
                const zip = new JSZip();
                for (const e of docxEntries) {
                    zip.file(e.relPath, e.buffer);
                }
                const zipBuffer = await zip.generateAsync({ type: "uint8array" });
                const zipPath = await this.writeBinarySafe(
                    baseExportDir,
                    `${folderName}.zip`,
                    zipBuffer,
                );
                notice.hide();
                if (failed.length === 0) {
                    new Notice(
                        `✅ 批量导出完成：${docxEntries.length} 个文件已汇总为一个 zip → ${zipPath}`,
                        8000,
                    );
                } else {
                    const detail = failed
                        .map((f) => `• ${f.name}：${f.error}`)
                        .join("\n");
                    new Notice(
                        `⚠️ 批量导出完成：${docxEntries.length} 个打包成功，${failed.length} 个失败 → ${zipPath}\n${detail}`,
                        15000,
                    );
                }
            } catch (error) {
                notice.hide();
                new Notice(
                    `❌ 打包 zip 失败：${error instanceof Error ? error.message : "未知错误"}`,
                    10000,
                );
            }
            return;
        }

        notice.hide();
        if (failed.length === 0) {
            new Notice(
                `✅ 批量导出完成：共 ${files.length} 个文件 → ${baseDir}`,
                8000,
            );
        } else {
            const detail = failed
                .map((f) => `• ${f.name}：${f.error}`)
                .join("\n");
            new Notice(
                `⚠️ 批量导出完成：${files.length - failed.length}/${files.length} 成功，${failed.length} 个失败：\n${detail}`,
                15000,
            );
        }
    }

    /** 打开文件夹选择器（命令面板触发） */
    async openFolderPicker(): Promise<void> {
        new FolderSuggestModal(this.app, this).open();
    }

    /** Zip 模式：导出主文档 + 所有链接的文档为一个 .zip 包 */
    async exportZip(
        file: TFile,
        content: string,
        opts?: { openAfter?: boolean; exportDirOverride?: string; fileName?: string },
    ): Promise<void> {
        // 第一步：导出主文档，收集 wikilink targets
        const mainResult = await exportToDocx(content, {
            app: this.app,
            vault: this.app.vault,
            sourceFile: file,
            wikilinkMode: "plain", // zip 内文档用普通模式
        });

        const targets = mainResult.wikilinkTargets;
        const visited = new Set<string>([file.path]);

        // 第二步：递归导出所有关联文档
        const attachments: { name: string; buffer: Uint8Array }[] = [];
        const depth = this.settings.recursionDepth;

        await this.collectLinkedDocs(targets, 1, depth, visited, attachments, file.path);

        // 第三步：用 JSZip 打包
        const JSZip = (await import("jszip")).default;
        const zip = new JSZip();

        // 主文档追加附件目录（无关联文档时「不递归」纯导出原文档，不加附录）
        let appendix = "";
        if (attachments.length > 0) {
            appendix = "\n\n---\n\n## 附件目录\n\n";
            for (let i = 0; i < attachments.length; i++) {
                appendix += `${i + 1}. ${attachments[i].name}\n`;
            }
        }

        // 重新导出主文档（带附件目录）
        const fullContent = content.endsWith("\n") ? content + appendix : content + "\n" + appendix;
        const finalResult = await exportToDocx(fullContent, {
            app: this.app,
            vault: this.app.vault,
            sourceFile: file,
            wikilinkMode: "plain",
        });

        zip.file(`${file.basename}.docx`, finalResult.buffer);
        for (const att of attachments) {
            zip.file(att.name, att.buffer);
        }

        const zipBuffer = await zip.generateAsync({ type: "uint8array" });

        const zipName = opts?.fileName ?? `${file.basename}_export2word.zip`;
        const saveOpts = {
            openAfter: opts?.openAfter ?? true,
            exportDirOverride: opts?.exportDirOverride,
        };

        if (Platform.isMobile) {
            await this.saveOnMobile(zipBuffer, zipName);
        } else {
            await this.saveOnDesktop(zipBuffer, zipName, file, saveOpts);
        }
    }

    /** 递归收集关联文档
     * @param sourcePath 当前文档路径，用于 Obsidian 最短路径链接解析
     */
    async collectLinkedDocs(
        targets: string[],
        currentDepth: number,
        maxDepth: number,
        visited: Set<string>,
        attachments: { name: string; buffer: Uint8Array }[],
        sourcePath: string,
    ): Promise<void> {
        if (currentDepth > maxDepth) return;

        for (const target of targets) {
            // 使用 Obsidian 原生链接解析（最短路径策略）
            // 例：[[v1.0.1-issues]] → docs/log/v1.0.1-issues.md
            const mdFile = this.app.metadataCache.getFirstLinkpathDest(target, sourcePath);

            if (!mdFile || !(mdFile instanceof TFile) || visited.has(mdFile.path)) continue;
            visited.add(mdFile.path);

            try {
                const linkedContent = await this.app.vault.read(mdFile);
                const linkedResult = await exportToDocx(linkedContent, {
                    app: this.app,
                    vault: this.app.vault,
                    sourceFile: mdFile,
                    wikilinkMode: "plain",
                });

                attachments.push({
                    name: `附件${attachments.length + 1}-${mdFile.basename}.docx`,
                    buffer: linkedResult.buffer,
                });

                // 递归收集下一层
                if (currentDepth < maxDepth && linkedResult.wikilinkTargets.length > 0) {
                    await this.collectLinkedDocs(
                        linkedResult.wikilinkTargets,
                        currentDepth + 1,
                        maxDepth,
                        visited,
                        attachments,
                        mdFile.path,
                    );
                }
            } catch {
                // 跳过无法导出的文件
                continue;
            }
        }
    }

    // ==========================================================
    // 桌面端保存
    // ==========================================================

    /**
     * 解析导出基础目录（单文件与批量导出共用同一设置）。
     * - "$sourceDir$"：
     *   - 单文件 → 源文件所在目录
     *   - 批量文件夹 → 该文件夹自身内部（zip/docx 放进文件夹里，与单文件"docx 放 md 旁边"语义一致）
     * - 空：vault 根目录的 _exports
     * - 其他：自定义路径（相对 vault 根）
     */
    /**
     * 树状收集关联文档（保留层级，用于附件目录编号）。
     * 链接提取统一走 converter 的 wikilinkTargets（与 zip 打包同一套解析），
     * 避免纯文本模式依赖 metadataCache 时因缓存时机不一致抓不到链接的问题。
     * maxDepth <= 0 表示「不递归」，仅导出原文档，返回空树。
     */
    private async collectAttachmentTree(root: TFile, maxDepth: number): Promise<AttachmentNode[]> {
        // 不递归：仅原文档，不展开任何 link
        if (maxDepth <= 0) return [];

        const visited = new Set<string>([root.path]);

        // 使用轻量级链接提取，避免完整 docx 转换
        const linkCache = new Map<string, string[]>();
        const getLinks = async (f: TFile): Promise<string[]> => {
            const cached = linkCache.get(f.path);
            if (cached) return cached;
            const content = await this.app.vault.read(f);
            const targets = extractWikilinkTargets(content);
            linkCache.set(f.path, targets);
            return targets;
        };

        const resolve = (link: string, fromPath: string): TFile | null => {
            const f = this.app.metadataCache.getFirstLinkpathDest(link, fromPath);
            return f instanceof TFile ? f : null;
        };

        const roots: AttachmentNode[] = [];
        for (const link of await getLinks(root)) {
            const f = resolve(link, root.path);
            if (!f || visited.has(f.path)) continue;
            visited.add(f.path);
            roots.push({ file: f, number: "", children: [] });
        }

        let currentLevel = roots;
        let depth = 2;
        while (depth <= maxDepth && currentLevel.length > 0) {
            const nextLevel: AttachmentNode[] = [];
            for (const parent of currentLevel) {
                for (const link of await getLinks(parent.file)) {
                    const f = resolve(link, parent.file.path);
                    if (!f || visited.has(f.path)) continue;
                    visited.add(f.path);
                    const child: AttachmentNode = { file: f, number: "", children: [] };
                    parent.children.push(child);
                    nextLevel.push(child);
                }
            }
            currentLevel = nextLevel;
            depth++;
        }

        const numberize = (nodes: AttachmentNode[], prefix: string | null): void => {
            nodes.forEach((n, i) => {
                n.number = prefix ? `${prefix}.${i + 1}` : `${i + 1}`;
                if (n.children.length) numberize(n.children, n.number);
            });
        };
        numberize(roots, null);
        return roots;
    }

    /** 将附件树渲染为 markdown 文本（每行：附件<编号> <文档名>） */
    private attachmentTreeToMarkdown(tree: AttachmentNode[]): string {
        if (tree.length === 0) return "（无关联文档）\n";
        const lines: string[] = [];
        const walk = (nodes: AttachmentNode[]): void => {
            for (const n of nodes) {
                lines.push(`附件${n.number}-${n.file.basename}`);
                if (n.children.length) walk(n.children);
            }
        };
        walk(tree);
        return lines.join("\n") + "\n";
    }

    resolveExportBaseDir(source: TFile | TFolder): string {
        const p = this.settings.exportPath;
        if (p === "$sourceDir$") {
            if (source instanceof TFolder) {
                // 批量导出文件夹：导出到该文件夹自身内部
                // vault 根目录无 path，回退 _exports 避免写到 vault 根
                return source.path ? source.path : normalizePath("_exports");
            }
            // 单文件：源文件所在目录
            const parent = source.parent;
            return parent && parent.path ? parent.path : normalizePath("_exports");
        }
        if (!p) {
            return normalizePath("_exports");
        }
        return normalizePath(p);
    }

    /** 写入二进制到指定目录，若同名已存在则加时间戳避免覆盖；返回最终写入路径 */
    async writeBinarySafe(
        dir: string,
        fileName: string,
        buffer: Uint8Array,
    ): Promise<string> {
        await this.ensureDir(dir);
        let path = normalizePath(`${dir}/${fileName}`);
        if (await this.app.vault.adapter.exists(path)) {
            const timestamp = new Date().toISOString().replace(/[:.]/g, "-").slice(0, 19);
            const extMatch = fileName.match(/\.(docx|zip)$/i);
            const ext = extMatch ? extMatch[0] : "";
            const nameWithoutExt = ext ? fileName.slice(0, -ext.length) : fileName;
            path = normalizePath(`${dir}/${nameWithoutExt}_${timestamp}${ext}`);
        }
        await this.app.vault.adapter.writeBinary(path, buffer.buffer as ArrayBuffer);
        return path;
    }

    async saveOnDesktop(
        buffer: Uint8Array,
        fileName: string,
        sourceFile: TFile,
        opts?: { openAfter?: boolean; exportDirOverride?: string },
    ): Promise<void> {
        const openAfter = opts?.openAfter ?? true;

        // 批量模式用调用方指定目录；否则用统一设置的默认路径
        const exportDir = opts?.exportDirOverride
            ? opts.exportDirOverride
            : this.resolveExportBaseDir(sourceFile);

        const finalPath = await this.writeBinarySafe(exportDir, fileName, buffer);

        // 自动打开（仅单文件模式）
        if (openAfter) {
            const abstract = this.app.vault.getAbstractFileByPath(finalPath);
            if (abstract instanceof TFile) {
                await this.app.workspace.getLeaf(false).openFile(abstract);
            }
        }
    }

    /** 递归创建目录（Obsidian adapter.mkdir 不保证创建父级目录） */
    async ensureDir(dirPath: string): Promise<void> {
        const parts = dirPath.split("/").filter((p) => p.length > 0);
        let cur = "";
        for (const p of parts) {
            cur = cur ? `${cur}/${p}` : p;
            if (!(await this.app.vault.adapter.exists(cur))) {
                await this.app.vault.adapter.mkdir(cur);
            }
        }
    }

    // ==========================================================
    // 移动端保存（使用 Data URI 下载）
    // ==========================================================

    async saveOnMobile(buffer: Uint8Array, fileName: string): Promise<void> {
        const blob = new Blob([buffer.buffer as ArrayBuffer], {
            type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        });

        const url = URL.createObjectURL(blob);
        const a = document.body.createEl("a", { cls: "md-to-word-download-hidden" });
        a.href = url;
        a.download = fileName;
        a.click();

        window.setTimeout(() => {
            a.remove();
            URL.revokeObjectURL(url);
        }, 100);

        await new Promise((resolve) => window.setTimeout(resolve, 500));
    }

    // ==========================================================
    // 设置管理
    // ==========================================================

    async loadSettings(): Promise<void> {
        this.settings = Object.assign({}, DEFAULT_SETTINGS, await this.loadData());
        // 迁移：超链接模式已隐藏（WPS 不触发自定义协议），旧设置值回退为纯文本
        if (this.settings.wikilinkMode === "hyperlink") {
            this.settings.wikilinkMode = "plain";
        }
    }

    async saveSettings(): Promise<void> {
        await this.saveData(this.settings);
    }
}

// ============================================================
// 文件夹选择器（命令面板选文件夹用）
// ============================================================

class FolderSuggestModal extends SuggestModal<string> {
    plugin: Export2WordPlugin;
    folders: string[];

    constructor(app: App, plugin: Export2WordPlugin) {
        super(app);
        this.plugin = plugin;
        this.setPlaceholder("输入或选择要导出的文件夹...");
        // 收集 vault 内所有文件夹路径
        this.folders = app.vault
            .getAllLoadedFiles()
            .filter((f) => f instanceof TFolder)
            .map((f) => f.path)
            .sort();
    }

    getSuggestions(query: string): string[] {
        const q = query.toLowerCase();
        return q
            ? this.folders.filter((p) => p.toLowerCase().includes(q))
            : this.folders;
    }

    renderSuggestion(value: string, el: HTMLElement): void {
        el.setText(value === "" ? "（vault 根目录）" : value);
    }

    onChooseSuggestion(folderPath: string, _evt?: MouseEvent | KeyboardEvent): void {
        const folder = this.app.vault.getAbstractFileByPath(folderPath);
        if (folder instanceof TFolder) {
            void this.plugin.exportFolder(folder);
        }
    }
}

// ============================================================
// 设置面板
// ============================================================

class Export2WordSettingTab extends PluginSettingTab {
    plugin: Export2WordPlugin;

    constructor(app: App, plugin: Export2WordPlugin) {
        super(app, plugin);
        this.plugin = plugin;
    }

    display(): void {
        const { containerEl } = this;
        containerEl.empty();

        // 标题（社区规范：设置面板标题不得包含插件名——左侧导航已显示）
        new Setting(containerEl).setName("导出选项").setHeading();
        containerEl.createEl("p", {
            text: "一键将 Obsidian Markdown 笔记导出为 Word (.docx) 文档",
            cls: "setting-item-description",
        });

        // Wikilink 处理方式
        new Setting(containerEl)
            .setName("Wikilink 处理方式")
            .setDesc("选择 [[内部链接]] 在导出时的处理方式")
            .addDropdown((dropdown) => {
                dropdown
                    .addOption("plain", "纯文本（显示笔记名）")
                    .addOption("zip", "附件打包（将关联笔记一起导出为 .zip）")
                    .setValue(this.plugin.settings.wikilinkMode)
                    .onChange(async (value) => {
                        this.plugin.settings.wikilinkMode = value as WikilinkMode;
                        await this.plugin.saveSettings();
                    });
            });

        // 递归深度（附件目录 + 附件打包共用）
        new Setting(containerEl)
            .setName("附件递归深度")
            .setDesc("导出「附件目录」与「附件打包」时，link 文档继续展开几层。0 = 不递归（仅导原文档，不展开任何 link）；1 层 = 仅直接 link；2 层 = 含其 link 的文档（如示例）；3 层 = 再深一层")
            .addDropdown((dropdown) => {
                dropdown
                    .addOption("0", "不递归（仅原文档）")
                    .addOption("1", "仅直接 link（1 层）")
                    .addOption("2", "包含间接 link（2 层）")
                    .addOption("3", "包含三层间接 link（3 层）")
                    .setValue(String(this.plugin.settings.recursionDepth))
                    .onChange(async (value) => {
                        this.plugin.settings.recursionDepth = parseInt(value, 10);
                        await this.plugin.saveSettings();
                    });
            })
            .setClass("export2word-recursion-setting");

        // 附件目录设置
        new Setting(containerEl).setName("附件目录").setHeading();

        new Setting(containerEl)
            .setName("导出时生成附件目录")
            .setDesc("开启后，普通导出会：①在文末（或单独文档）列出本笔记 link 的关联文档层级编号（附件1 / 附件1.1）；②把这些关联文档实际导出为独立的 .docx 文件，归入「<笔记名>附件」子文件夹（如 A附件/甲.docx）")
            .addToggle((toggle) => {
                toggle
                    .setValue(this.plugin.settings.attachmentIndexEnabled)
                    .onChange(async (value) => {
                        this.plugin.settings.attachmentIndexEnabled = value;
                        await this.plugin.saveSettings();
                        this.display();
                    });
            });

        new Setting(containerEl)
            .setName("附件目录位置")
            .setDesc("末尾：把附件目录追加到主文档末尾。单独文档：额外生成一个「<笔记名>-附件目录.docx」，与主文档同位置")
            .addDropdown((dropdown) => {
                dropdown
                    .addOption("end", "追加到主文档末尾")
                    .addOption("separate", "单独生成一个附件目录文档")
                    .setValue(this.plugin.settings.attachmentIndexLocation)
                    .onChange(async (value) => {
                        this.plugin.settings.attachmentIndexLocation = value as "end" | "separate";
                        await this.plugin.saveSettings();
                    });
            });

        // 导出路径
        new Setting(containerEl)
            .setName("默认导出路径")
            .setDesc("单文件与批量导出共用的默认保存位置（批量导出会在该位置下再建以文件夹名命名的子目录）")
            .addDropdown((dropdown) => {
                dropdown
                    .addOption("_exports", "vault 根目录的 _exports 文件夹")
                    .addOption("$sourceDir$", "源文件所在目录")
                    .addOption("$custom$", "自定义路径（在下方输入）")
                    .setValue(
                        this.plugin.settings.exportPath === "$sourceDir$" ? "$sourceDir$" :
                        this.plugin.settings.exportPath ? "$custom$" : "_exports"
                    )
                    .onChange(async (value) => {
                        if (value === "_exports") {
                            this.plugin.settings.exportPath = "";
                        } else if (value === "$sourceDir$") {
                            this.plugin.settings.exportPath = "$sourceDir$";
                        }
                        await this.plugin.saveSettings();
                        this.display();
                    });
            });

        // 自定义路径输入（仅在选自定义时显示）
        if (!this.plugin.settings.exportPath || this.plugin.settings.exportPath === "$sourceDir$") {
            // 隐藏自定义输入
        } else {
            new Setting(containerEl)
                .setName("自定义路径")
                .setDesc("相对于 vault 根目录的路径，如：MyExports")
                .addText((text) => {
                    text
                        .setPlaceholder("如：MyExports")
                        .setValue(this.plugin.settings.exportPath)
                        .onChange(async (value) => {
                            this.plugin.settings.exportPath = value.trim();
                            await this.plugin.saveSettings();
                        });
                });
        }

        // 批量导出设置
        new Setting(containerEl).setName("批量导出（文件夹）").setHeading();

        new Setting(containerEl)
            .setName("批量导出打包成一个 zip")
            .setDesc(
                "关闭（默认）：每个 .md 生成独立 .docx，按原目录层级写入导出路径。开启：整个文件夹所有 .md 转换后汇总压成「一个」.zip（内含全部 .docx，保持原目录层级）",
            )
            .addToggle((toggle) => {
                toggle
                    .setValue(this.plugin.settings.batchZipMode)
                    .onChange(async (value) => {
                        this.plugin.settings.batchZipMode = value;
                        await this.plugin.saveSettings();
                    });
            });

        // 快捷键说明（社区规范不允许默认热键，请用户在 Obsidian 快捷键设置中自行绑定）
        new Setting(containerEl).setName("快捷键").setHeading();

        new Setting(containerEl)
            .setName("绑定导出快捷键")
            .setDesc("本插件不设置默认热键。如需快捷键，请在 Obsidian「设置 → 快捷键」中搜索「导出为 Word」自行绑定。")
            .setDisabled(true);

        // 版本信息
        containerEl.createDiv({
            cls: "export2word-footer",
        }).createEl("p", {
            text: `MD to Word v${this.plugin.manifest.version} · MIT License`,
            cls: "setting-item-description",
        });
    }
}
