/**
 * export2word - 插件设置
 * v1.0.0
 */

export type WikilinkMode = "plain" | "hyperlink" | "zip";

export interface Export2WordSettings {
    /** Wikilink 处理方式 */
    wikilinkMode: WikilinkMode;
    /** 附件目录 / 附件打包递归深度：0 = 不递归（仅原文档）；1/2/3 = 展开对应层数 */
    recursionDepth: number;
    /** 默认导出路径（单文件与批量共用；空 = vault 根的 _exports；"$sourceDir$" = 源目录；其他 = 自定义路径） */
    exportPath: string;
    /** 批量导出时，是否将所有 .docx 汇总压成一个 .zip（而非每个 .md 独立 .docx）；默认 false */
    batchZipMode: boolean;
    /** 是否生成附件目录：普通导出时，按递归深度列出关联文档的层级清单（附件1 / 附件1.1） */
    attachmentIndexEnabled: boolean;
    /** 附件目录位置：end = 追加到主文档末尾；separate = 单独生成一个「附件目录.docx」 */
    attachmentIndexLocation: "end" | "separate";
}

export const DEFAULT_SETTINGS: Export2WordSettings = {
    wikilinkMode: "plain",
    recursionDepth: 1,
    exportPath: "",
    batchZipMode: false,
    attachmentIndexEnabled: false,
    attachmentIndexLocation: "end",
};
