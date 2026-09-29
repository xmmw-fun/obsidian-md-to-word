/**
 * markdown-it 插件：解析 Obsidian Wikilink 语法
 * [[笔记名]] 或 [[笔记名|别名]] 或 [[笔记名#锚点]]
 */
import type MarkdownIt from "markdown-it";

export function wikilinkPlugin(md: MarkdownIt): void {
    // 在行内解析阶段插入
    md.inline.ruler.before("link", "wikilink", (state, silent) => {
        let pos = state.pos;
        const src = state.src;

        // 检测两种形态：
        //   [[...]]   → 笔记链接
        //   ![[...]]  → 嵌入（嵌入图 / 嵌入笔记），需吃掉开头的 !，避免被当作文本
        let isEmbed = false;
        if (
            src.charCodeAt(pos) === 0x21 /* ! */ &&
            src.charCodeAt(pos + 1) === 0x5b /* [ */ &&
            src.charCodeAt(pos + 2) === 0x5b /* [ */
        ) {
            isEmbed = true;
            pos += 1; // 跳过 !，后续与 [[ 逻辑一致
        } else if (
            src.charCodeAt(pos) === 0x5b /* [ */ &&
            src.charCodeAt(pos + 1) === 0x5b /* [ */
        ) {
            // 普通 [[ 链接
        } else {
            return false;
        }

        // 查找闭合 ]]
        const endPos = src.indexOf("]]", pos + 2);
        if (endPos === -1) return false;

        const raw = src.slice(pos + 2, endPos);
        if (!raw.trim()) return false;

        if (!silent) {
            const token = state.push("wikilink", "span", 0);
            token.content = raw;
            token.markup = "wikilink";
            if (isEmbed) token.attrSet("data-embed", "true");

            // 解析 target | alias 或 target#anchor
            const pipeIdx = raw.indexOf("|");
            const hashIdx = raw.indexOf("#");

            if (pipeIdx !== -1) {
                token.attrSet("data-target", raw.slice(0, pipeIdx).trim());
                token.attrSet("data-alias", raw.slice(pipeIdx + 1).trim());
            } else if (hashIdx !== -1) {
                token.attrSet("data-target", raw.slice(0, hashIdx).trim());
                token.attrSet("data-anchor", raw.slice(hashIdx + 1).trim());
            } else {
                token.attrSet("data-target", raw.trim());
            }
        }

        state.pos = endPos + 2;
        return true;
    });

    // 注册 wikilink 的渲染规则（在 converter 中使用，这里提供一个默认占位）
    md.renderer.rules.wikilink = (tokens, idx) => {
        const token = tokens[idx];
        const target = token.attrGet("data-target") || token.content;
        const alias = token.attrGet("data-alias");
        return alias || target;
    };
}
