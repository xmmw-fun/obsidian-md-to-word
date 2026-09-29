// 替代 npm 包 `immediate`（docx → jszip → lie 的依赖链）。
// 原包包含 `document.createElement("script")` 特性检测代码，
// 会被 Obsidian 社区插件审查的静态扫描判为「运行时创建 script 元素」（Error 级）。
// 本 shim 用 Promise 微任务提供等价的「尽快异步执行」语义。
// （上架审查整改，2026-09-29，署名 XMMW）
module.exports = function immediate(task) {
	Promise.resolve().then(task);
};
