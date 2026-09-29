// 替代 npm 包 `setimmediate`（jszip 的依赖）。
// 原包特性检测里含 `document.createElement("script")`，
// 会被 Obsidian 社区插件审查的静态扫描判为「运行时创建 script 元素」（Error 级）。
// Obsidian 桌面端基于 Electron，全局已有 setImmediate；这里仅在缺失时用
// Promise 微任务兜底，语义与 setImmediate 一致（尽快异步执行）。
// （上架审查整改，2026-09-29，署名 XMMW）
(function (global) {
	if (typeof global.setImmediate !== "function") {
		global.setImmediate = function (fn) {
			var args = Array.prototype.slice.call(arguments, 1);
			Promise.resolve().then(function () { fn.apply(null, args); });
		};
		global.clearImmediate = function () {};
	}
})(typeof window !== "undefined" ? window : globalThis);
