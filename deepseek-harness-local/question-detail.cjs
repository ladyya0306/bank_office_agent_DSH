'use strict';

const MARKER = '<!--dsh-fill-locations:v1-->';
const PATCH_MARKER = '"data-dsh-fill-locations": "v1"';

const START = 'question.detail !== void 0 && (0, react_jsx_runtime.jsx)("div", {';
const LABELS = 'labels: markdownLabels';
const NEXT_DETAIL = '}), (0, react_jsx_runtime.jsxs)("div", {';

const REPLACEMENT = `question.detail !== void 0 && (question.detail.startsWith("${MARKER}")
							? (0, react_jsx_runtime.jsxs)("details", {
								"data-dsh-fill-locations": "v1",
								className: QuestionComposer_module_css_default.detail,
								children: [(0, react_jsx_runtime.jsx)("summary", { children: "查看填写位置" }),
									(0, react_jsx_runtime.jsx)(_deepseek_ai_dsh_client_ui_primitives.MarkdownText, {
										text: question.detail.slice("${MARKER}".length).trim(),
										labels: markdownLabels
									})]
							})
							: (0, react_jsx_runtime.jsx)("div", {
								className: QuestionComposer_module_css_default.detail,
								children: (0, react_jsx_runtime.jsx)(_deepseek_ai_dsh_client_ui_primitives.MarkdownText, {
									text: question.detail,
									labels: markdownLabels
								})
							}))`;

/** Idempotently patch the known DSH question-detail render site. */
function patchQuestionDetail(source) {
	if (typeof source !== 'string') throw new TypeError('client source must be a string');
	if (source.includes(PATCH_MARKER)) return source;
	const first = source.indexOf(START);
	if (first < 0 || source.indexOf(START, first + START.length) >= 0) {
		throw new Error('question detail start anchor did not match exactly once');
	}
	const labelsAt = source.indexOf(LABELS, first + START.length);
	const closeAt = labelsAt < 0 ? -1 : source.indexOf(NEXT_DETAIL, labelsAt + LABELS.length);
	if (labelsAt < 0 || closeAt < 0) {
		throw new Error('question detail body or trailing anchor did not match');
	}
	const end = closeAt + 2;
	return source.slice(0, first) + REPLACEMENT + source.slice(end);
}

module.exports = { MARKER, START, patchQuestionDetail };
