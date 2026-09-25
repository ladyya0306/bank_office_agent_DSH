'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const vm = require('node:vm');
const { MARKER, START, patchQuestionDetail } = require('../question-detail.cjs');

const renderSite = `question.detail !== void 0 && (0, react_jsx_runtime.jsx)("div", {
							className: QuestionComposer_module_css_default.detail,
							children: (0, react_jsx_runtime.jsx)(_deepseek_ai_dsh_client_ui_primitives.MarkdownText, {
								text: question.detail,
								labels: markdownLabels
							})
						}), (0, react_jsx_runtime.jsxs)("div", {
							className: QuestionComposer_module_css_default.options,
							children: []
						})`;

test('wraps only marked tool location details in collapsed native details', () => {
	assert.ok(renderSite.includes(START));
	const patched = patchQuestionDetail(`prefix\n${renderSite}\nsuffix`);
	assert.match(patched, /question\.detail\.startsWith\("<!--dsh-fill-locations:v1-->"\)/);
	assert.match(patched, /jsx\)\("summary", \{ children: "查看填写位置" \}\)/);
	assert.match(patched, /data-dsh-fill-locations/);
	assert.match(patched, /text: question\.detail\.slice\(/);
	assert.equal(patchQuestionDetail(patched), patched, 'patch is idempotent');
});

test('keeps unmarked detail in the original markdown renderer branch', () => {
	const patched = patchQuestionDetail(renderSite);
	assert.match(patched, /: \(0, react_jsx_runtime\.jsx\)\("div"/);
	assert.match(patched, /text: question\.detail,/);
	assert.ok(MARKER.startsWith('<!--'));
});

test('fails clearly when the installed DSH anchor changes', () => {
	assert.throws(() => patchQuestionDetail('changed renderer'), /anchor did not match exactly once/);
});

test('renders collapsed locations and keeps ordinary supporting details visible', () => {
  const render = (detail) => vm.runInNewContext('[' + patchQuestionDetail(renderSite) + ']', {
    question: { detail }, markdownLabels: {},
    react_jsx_runtime: { jsx: (type, props) => ({ type, props }), jsxs: (type, props) => ({ type, props }) },
    QuestionComposer_module_css_default: { detail: 'detail', options: 'options' },
    _deepseek_ai_dsh_client_ui_primitives: { MarkdownText: 'markdown' },
  });
  const locations = render(MARKER + '\n- synthetic.docx: B1')[0];
  assert.equal(locations.type, 'details');
  assert.equal(locations.props.open, undefined);
  assert.equal(locations.props.children[0].type, 'summary');
  assert.equal(locations.props.children[1].props.text, '- synthetic.docx: B1');
  const ordinary = render('ordinary supporting detail')[0];
  assert.equal(ordinary.type, 'div');
  assert.equal(ordinary.props.children.props.text, 'ordinary supporting detail');
  assert.equal(render(undefined)[0], false);
});
