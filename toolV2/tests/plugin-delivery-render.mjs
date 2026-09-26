import assert from 'node:assert/strict';
import { renderOfficeOutput } from '../dsh-plugin/render.mjs';

const files = Array.from({ length: 16 }, (_, index) => `C:\\work\\output\\result-${index + 1}.docx`);
const response = {
  ok: true,
  delivery: {
    versions: [{ path: files[0], version: 1, template_path: 'target/source-template.docx' }],
    attachments: [...files],
    attachment_groups: [files.slice(0, 8), files.slice(8)],
    max_attachments_per_group: 8,
  },
};
const original = structuredClone(response);

const rendered = JSON.parse(renderOfficeOutput(response));

assert.equal(rendered.delivery.present_calls.length, 2);
assert.deepEqual(rendered.delivery.present_calls.map((call) => call.files.length), [8, 8]);
assert.deepEqual(rendered.delivery.present_calls[0].files[0], {
  path: files[0], description: 'output / source-template.docx',
});
assert.equal(Object.hasOwn(rendered.delivery, 'attachments'), false);
assert.equal(Object.hasOwn(rendered.delivery, 'attachment_groups'), false);
assert.deepEqual(response, original);

console.log('delivery render grouping passed');
