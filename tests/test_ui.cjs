// State-flow tests for the task pane, using Node's built-in runner.
const { test } = require('node:test');
const assert = require('node:assert/strict');
const vm = require('node:vm');
const fs = require('node:fs');
const path = require('node:path');

function pane() {
  const elements = new Map();
  const node = () => ({ disabled: false, hidden: false, textContent: '', files: [], children: [],
    events: {}, addEventListener(event, fn) { this.events[event] = fn; },
    replaceChildren() { this.children = []; }, append(item) { this.children.push(item); },
    click() {}, remove() {} });
  const get = id => { if (!elements.has(id)) elements.set(id, node()); return elements.get(id); };
  let reply = { count: 3, images: [{ sheet_name: 'Sheet1', cell: 'A1' }], converted: 3, filename: 'out.xlsx', workbookBase64: 'eA==' };
  let okay = true;
  const context = { document: { querySelector: get, createElement: node, body: node() },
    fetch: async () => ({ ok: okay, json: async () => reply }),
    Uint8Array, Blob, URL, setTimeout: fn => fn(), btoa, atob };
  vm.runInNewContext(fs.readFileSync(path.join(__dirname, '../plugins/shared/app.js'), 'utf8'), context);
  return { get, reply(value, ok = true) { reply = value; okay = ok; },
    choose(name = 'sample.xlsx') { get('#workbook').files = [{ name, size: 1, arrayBuffer: async () => new Uint8Array([1]).buffer }]; get('#workbook').events.change(); } };
}

test('selection, detection, repair, download and changed input invalidate result', async () => {
  const ui = pane();
  assert.equal(ui.get('#save').disabled, true);
  assert.equal(ui.get('#repair').disabled, true);
  ui.choose();
  assert.equal(ui.get('#repair').disabled, false);
  await ui.get('#inspect').events.click();
  assert.match(ui.get('#status').textContent, /检测到 3/);
  assert.equal(ui.get('#save').disabled, true);
  await ui.get('#repair').events.click();
  assert.equal(ui.get('#save').disabled, false);
  ui.get('#save').events.click();
  assert.match(ui.get('#status').textContent, /已发起下载/);
  ui.choose('another.xlsx');
  assert.equal(ui.get('#save').disabled, true);
});

test('invalid input and failure cannot save stale output', async () => {
  const ui = pane();
  ui.choose('bad.xlsm');
  assert.equal(ui.get('#repair').disabled, true);
  ui.choose();
  await ui.get('#repair').events.click();
  ui.reply({ error: 'Missing media' }, false);
  await ui.get('#repair').events.click();
  assert.equal(ui.get('#save').disabled, true);
  assert.equal(ui.get('#repair').disabled, false);
  assert.match(ui.get('#status').textContent, /Missing media/);
});
