'use strict';
// Code.gs を Node の vm でロードし、モックしたGASグローバルの上で実行できる状態にして返す。
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { buildSandbox } = require('./mock-gas');

const CODE_PATH = path.join(__dirname, '..', 'src', 'Code.gs');

function loadCode(mockOpts) {
  const src = fs.readFileSync(CODE_PATH, 'utf8');
  const built = buildSandbox(mockOpts);
  const context = vm.createContext(Object.assign({}, built.sandbox));
  vm.runInContext(src, context, { filename: 'Code.gs' });
  return Object.assign({ context }, built);
}

module.exports = { loadCode, CODE_PATH };
