const fs = require('node:fs');
const path = require('node:path');
const { JSDOM, requestInterceptor, VirtualConsole } = require('jsdom');
const root = path.resolve(__dirname, '..');
async function openSite(transform = text => text, beforeParse = () => {}) {
  const errors = [];
  const resources = { interceptors: [requestInterceptor(request => {
      const u = new URL(request.url);
      let file;
      if (u.hostname === 'site.test') file = path.join(root, decodeURIComponent(u.pathname));
      else if (u.pathname.includes('jszip')) file = require.resolve('jszip/dist/jszip.min.js');
      else if (u.pathname.includes('Sortable')) file = require.resolve('sortablejs/Sortable.min.js');
      if (!file || !fs.existsSync(file)) return new Response('');
      let content = fs.readFileSync(file);
      if (file === path.join(root, 'script.js')) content = Buffer.from(transform(content.toString()));
      return new Response(content, { headers: { 'Content-Type': file.endsWith('.css') ? 'text/css' : 'application/javascript' } });
  })] };
  const console = new VirtualConsole();
  console.on('jsdomError', error => errors.push(error.message));
  const dom = new JSDOM(fs.readFileSync(path.join(root, 'index.html'), 'utf8'), {
    url: 'https://site.test/index.html', runScripts: 'dangerously', resources,
    pretendToBeVisual: true, virtualConsole: console,
    beforeParse(w) { w.setImmediate = setImmediate; w.clearImmediate = clearImmediate; w.alert = () => {}; w.confirm = () => true; beforeParse(w); }
  });
  await new Promise(resolve => dom.window.addEventListener('load', resolve, { once: true }));
  return { dom, w: dom.window, errors };
}
async function importFile(w, selector, content, name = 'project.json') {
  const input = w.document.querySelector(selector);
  const file = new w.File([content], name, { type: 'application/octet-stream' });
  Object.defineProperty(input, 'files', { configurable: true, value: [file] });
  input.dispatchEvent(new w.Event('change'));
  await new Promise(resolve => setTimeout(resolve, 80));
}
module.exports = { openSite, importFile };
