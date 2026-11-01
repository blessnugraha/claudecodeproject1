// Bundles src/ into one self-contained HTML file: dist/idx-xbrl-workbench.html
// Every script and the stylesheet are inlined, so the page works offline from a double-click.
import { readFileSync, writeFileSync, mkdirSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = dirname(fileURLToPath(import.meta.url));
const srcDir = resolve(root, 'src');
const out = resolve(root, 'dist', 'idx-xbrl-workbench.html');

const read = (rel) => readFileSync(resolve(srcDir, rel), 'utf8');
// keep inline code from closing its own <script> element early
const safeJs = (js) => js.replace(/<\/script/gi, '<\\/script').replace(/<!--/g, '<\\!--');

let html = read('index.html');
html = html.replace(/<link rel="stylesheet" href="styles\.css">/, () => `<style>\n${read('styles.css')}</style>`);
html = html.replace(/<script src="([^"]+)"><\/script>/g, (_, src) => `<script>\n${safeJs(read(src))}\n</script>`);

if (/<script src=|href="styles\.css"/.test(html)) throw new Error('A local resource was not inlined');
mkdirSync(dirname(out), { recursive: true });
writeFileSync(out, html);
console.log(`wrote ${out} (${(html.length / 1024).toFixed(0)} KB)`);
