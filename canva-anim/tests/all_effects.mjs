// 全効果・全方向の p:timing を既存PPTXの1枚目に埋め込んだファイルを作る(LibreOffice 等で開いて確認する用)。
// 使い方: node tests/all_effects.mjs <sample.pptx> <出力.pptx>
import { createRequire } from 'node:module';
import fs from 'node:fs';
const require = createRequire(import.meta.url);
const JSZip = require('../vendor/jszip.min.js');
const A = require('../pptx-anim.js');

const [, , src, out] = process.argv;
const zip = await JSZip.loadAsync(fs.readFileSync(src));
const path = 'ppt/slides/slide1.xml';
let xml = await zip.file(path).async('string');

const anims = [];
for (const [effect, def] of Object.entries(A.EFFECTS)) {
  const dirs = def.dirs ? ['left', 'right', 'top', 'bottom'] : [null];
  for (const dir of dirs) anims.push({ spid: 4, effect, dir, trigger: anims.length % 3 === 0 ? 'click' : 'after', dur: 400, delay: 100 });
}
const timing = A.buildTiming(anims, { 4: 'sp' }).replace(/ xmlns:p="[^"]+"/, '');
xml = xml.replace('</p:sld>', timing + '</p:sld>');
zip.file(path, xml, { createFolders: false });
fs.writeFileSync(out, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }));
console.log('effects written:', anims.length);
