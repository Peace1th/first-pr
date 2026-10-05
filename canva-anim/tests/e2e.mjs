// 使い方: PLAYWRIGHT_DIR=<playwright の場所(省略可)> node tests/e2e.mjs <sample.pptx> <出力.pptx> [スクショ接頭辞]
import { createRequire } from 'node:module';
const { chromium } = createRequire(import.meta.url)(process.env.PLAYWRIGHT_DIR || 'playwright');
import path from 'node:path';
import { pathToFileURL } from 'node:url';

const [, , sample, out, shot = ''] = process.argv;
const url = pathToFileURL(path.resolve(import.meta.dirname, '..', 'index.html')).href;
const browser = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium-1194/chrome-linux/chrome' }).catch(() => chromium.launch());
const page = await browser.newPage({ viewport: { width: 1400, height: 1000 }, acceptDownloads: true });
const errors = [];
page.on('pageerror', (e) => errors.push('pageerror: ' + e.message));
page.on('console', (m) => { if (m.type() === 'error') errors.push('console: ' + m.text()); });

await page.goto(url);
await page.setInputFiles('#file', sample);
await page.waitForSelector('#app:not([hidden])');
const shapes = await page.locator('#stage .shp').count();
console.log('shapes on slide 1:', shapes, '| slides:', await page.locator('#slideList li').count());

// 1枚目: 一括でフェード(最初だけクリック、残りは自動)
await page.selectOption('#bulkEffect', 'fade');
await page.selectOption('#bulkTrigger', 'click-after');
await page.click('#bulkSlide');
console.log('order items after bulk:', await page.locator('#orderList li button.pick').count());

// 画像を選んで「ズーム + 前と同時」の2つ目のアニメを追加
await page.locator('#elList button', { hasText: '画像' }).click();
await page.click('text=＋ アニメーションを追加');
const cards = page.locator('.anim-card');
await cards.nth(0).locator('select').nth(0).selectOption('fly');
await cards.nth(0).locator('select').nth(1).selectOption('bottom');
await cards.nth(0).locator('select').nth(2).selectOption('with');
await cards.nth(1).locator('select').nth(0).selectOption('fadeOut');
await page.selectOption('#transition', 'push');

// 2枚目: ワイプ(下から)、すべて自動
await page.locator('#slideList button').nth(1).click();
await page.selectOption('#bulkEffect', 'wipe');
await page.selectOption('#bulkDir', 'bottom');
await page.selectOption('#bulkTrigger', 'auto');
await page.click('#bulkSlide');
await page.selectOption('#transition', 'fade');

// プレビュー再生
await page.click('#play');
await page.waitForTimeout(700);
if (shot) await page.screenshot({ path: shot + '_play.png' });
await page.click('#stop');
await page.locator('#slideList button').nth(0).click();
await page.locator('#stage .shp').nth(3).click();
if (shot) await page.screenshot({ path: shot + '_edit.png', fullPage: true });

const [dl] = await Promise.all([page.waitForEvent('download'), page.click('#download')]);
await dl.saveAs(out);
console.log('saved', dl.suggestedFilename());
console.log('errors:', errors.length ? errors : 'none');
await browser.close();
process.exit(errors.length ? 1 : 0);
