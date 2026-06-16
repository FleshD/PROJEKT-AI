import { createServer } from 'node:http';
import { readFile } from 'node:fs/promises';
import { extname, join, normalize } from 'node:path';
import { fileURLToPath } from 'node:url';
import { chromium } from 'playwright';

const root = normalize(fileURLToPath(new URL('..', import.meta.url)));
const port = Number(process.env.PORT) || 8000;
const screenshotPath = join(root, 'snake-game', 'screenshot.png');

const contentTypes = {
  '.html': 'text/html; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.png': 'image/png',
};

const server = createServer(async (request, response) => {
  const requestUrl = new URL(request.url ?? '/', `http://127.0.0.1:${port}`);
  const pathname = requestUrl.pathname === '/snake-game/' ? '/snake-game/index.html' : requestUrl.pathname;
  const filePath = normalize(join(root, pathname));

  if (!filePath.startsWith(root)) {
    response.writeHead(403);
    response.end('Forbidden');
    return;
  }

  try {
    const body = await readFile(filePath);
    response.writeHead(200, { 'content-type': contentTypes[extname(filePath)] ?? 'application/octet-stream' });
    response.end(body);
  } catch {
    response.writeHead(404);
    response.end('Not found');
  }
});

await new Promise((resolve) => server.listen(port, '127.0.0.1', resolve));

try {
  const browser = await chromium.launch();
  const page = await browser.newPage({ viewport: { width: 1000, height: 900 } });
  await page.goto(`http://127.0.0.1:${port}/snake-game/`);
  await page.getByRole('button', { name: 'Start / restart' }).click();
  await page.keyboard.press('ArrowDown');
  await page.waitForTimeout(250);
  await page.screenshot({ path: screenshotPath, fullPage: true });
  await browser.close();
  console.log(`Playwright verification passed. Screenshot saved to ${screenshotPath}`);
} finally {
  server.close();
}
