/**
 * Playwright 截图脚本 — 由 Python 后端通过子进程调用。
 *
 * 用法: node screenshot.js <questionId> <cookie> <token> <domain> [selector]
 *   domain: 莫顿环境域名（如 https://tms.testing.xueqiulearning.com），
 *           由后端按当前运行环境（测试/正式）传入，截图与生成保持一致。
 * 输出: PNG 二进制流写入 stdout，错误信息写入 stderr
 *
 * 优先使用系统 Chrome，避免下载 Chromium 浏览器（约 300MB）。
 * 如果找不到系统 Chrome，回退到 Playwright 自带的 Chromium。
 */

const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');
const { URL } = require('url');

// 系统 Chrome 可能的安装路径（使用正斜杠，兼容 Node.js）
const CHROME_PATHS = [
  'C:/Program Files/Google/Chrome/Application/chrome.exe',
  'C:/Program Files (x86)/Google/Chrome/Application/chrome.exe',
  path.join((process.env.LOCALAPPDATA || '').replace(/\\/g, '/'), 'Google/Chrome/Application/chrome.exe'),
];

function findSystemChrome() {
  for (const p of CHROME_PATHS) {
    if (p && fs.existsSync(p)) return p;
  }
  return null;
}

async function main() {
  const questionId = process.argv[2];
  const cookie = process.argv[3];
  const token = process.argv[4];
  const domain = process.argv[5] || 'https://tms.testing.xueqiulearning.com'; // 默认测试环境
  const selector = process.argv[6] || '';

  if (!questionId || !cookie || !token) {
    process.stderr.write('Usage: node screenshot.js <questionId> <cookie> <token> <domain> [selector]\n');
    process.exit(1);
  }

  // 跟随当前运行环境：用后端传入的 domain 拼接 widget 地址
  let widgetHost;
  try {
    widgetHost = new URL(domain).host;
  } catch (_) {
    widgetHost = 'tms.testing.xueqiulearning.com';
  }
  const WIDGET_URL = `https://${widgetHost}/fe_morton/morton/widget?id=${questionId}`;
  process.stderr.write(`[INFO] 截图环境域名: ${widgetHost}\n`);

  let browser;
  try {
    // 查找系统 Chrome
    const chromePath = findSystemChrome();

    if (chromePath) {
      process.stderr.write(`[INFO] Using system Chrome: ${chromePath}\n`);
      browser = await chromium.launch({
        executablePath: chromePath,
        headless: true,
        args: [
          '--no-sandbox',
          '--disable-gpu',
          '--disable-dev-shm-usage',
          '--disable-extensions',
          '--disable-background-networking',
          '--disable-sync',
          '--no-first-run',
        ],
      });
    } else {
      process.stderr.write('[INFO] System Chrome not found, using Playwright Chromium\n');
      browser = await chromium.launch({
        headless: true,
        args: [
          '--no-sandbox',
          '--disable-gpu',
          '--disable-dev-shm-usage',
          '--disable-extensions',
        ],
      });
    }

    const context = await browser.newContext({
      viewport: { width: 1334, height: 750 },
      deviceScaleFactor: 1,
    });

    // 注入 Cookie（domain 跟随当前环境，与 widget 地址一致）
    const cookies = [];
    for (const pair of cookie.split(';')) {
      const trimmed = pair.trim();
      if (trimmed.includes('=')) {
        const [k, ...vParts] = trimmed.split('=');
        cookies.push({
          name: k.trim(),
          value: vParts.join('=').trim(),
          domain: widgetHost,
          path: '/',
        });
      }
    }
    if (cookies.length > 0) {
      await context.addCookies(cookies);
    }

    // 设置 Authorization 头
    await context.setExtraHTTPHeaders({
      'Authorization': token,
    });

    const page = await context.newPage();

    // 导航到莫顿题目页面（用 domcontentloaded 更可靠，SPA 可能一直有网络请求）
    await page.goto(WIDGET_URL, { waitUntil: 'domcontentloaded', timeout: 30000 });

    // 等待画布渲染（SPA 需要等待 React/Vue 渲染完成）
    try {
      await page.waitForSelector(
        'canvas, .question-container, .widget-container, .morton-canvas, [class*="canvas"], [class*="widget"]',
        { timeout: 10000 }
      );
    } catch (_) {
      // 找不到特定元素时，额外等待
      await page.waitForTimeout(3000);
    }

    // 额外等待确保渲染完成
    await page.waitForTimeout(1500);

    // 截图
    let buffer;
    if (selector) {
      const element = await page.$(selector);
      if (element) {
        buffer = await element.screenshot({ type: 'png' });
      } else {
        buffer = await page.screenshot({ type: 'png', fullPage: false });
      }
    } else {
      buffer = await page.screenshot({ type: 'png', fullPage: false });
    }

    // 写入 stdout（二进制）
    process.stdout.write(buffer);

    await browser.close();
  } catch (err) {
    process.stderr.write(`SCREENSHOT_ERROR: ${err.message}\n`);
    if (browser) {
      try { await browser.close(); } catch (_) {}
    }
    process.exit(2);
  }
}

main();
