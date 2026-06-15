/**
 * Playwright 截图脚本 — 由 Python 后端通过子进程调用。
 *
 * 用法: node screenshot.js <questionId> <cookie> <token> [selector]
 * 输出: PNG 二进制流写入 stdout，错误信息写入 stderr
 *
 * 优先使用系统 Chrome，避免下载 Chromium 浏览器（约 300MB）。
 * 如果找不到系统 Chrome，回退到 Playwright 自带的 Chromium。
 */

const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');

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
  const selector = process.argv[5] || '';

  if (!questionId || !cookie || !token) {
    process.stderr.write('Usage: node screenshot.js <questionId> <cookie> <token> [selector]\n');
    process.exit(1);
  }

  const WIDGET_URL = `https://tms-mx.xueqiulearning.com/fe_morton/morton/widget?id=${questionId}`;

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

    // 注入 Cookie
    const cookies = [];
    for (const pair of cookie.split(';')) {
      const trimmed = pair.trim();
      if (trimmed.includes('=')) {
        const [k, ...vParts] = trimmed.split('=');
        cookies.push({
          name: k.trim(),
          value: vParts.join('=').trim(),
          domain: 'tms-mx.xueqiulearning.com',
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
