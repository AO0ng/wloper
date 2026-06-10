const { chromium } = require("playwright");
const XLSX = require("xlsx");
const path = require("path");
const fs = require("fs");

// ==================== 配置区 ====================
const CONFIG = {
  referenceDate: "2026-06-09",
  stations: ["梓坊", "先锋", "土桥", "渣津", "瑞昌", "萍乡", "虬津"],
  gotoTimeout: 60000,
  loginUrl: "http://weixin.jxsswj.cn/jxhydp-app/#login",
  dataUrl: "http://weixin.jxsswj.cn/jxhydp-app/#hydataview",
  username: "夏雄",
  password: "Jxsw_123",
};

const OUTPUT_DIR = path.join(__dirname, "output");

// ==================== 日期工具 ====================
function getWeekRange(refDateStr) {
  const ref = new Date(refDateStr + "T00:00:00");
  const dayOfWeek = ref.getDay();
  const daysFromMonday = dayOfWeek === 0 ? 6 : dayOfWeek - 1;
  const thisMonday = new Date(ref);
  thisMonday.setDate(ref.getDate() - daysFromMonday);
  const lastMonday = new Date(thisMonday);
  lastMonday.setDate(thisMonday.getDate() - 7);
  const lastSunday = new Date(lastMonday);
  lastSunday.setDate(lastMonday.getDate() + 6);

  const fmt = (d) => {
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, "0");
    const day = String(d.getDate()).padStart(2, "0");
    return `${y}-${m}-${day}`;
  };

  const dates = [];
  for (let i = 0; i < 7; i++) {
    const d = new Date(lastMonday);
    d.setDate(lastMonday.getDate() + i);
    dates.push(fmt(d));
  }

  return { monday: fmt(lastMonday), sunday: fmt(lastSunday), dates };
}

// ==================== 登录 ====================
async function doLogin(page) {
  await page.goto(CONFIG.loginUrl, { waitUntil: "domcontentloaded", timeout: CONFIG.gotoTimeout });
  await page.waitForTimeout(5000);

  // 登录页面是 ExtJS 构建的，输入框可能是 ExtJS 组件
  // 尝试多种方式找到用户名和密码输入框
  let usernameInput = page.locator('input[name="username"]');
  if (await usernameInput.count() === 0) {
    // 找可见的、可编辑的、非密码的 input
    usernameInput = page.locator('input:not([readonly]):not([type="password"]):visible').first();
  }

  let passwordInput = page.locator('input[name="password"]');
  if (await passwordInput.count() === 0) {
    passwordInput = page.locator('input[type="password"]:visible').first();
  }

  const uc = await usernameInput.count();
  const pc = await passwordInput.count();
  console.log(`  找到用户名框: ${uc}, 密码框: ${pc}`);

  if (uc === 0 || pc === 0) {
    console.log("  ⚠ 未找到登录输入框，尝试截图保存...");
    await page.screenshot({ path: path.join(OUTPUT_DIR, "login-page.png"), fullPage: true });
    return false;
  }

  await usernameInput.fill(CONFIG.username);
  await passwordInput.fill(CONFIG.password);

  const loginBtn = page.locator('a:has-text("登 录"), button:has-text("登录"), button:has-text("登 录")');
  if (await loginBtn.count() > 0) {
    await loginBtn.first().click();
  } else {
    await passwordInput.press("Enter");
  }
  await page.waitForTimeout(5000);
  return true;
}

// ==================== 数据提取 ====================
async function extractStationData(page) {
  const result = await page.evaluate(() => {
    const results = [];
    // 找包含大量列（>=30）的 data view — 即有1日~31日列的
    const views = document.querySelectorAll("[data-boundview]");
    let targetView = null;
    for (const v of views) {
      const headers = v.querySelectorAll(".x-column-header-text-inner, .x-column-header-text");
      if (headers.length >= 30) {
        targetView = v;
        break;
      }
    }
    if (!targetView) {
      targetView = document.querySelector('[data-boundview="tableview-1275"]');
    }
    if (!targetView) {
      return { results, error: "未找到数据 grid view" };
    }

    // 提取表头
    const headerEls = targetView.querySelectorAll(".x-column-header-text-inner, .x-column-header-text");
    const headers = [];
    for (const h of headerEls) {
      headers.push((h.textContent || "").trim());
    }

    const stationNameIdx = headers.findIndex(h => h === "站名" || h.includes("站名"));
    const stationCodeIdx = headers.findIndex(h => h === "站码" || h.includes("站码"));
    const riverIdx = headers.findIndex(h => h === "河流" || h.includes("河"));
    const day1Idx = headers.findIndex(h => h === "1日" || h === "1");

    // 提取数据行
    const rows = targetView.querySelectorAll(".x-grid-item");
    for (const row of rows) {
      const cells = row.querySelectorAll(".x-grid-cell-inner");
      const texts = [];
      for (const cell of cells) {
        texts.push((cell.textContent || "").trim());
      }
      if (texts.length < 3) continue;

      const stationCode = stationCodeIdx >= 0 ? (texts[stationCodeIdx] || "") : (texts[2] || "");
      const stationName = stationNameIdx >= 0 ? (texts[stationNameIdx] || "") : (texts[3] || "");
      const riverName = riverIdx >= 0 ? (texts[riverIdx] || "") : (texts[4] || "");

      if (!/^\d+$/.test(stationCode)) continue;

      const dailyFlows = [];
      const startCol = day1Idx >= 0 ? day1Idx : 8;
      for (let i = startCol; i < texts.length; i++) {
        const val = parseFloat(texts[i]);
        if (!isNaN(val)) {
          dailyFlows.push({ day: i - startCol + 1, flow: val });
        }
      }

      results.push({ stationCode, stationName, riverName, dailyFlows });
    }

    return {
      results,
      debug: {
        viewId: targetView.getAttribute("data-boundview"),
        headerCount: headers.length,
        headers: headers.slice(0, 40),
        stationNameIdx, stationCodeIdx, riverIdx, day1Idx,
        rowCount: results.length,
        firstRow: results.length > 0 ? results[0] : null,
      }
    };
  });

  return result;
}

// ==================== 主流程 ====================
(async () => {
  const week = getWeekRange(CONFIG.referenceDate);
  console.log(`=== 水文周报抓取 ===`);
  console.log(`基准日期: ${CONFIG.referenceDate}`);
  console.log(`抓取周期: ${week.monday} ~ ${week.sunday}`);
  console.log(`目标站点: ${CONFIG.stations.join(", ")}`);
  console.log("");

  fs.mkdirSync(OUTPUT_DIR, { recursive: true });

  const browser = await chromium.launch({ headless: true, args: ["--no-sandbox"] });
  const context = await browser.newContext({ locale: "zh-CN", viewport: { width: 1920, height: 1080 } });
  const page = await context.newPage();

  try {
    // [1] 登录
    console.log("[1/4] 登录...");
    const loggedIn = await doLogin(page);
    if (!loggedIn) {
      console.log("  登录失败，请查看 output/login-page.png");
      return;
    }
    console.log("  ✓ 登录完成");

    // [2] 进入数据视图并点击菜单
    console.log("[2/4] 进入逐日平均流量统计...");
    await page.goto(CONFIG.dataUrl, { waitUntil: "domcontentloaded", timeout: CONFIG.gotoTimeout });
    await page.waitForTimeout(3000);

    await page.waitForSelector(".x-treelist-item-text", { timeout: 15000 }).catch(() => {});
    await page.waitForTimeout(2000);

    const treeItems = await page.$$(".x-treelist-item-text");
    let menuClicked = false;
    for (const item of treeItems) {
      const text = (await item.textContent()).trim();
      if (text.includes("逐日平均流量")) {
        await item.click();
        menuClicked = true;
        break;
      }
    }
    if (!menuClicked) {
      await page.locator('text="逐日平均流量统计"').first().click();
    }

    console.log("  等待数据加载...");
    await page.waitForTimeout(5000);
    try {
      await page.waitForSelector(".x-grid-item", { timeout: 15000 });
    } catch (_) {
      console.log("  ⚠ .x-grid-item 未出现");
    }
    await page.waitForTimeout(3000);
    console.log("  ✓ 菜单已点击");

    // [3] 提取站点数据
    console.log("[3/4] 提取站点数据...");
    const { results: allData, debug } = await extractStationData(page);

    if (debug) {
      console.log(`  [调试] viewId: ${debug.viewId}, 表头${debug.headerCount}列`);
      console.log(`  [调试] 站码@${debug.stationCodeIdx} 站名@${debug.stationNameIdx} 河流@${debug.riverIdx} 1日@${debug.day1Idx}`);
      console.log(`  [调试] 表头: ${(debug.headers || []).join(" | ")}`);
      if (debug.firstRow) {
        console.log(`  [调试] 首行: ${debug.firstRow.stationCode} ${debug.firstRow.stationName} 流量${debug.firstRow.dailyFlows.length}天`);
      }
    }

    console.log(`  共提取 ${allData.length} 条记录`);

    const stationMap = new Map();
    for (const data of allData) {
      if (data.stationName) stationMap.set(data.stationName, data);
    }
    console.log(`  共 ${stationMap.size} 个不同站点`);

    if (stationMap.size > 0) {
      console.log(`  可用站名(前20): ${[...stationMap.keys()].slice(0, 20).join(", ")}`);
    }

    // [4] 匹配站点 & 计算周总量 & 导出 Excel
    console.log("[4/4] 匹配站点、计算周总量...");
    const results = [];

    for (const stationName of CONFIG.stations) {
      const data = stationMap.get(stationName);
      if (data) {
        const flowSum = data.dailyFlows.reduce((s, f) => s + f.flow, 0);
        const weekTotal = Math.round(flowSum * 86400 / 10000 * 100) / 100;

        const row = {
          站码: data.stationCode,
          站名: data.stationName,
          河流: data.riverName,
          七日合计: Math.round(flowSum * 100) / 100,
          周总量: weekTotal,
        };

        for (let i = 0; i < 7 && i < data.dailyFlows.length; i++) {
          row[week.dates[i]] = data.dailyFlows[i].flow;
        }

        results.push(row);
        console.log(`  ✓ ${stationName} 日均${Math.round(flowSum / 7 * 100) / 100} m³/s  周总量${weekTotal}万m³`);
      } else {
        console.log(`  ✗ ${stationName}: 未找到`);
      }
    }

    if (results.length === 0) {
      console.log("\n⚠ 未匹配到任何站点。");
      console.log("  请对比上面输出的「可用站名」和 CONFIG.stations。");
      return;
    }

    const filename = `水文周报_${week.monday}_${week.sunday}.xlsx`;
    const filepath = path.join(OUTPUT_DIR, filename);

    const wb = XLSX.utils.book_new();
    const wsDetail = XLSX.utils.json_to_sheet(results);
    XLSX.utils.book_append_sheet(wb, wsDetail, "水文周报");

    const summaryRows = results.map((r) => ({
      站码: r.站码, 站名: r.站名, 河流: r.河流,
      七日合计: r.七日合计, 周总量: r.周总量,
    }));
    const wsSummary = XLSX.utils.json_to_sheet(summaryRows);
    XLSX.utils.book_append_sheet(wb, wsSummary, "汇总");

    XLSX.writeFile(wb, filepath);

    console.log(`\n✅ 导出完成: ${filepath}`);
    console.log(`   站点: ${results.length}/${CONFIG.stations.length}`);
  } catch (err) {
    console.error(`\n❌ 运行出错: ${err.message}`);
    console.error(err.stack);
  } finally {
    await browser.close();
    console.log("\n浏览器已关闭");
  }
})();
