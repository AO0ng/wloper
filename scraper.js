const { chromium } = require("playwright");
const XLSX = require("xlsx");
const path = require("path");
const fs = require("fs");

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
  await page.waitForTimeout(3000);
  await page.locator('input[name="username"]').fill(CONFIG.username);
  await page.locator('input[name="password"]').fill(CONFIG.password);
  await page.waitForTimeout(500);

  const loginBtn = page.locator('a:has-text("登 录")').first();
  if (await loginBtn.count() > 0) {
    await loginBtn.click({ force: true });
  } else {
    await page.locator('input[name="password"]').press("Enter");
  }
  await page.waitForTimeout(5000);
  return true;
}

// ==================== 通过 ExtJS Store API 提取全部数据 ====================
async function extractFromStore(page) {
  return await page.evaluate((targetStations) => {
    const grid = window.Ext.ComponentQuery.query("grid#grid-1232")[0];
    if (!grid) return { error: "grid-1232 未找到" };

    const store = grid.getStore();
    const total = store.getCount();
    const results = [];

    for (const name of targetStations) {
      const idx = store.find("stnm", name);
      if (idx < 0) {
        // 模糊匹配
        let found = false;
        for (let i = 0; i < total; i++) {
          const rec = store.getAt(i);
          if ((rec.data.stnm || "").includes(name)) {
            results.push({ stationName: name, data: rec.data, index: i });
            found = true;
            break;
          }
        }
        if (!found) results.push({ stationName: name, data: null, index: -1 });
      } else {
        const rec = store.getAt(idx);
        results.push({ stationName: name, data: rec.data, index: idx });
      }
    }

    return { total, results };
  }, CONFIG.stations);
}

// ==================== 主流程 ====================
(async () => {
  const week = getWeekRange(CONFIG.referenceDate);

  console.log("=".repeat(50));
  console.log("  水文数据周报抓取工具");
  console.log("=".repeat(50));
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
    console.log("[1/3] 登录...");
    if (!await doLogin(page)) { console.log("  登录失败"); return; }
    console.log("  ✓ 登录完成");

    // [2] 进入数据页面并提取
    console.log("[2/3] 进入逐日平均流量统计...");
    await page.goto(CONFIG.dataUrl, { waitUntil: "domcontentloaded", timeout: CONFIG.gotoTimeout });
    await page.waitForTimeout(5000);

    const menuClicked = await page.evaluate(() => {
      const items = document.querySelectorAll(".x-treelist-item-text, .x-treelist-item");
      for (const item of items) {
        if ((item.textContent || "").includes("逐日平均流量")) { item.click(); return true; }
      }
      return false;
    });
    if (!menuClicked) { console.log("  ⚠ 未找到菜单"); return; }
    console.log("  点击菜单: 逐日平均流量统计");

    // 等待 grid store 加载
    console.log("  等待数据...");
    await page.waitForTimeout(8000);

    // 通过 ExtJS API 提取
    console.log("[3/3] 提取并计算...");
    const { total, results: storeResults, error } = await extractFromStore(page);
    if (error) { console.log(`  ✗ ${error}`); return; }
    console.log(`  Store 总记录: ${total}`);

    const startDay = parseInt(week.dates[0].split("-")[2]);
    const endDay = parseInt(week.dates[6].split("-")[2]);
    const excelRows = [];

    for (const sr of storeResults) {
      if (!sr.data) {
        console.log(`  ✗ ${sr.stationName}: 未找到`);
        continue;
      }

      const d = sr.data;
      let flowSum = 0;
      let dayCount = 0;
      const detail = {};

      for (let day = startDay; day <= endDay; day++) {
        const key = `dq${day}`;
        const val = d[key];
        const flow = (val !== undefined && val !== null && val !== "") ? parseFloat(val) : null;
        detail[week.dates[dayCount]] = flow;
        if (flow !== null && !isNaN(flow)) {
          flowSum += flow;
          dayCount++;
        }
        dayCount = Math.max(dayCount, 0); // keep count correct
      }

      // 重新计数
      let actualSum = 0;
      let actualDays = 0;
      for (let day = startDay; day <= endDay; day++) {
        const val = d[`dq${day}`];
        const flow = (val !== undefined && val !== null && val !== "") ? parseFloat(val) : null;
        if (flow !== null && !isNaN(flow)) {
          actualSum += flow;
          actualDays++;
        }
      }

      if (actualDays === 0) {
        console.log(`  ✗ ${sr.stationName}: 无流量数据`);
        continue;
      }

      const weekTotal = Math.round(actualSum * 86400 / 10000 * 100) / 100;

      const row = {
        站码: d.stcd,
        站名: d.stnm,
        河流: d.rvnm,
        管理单位: d.admsnm,
        七日合计: Math.round(actualSum * 100) / 100,
        周总量: weekTotal,
      };
      for (let day = startDay; day <= endDay; day++) {
        const dateStr = week.dates[day - startDay];
        const val = d[`dq${day}`];
        row[dateStr] = (val !== undefined && val !== null && val !== "") ? parseFloat(val) : null;
      }

      excelRows.push(row);
      console.log(`  ✓ ${sr.stationName}(${d.stcd}) 日均${Math.round(actualSum / actualDays * 100) / 100} m³/s  周总量${weekTotal}万m³`);
    }

    if (excelRows.length === 0) {
      console.log("\n⚠ 未匹配到任何站点。");
      return;
    }

    // 导出 Excel
    const filename = `水文周报_${week.monday}_${week.sunday}.xlsx`;
    const filepath = path.join(OUTPUT_DIR, filename);

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(excelRows), "水文周报");

    const summaryRows = excelRows.map(r => ({
      站码: r.站码, 站名: r.站名, 河流: r.河流, 管理单位: r.管理单位,
      七日合计: r.七日合计, 周总量: r.周总量,
    }));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(summaryRows), "汇总");

    XLSX.writeFile(wb, filepath);

    console.log(`\n✅ 导出完成: ${filepath}`);
    console.log(`   站点: ${excelRows.length}/${CONFIG.stations.length}`);
  } catch (err) {
    console.error(`\n❌ 运行出错: ${err.message}`);
    console.error(err.stack);
  } finally {
    await browser.close();
    console.log("\n浏览器已关闭");
  }
})();
