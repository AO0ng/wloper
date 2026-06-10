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
  console.log("  访问登录页...");
  await page.goto(CONFIG.loginUrl, { waitUntil: "domcontentloaded", timeout: CONFIG.gotoTimeout });
  await page.waitForTimeout(3000);

  await page.locator('input[name="username"]').fill(CONFIG.username);
  await page.waitForTimeout(300);
  await page.locator('input[name="password"]').fill(CONFIG.password);
  await page.waitForTimeout(300);

  const loginBtn = page.locator('a:has-text("登 录")').first();
  if (await loginBtn.count() > 0) {
    await loginBtn.click({ force: true });
    console.log("  点击登录");
  } else {
    await page.locator('input[name="password"]').press("Enter");
  }
  await page.waitForTimeout(5000);
  return true;
}

// ==================== 数据提取 ====================
async function extractAllData(page) {
  return await page.evaluate(() => {
    const allRows = document.querySelectorAll(".x-grid-item");
    const stationRows = [];  // 8-cell: 站码、站名、河流、月均值
    const flowRows = [];     // 30-cell: 1日~30日流量

    for (let idx = 0; idx < allRows.length; idx++) {
      const row = allRows[idx];
      const cells = row.querySelectorAll(".x-grid-cell-inner");
      const texts = [];
      for (const cell of cells) texts.push((cell.textContent || "").trim());
      const recordIndex = parseInt(row.getAttribute("data-recordindex")) || -1;

      if (texts.length === 8) {
        // 站点信息行: [0]=空 [1]=序号 [2]=站码 [3]=站名 [4]=河流 [5]=月平均 [6]=月最大 [7]=月最小
        stationRows.push({
          recordIndex,
          stationCode: texts[2],
          stationName: texts[3],
          riverName: texts[4],
          monthAvg: parseFloat(texts[5]) || 0,
        });
      } else if (texts.length === 30) {
        // 逐日流量行: [0]=1日 [1]=2日 ... [29]=30日
        const flows = texts.map(t => {
          if (t === "-" || t === "") return null;
          const v = parseFloat(t);
          return isNaN(v) ? null : v;
        });
        flowRows.push({ recordIndex, flows });
      }
    }

    return { stationRows, flowRows };
  });
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
    console.log("[1/4] 登录...");
    if (!await doLogin(page)) { console.log("  登录失败"); return; }
    console.log("  ✓ 登录完成");

    // [2] 进入数据页面
    console.log("[2/4] 进入逐日平均流量统计...");
    await page.goto(CONFIG.dataUrl, { waitUntil: "domcontentloaded", timeout: CONFIG.gotoTimeout });
    await page.waitForTimeout(5000);

    const menuClicked = await page.evaluate(() => {
      const items = document.querySelectorAll(".x-treelist-item-text, .x-treelist-item");
      for (const item of items) {
        if ((item.textContent || "").includes("逐日平均流量")) { item.click(); return true; }
      }
      return false;
    });
    console.log(menuClicked ? "  点击菜单: 逐日平均流量统计" : "  ⚠ 未找到菜单");

    // 等待数据
    console.log("  等待数据加载...");
    for (let i = 0; i < 6; i++) {
      await page.waitForTimeout(4000);
      const count = await page.locator(".x-grid-item").count();
      console.log(`    轮次${i + 1}: ${count} 行`);
      if (count > 50) break;
    }

    // [3] 提取
    console.log("[3/4] 提取站点数据...");
    const { stationRows, flowRows } = await extractAllData(page);
    console.log(`  站点行: ${stationRows.length}, 流量行: ${flowRows.length}`);

    // 按 recordIndex 建立映射
    const flowMap = new Map();
    for (const f of flowRows) flowMap.set(f.recordIndex, f.flows);

    const stationMap = new Map();
    for (const s of stationRows) {
      const flows = flowMap.get(s.recordIndex) || [];
      stationMap.set(s.stationName, { ...s, dailyFlows: flows });
    }
    console.log(`  共 ${stationMap.size} 个不同站点`);
    const sampleNames = [...stationMap.keys()].slice(0, 10);
    console.log(`  可用站名(前10): ${sampleNames.join(", ")}`);

    // [4] 匹配 & 导出
    console.log("[4/4] 匹配站点、计算周总量...");
    const startDay = parseInt(week.dates[0].split("-")[2]); // 周一的日
    const endDay = parseInt(week.dates[6].split("-")[2]);   // 周日的日
    const results = [];

    for (const stationName of CONFIG.stations) {
      const s = stationMap.get(stationName);
      if (!s) { console.log(`  ✗ ${stationName}: 未找到`); continue; }

      // 取上周对应日期的流量 (日索引 = 日-1)
      let flowSum = 0;
      let dayCount = 0;
      const detail = {};
      for (let d = startDay; d <= endDay; d++) {
        const flow = s.dailyFlows[d - 1]; // 数组索引
        detail[week.dates[dayCount]] = flow;
        if (flow !== null && flow !== undefined) {
          flowSum += flow;
          dayCount++;
        }
      }

      if (dayCount === 0) { console.log(`  ✗ ${stationName}: 无流量数据`); continue; }

      const weekTotal = Math.round(flowSum * 86400 / 10000 * 100) / 100;
      const row = {
        站码: s.stationCode,
        站名: s.stationName,
        河流: s.riverName,
        七日合计: Math.round(flowSum * 100) / 100,
        周总量: weekTotal,
      };
      Object.assign(row, detail);

      results.push(row);
      console.log(`  ✓ ${stationName} 日均${Math.round(flowSum / dayCount * 100) / 100} m³/s  周总量${weekTotal}万m³`);
    }

    if (results.length === 0) {
      console.log("\n⚠ 未匹配到任何站点。");
      return;
    }

    // 导出 Excel
    const filename = `水文周报_${week.monday}_${week.sunday}.xlsx`;
    const filepath = path.join(OUTPUT_DIR, filename);

    const wb = XLSX.utils.book_new();
    const wsDetail = XLSX.utils.json_to_sheet(results);
    XLSX.utils.book_append_sheet(wb, wsDetail, "水文周报");

    const summaryRows = results.map(r => ({
      站码: r.站码, 站名: r.站名, 河流: r.河流,
      七日合计: r.七日合计, 周总量: r.周总量,
    }));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(summaryRows), "汇总");

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
