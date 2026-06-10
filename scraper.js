const { chromium } = require("playwright");
const XLSX = require("xlsx");
const path = require("path");
const fs = require("fs");

// ==================== 配置区 ====================
const CONFIG = {
  referenceDate: "2026-06-09",
  stations: ["梓坊", "先锋", "土桥", "渣津", "瑞昌", "萍乡", "虬津"],
  gotoTimeout: 60000,
  pageLoadWait: 5000,
  loginUrl: "http://weixin.jxsswj.cn/jxhydp-app/#login",
  dataUrl: "http://weixin.jxsswj.cn/jxhydp-app/#hydataview",
  username: "夏雄",
  password: "Jxsw_123",
};

const OUTPUT_DIR = path.join(__dirname, "output");
const DEBUG_DIR = path.join(__dirname, "debug");

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

// ==================== 调试辅助 ====================
function debugLog(page, msg) {
  const ts = new Date().toISOString().slice(11, 19);
  return page.evaluate((m) => { console.log(`[${m}]`); }, `${ts} ${msg}`);
}

async function debugSave(page, label) {
  try {
    fs.mkdirSync(DEBUG_DIR, { recursive: true });
    const html = await page.content();
    fs.writeFileSync(path.join(DEBUG_DIR, `${label}.html`), html);
    await page.screenshot({ path: path.join(DEBUG_DIR, `${label}.png`), fullPage: true });
    console.log(`  [调试] ${label}.html + ${label}.png 已保存`);
  } catch (e) {
    console.log(`  [调试] 保存失败: ${e.message}`);
  }
}

// ==================== 数据提取（多策略） ====================
async function extractStationData(page) {
  const results = [];

  // 等待可能的 AJAX 加载
  await page.waitForTimeout(3000);

  // 策略1：.x-grid-item（ExtJS 6+ 经典主题）
  let rows1 = await page.locator(".x-grid-item").all();
  console.log(`  [策略1] .x-grid-item: ${rows1.length} 行`);

  // 策略2：.x-grid-row（旧版 ExtJS）
  let rows2 = await page.locator(".x-grid-row").all();
  console.log(`  [策略2] .x-grid-row: ${rows2.length} 行`);

  // 策略3：table.x-grid-table（ExtJS 表格）
  let rows3 = await page.locator("table.x-grid-table tr").all();
  console.log(`  [策略3] table.x-grid-table tr: ${rows3.length} 行`);

  // 策略4：.x-grid-view 下的所有 table tr
  let rows4 = await page.locator(".x-grid-view table tr").all();
  console.log(`  [策略4] .x-grid-view table tr: ${rows4.length} 行`);

  // 策略5：查找所有带 data- 属性的 grid 容器
  const dataViews = await page.locator("[data-boundview]").all();
  console.log(`  [策略5] [data-boundview]: ${dataViews.length} 个`);
  for (const dv of dataViews) {
    const boundview = await dv.getAttribute("data-boundview");
    const trCount = await dv.locator("table tr").count();
    console.log(`    ${boundview}: ${trCount} 行`);
  }

  // 策略6：直接查所有 table，打印表头
  const allTables = await page.locator("table").all();
  console.log(`  [策略6] 页面共 ${allTables.length} 个 table`);
  for (let i = 0; i < allTables.length; i++) {
    const table = allTables[i];
    const trs = await table.locator("tr").all();
    if (trs.length < 2) continue;
    // 读第一行看表头
    const headerCells = await trs[0].locator("th, td").all();
    const headerTexts = [];
    for (const hc of headerCells.slice(0, 15)) {
      headerTexts.push((await hc.textContent()).trim().substring(0, 10));
    }
    console.log(`    table[${i}] ${trs.length}行 表头: [${headerTexts.join(" | ")}]`);
  }

  // 策略7：检查 iframe
  const iframes = await page.locator("iframe").all();
  console.log(`  [策略7] iframe: ${iframes.length} 个`);

  // 实际提取：优先用 .x-grid-item
  let rows = rows1;
  let extractor = "grid-item";

  if (rows.length === 0) rows = rows3; extractor = "grid-table";
  if (rows.length === 0) rows = rows2; extractor = "grid-row";
  if (rows.length === 0) rows = rows4; extractor = "grid-view";

  console.log(`  使用策略: ${extractor}, ${rows.length} 行`);

  for (const row of rows) {
    const cells = await row.locator(".x-grid-cell-inner, .x-grid-cell, td, th").all();
    const texts = [];
    for (const cell of cells) {
      texts.push((await cell.textContent()).trim());
    }
    if (texts.length < 3) continue;

    // 打印前3行样本
    if (results.length < 3) {
      console.log(`  样本行[${results.length}]: ${JSON.stringify(texts.slice(0, 10))}`);
    }

    const stationName = texts[1] || "";
    const stationCode = texts[0] || "";
    const riverName = texts[2] || "";
    const dailyFlows = [];
    for (let i = 6; i < texts.length; i++) {
      const val = parseFloat(texts[i]);
      if (!isNaN(val)) {
        dailyFlows.push({ day: i - 5, flow: val });
      }
    }
    results.push({ stationCode, stationName, riverName, dailyFlows });
  }

  return results;
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
  fs.mkdirSync(DEBUG_DIR, { recursive: true });

  const browser = await chromium.launch({ headless: true, args: ["--no-sandbox"] });
  const context = await browser.newContext({ locale: "zh-CN", viewport: { width: 1920, height: 1080 } });
  const page = await context.newPage();

  try {
    // [1] 登录
    console.log("[1/4] 登录...");
    await page.goto(CONFIG.loginUrl, { waitUntil: "networkidle", timeout: CONFIG.gotoTimeout });
    await page.waitForTimeout(3000);

    const usernameInput = page.locator('input[name="username"]');
    const passwordInput = page.locator('input[name="password"]');
    const loginBtn = page.locator('a:has-text("登 录"), button:has-text("登录"), button:has-text("登 录")');

    if (await usernameInput.count() > 0) {
      await usernameInput.fill(CONFIG.username);
      await passwordInput.fill(CONFIG.password);
      if (await loginBtn.count() > 0) {
        await loginBtn.first().click();
      } else {
        await passwordInput.press("Enter");
      }
      await page.waitForTimeout(4000);
    }
    console.log("  ✓ 登录完成");

    // [2] 进入数据视图并点击菜单
    console.log("[2/4] 进入逐日平均流量统计...");
    await page.goto(CONFIG.dataUrl, { waitUntil: "networkidle", timeout: CONFIG.gotoTimeout });
    await page.waitForTimeout(CONFIG.pageLoadWait);

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
    await page.waitForTimeout(5000);
    console.log("  ✓ 菜单已点击");

    // 保存页面状态供调试
    await debugSave(page, "page-state");

    // [3] 提取站点数据
    console.log("[3/4] 提取站点数据...");
    const allData = await extractStationData(page);
    console.log(`  共提取 ${allData.length} 条记录`);

    // 建立站名索引
    const stationMap = new Map();
    for (const data of allData) {
      if (data.stationName) stationMap.set(data.stationName, data);
    }
    console.log(`  共 ${stationMap.size} 个不同站点`);

    // 打印可用站名
    if (stationMap.size > 0) {
      console.log(`  可用站名: ${[...stationMap.keys()].slice(0, 20).join(", ")}`);
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
      console.log("  请查看 debug/page-state.html 分析页面结构。");
      console.log("  如果页面有筛选条件（如日期、管理单位），可能需要先设置。");
      return;
    }

    // 导出 Excel
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
