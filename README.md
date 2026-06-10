# 水文数据周报抓取工具

自动登录 [江西省水文资料在线整编系统](http://weixin.jxsswj.cn/jxhydp-app/)，提取「逐日平均流量统计」数据，计算周总量并导出 Excel。

## 快速开始

```cmd
cd /d C:\Users\夏雄\Documents\Codex\2026-06-08\new-chat-2\work
node scraper.js
```

## 配置说明

编辑 `scraper.js` 顶部的 `CONFIG` 对象：

| 配置项 | 说明 | 示例 |
|--------|------|------|
| `referenceDate` | 基准日期，自动抓取此日期**上一周**（周一~周日） | `"2026-06-02"` → 抓取 5/25~5/31 |
| `stationTypes` | 站类：`"基本站"` / `"非基本站"` / 两个都写 | `["基本站", "非基本站"]` |
| `stations` | 要抓取的站点名列表（需与网页上一致） | `["梓坊", "虬津"]` |
| `gotoTimeout` | 页面加载超时（毫秒） | `60000` |

### 日期推算规则

```
referenceDate → 找 referenceDate 所在周的周一 → 减 7 天 = 上周一 → 上周日
```

| referenceDate | 上周一 | 上周日 | 说明 |
|:---|:---|:---|:---|
| `2026-06-09`（周二） | 6/1 | 6/7 | 同月 |
| `2026-07-07`（周二） | 6/29 | 7/5 | ⚡ 跨月自动切换 |
| `2026-06-02`（周二） | 5/25 | 5/31 | ⚡ 自动切换到5月 |

### 站类切换

右上角有"基本站"和"非基本站"两个选项，脚本会自动依次切换并抓取两批数据，合并后匹配站点。

只需要在 `stations` 里写上站名（不管是基本站还是非基本站），脚本会自动找到。

## 输出文件

- `output/水文周报_YYYY-MM-DD_YYYY-MM-DD.xlsx`
- 两个 Sheet：
  - **水文周报**：站码、站名、河流、站类、周一~周日流量、七日合计、周总量
  - **汇总**：站码、站名、河流、站类、七日合计、周总量

## 常见问题

### 站点未找到
- 运行后会打印各站类的可用站点列表，从里面复制站名
- 站名必须与网页上完全一致（含括号等特殊字符）

### 跨月抓取
- 脚本自动检测，跨月时依次切换月份提取再合并

### 网络超时
- 增大 `gotoTimeout` 和 `pageLoadWait`

### 定时运行（Windows 任务计划程序）
```powershell
powershell.exe -ExecutionPolicy Bypass -File "C:\Users\夏雄\Documents\Codex\2026-06-08\new-chat-2\work\run.ps1"
```


cd /d C:\Users\夏雄\Documents\Codex\2026-06-08\new-chat-2\work
node scraper.js
