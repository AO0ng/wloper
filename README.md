# 水文数据周报抓取工具

自动登录 [江西省水文资料在线整编系统](http://weixin.jxsswj.cn/jxhydp-app/)，提取「逐日平均流量统计」数据，匹配目标站点计算周总量并导出 Excel。

## 快速开始

```cmd
cd /d C:\Users\夏雄\Documents\Codex\2026-06-08\new-chat-2\work
node scraper.js
```

## 配置说明

编辑 `scraper.js` 顶部的 `CONFIG` 对象：

| 配置项 | 说明 | 示例 |
|--------|------|------|
| `referenceDate` | 基准日期，自动抓取上一周（周一~周日） | `"2026-06-09"` → 抓取 6/1~6/7 |
| `stations` | 要抓取的站点名列表（需与网页上一致） | `["梓坊", "虬津"]` |
| `username` / `password` | 登录账号密码 | — |
| `gotoTimeout` | 页面加载超时（毫秒） | `60000` |

### 日期推算

```
referenceDate → 所在周的周一 → 减 7 天 = 上周一 ~ 上周日
```

| referenceDate | 抓取周期 |
|:---|:---|
| `2026-06-09`（周二） | 6/1 ~ 6/7 |
| `2026-07-07`（周二） | 6/29 ~ 7/5 |

## 输出

- `output/水文周报_YYYY-MM-DD_YYYY-MM-DD.xlsx`
- 两个 Sheet：**水文周报**（含逐日流量）、**汇总**

## 常见问题

### 站点未找到
- 运行后会打印全部可用站名，从里面核对站名
- 站名必须与网页上完全一致

### 网络超时
- 增大 `gotoTimeout`

### 定时运行（Windows 任务计划程序）
```powershell
powershell.exe -ExecutionPolicy Bypass -File "C:\Users\夏雄\Documents\Codex\2026-06-08\new-chat-2\work\run.ps1"
```
