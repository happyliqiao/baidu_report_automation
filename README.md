# 百度日报自动化程序说明

本目录 `C:\Users\liqia\baidu_report_automation` 是合并后的唯一主程序目录。以后需要修改这个自动化时，先阅读本 README，再看对应脚本。

原来的 `C:\Users\liqia\baidu_report_automation_versions` 已合并进本目录的备份区：

- `backups\baidu_report_automation_versions_merged\20260525lq1720`
- `backups\20260525lq1720`

当前实际运行使用的是主目录根部的脚本和 `config.json`，不会自动使用 `backups` 里的旧版本。

## 程序目标

这个程序用于自动生成推广日报，并把日报数据同步到石墨表格。

当前覆盖的数据来源：

- 百度推广：`baidu_report.ps1`
- 360 推广：`ad_report.ps1`
- 百度 access token 刷新：`refresh_baidu_token.ps1`
- 石墨账号分表同步：`sync_shimo_account_sheets.ps1` + `scripts\shimo_account_sheet_sync.js`
- 总入口：`run_daily_report.ps1`

## 主要文件

- `config.json`：核心配置，包含百度账号、360 账号、输出 CSV 路径、石墨表格同步配置和当前 token。这个文件里有密钥和 token，修改前要备份。
- `run_daily_report.ps1`：日常运行入口，负责串起缺失日期扫描、拉取广告数据、同步石墨。
- `baidu_report.ps1`：调用百度 API，生成 `data\baidu_daily_report.csv` 和带时间戳的历史 CSV。
- `ad_report.ps1`：调用 360 API，生成 `data\360_daily_report.csv` 和带时间戳的历史 CSV。
- `refresh_baidu_token.ps1`：刷新百度账号 accessToken，并可写回 `config.json`。
- `sync_shimo_account_sheets.ps1`：PowerShell 包装入口，调用 Node 脚本处理石墨表格。
- `scripts\shimo_account_sheet_sync.js`：用 Playwright/Chrome 打开石墨表格，扫描缺失日期或写入数据。
- `test_automation.ps1`：安全测试脚本，只检查结构、配置和语法，不访问线上接口、不写石墨表。
- `data\`：输出 CSV、缺失计划 JSON、石墨 Chrome 用户数据。
- `logs\`：运行日志。
- `backups\`：历史备份，不是当前运行入口。

## 默认运行逻辑

执行：

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\run_daily_report.ps1
```

默认逻辑如下：

1. `run_daily_report.ps1` 默认启用 `-UseShimoMissingDates`。
2. 程序先调用 `sync_shimo_account_sheets.ps1 -ScanMissing`，打开石墨表格，扫描哪些账号和日期缺数据。
3. 扫描结果保存到 `data\shimo_missing_plan_yyyyMMdd_HHmmss.json`。
4. 程序根据 `config.json` 判断缺失账号属于百度还是 360。
5. 如果有百度账号缺数据，先调用 `refresh_baidu_token.ps1 -UpdateConfig -SkipMissingRefreshToken` 刷新 token。
6. 程序按缺失日期范围调用：
   - `baidu_report.ps1` 拉取百度数据
   - `ad_report.ps1` 拉取 360 数据
7. 拉取完成后，再调用 `sync_shimo_account_sheets.ps1` 把 CSV 数据写入石墨账号分表。
8. 日志写入 `logs\`，报表 CSV 写入 `data\`。

如果石墨没有扫描到缺失日期，程序会直接退出，不拉取广告数据。

## 常用命令

进入目录：

```powershell
cd C:\Users\liqia\baidu_report_automation
```

安全测试：

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\test_automation.ps1
```

默认按石墨缺失日期补数据：

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\run_daily_report.ps1
```

指定某一天，并仍然按石墨缺失逻辑执行：

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\run_daily_report.ps1 -Date 2026-05-26
```

指定日期范围，并仍然按石墨缺失逻辑执行：

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\run_daily_report.ps1 -StartDate 2026-05-20 -EndDate 2026-05-26
```

不扫描石墨，直接按日期拉取百度和 360 数据，然后同步石墨：

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "& '.\run_daily_report.ps1' -UseShimoMissingDates:`$false -Date '2026-05-26'"
```

只拉取广告数据，不同步石墨：

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "& '.\run_daily_report.ps1' -UseShimoMissingDates:`$false -SkipShimoSync -Date '2026-05-26'"
```

首次或登录状态失效时，设置石墨浏览器登录状态：

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\sync_shimo_account_sheets.ps1 -Setup
```

## 修改时优先看哪里

- 修改账号、token、输出路径、石墨表配置：看 `config.json`。
- 修改日常执行顺序：看 `run_daily_report.ps1`。
- 修改百度 API 字段或汇总逻辑：看 `baidu_report.ps1`。
- 修改 360 API 字段或汇总逻辑：看 `ad_report.ps1`。
- 修改石墨表格插入、复制模板行、字段映射、缺失扫描：看 `scripts\shimo_account_sheet_sync.js`。
- 修改 Node/Chrome 调用参数：看 `sync_shimo_account_sheets.ps1`。

## 注意事项

- `config.json` 中有真实 token、密钥和账号配置，不要随意公开。
- `data\chrome-shimo-profile` 保存石墨登录状态，删除后需要重新运行 `sync_shimo_account_sheets.ps1 -Setup` 登录。
- `node_modules` 是依赖目录，通常不要手动修改。
- `logs` 和 `data` 中大量文件是运行产物，不代表程序逻辑。
- `backups` 是历史版本保存区，除非要回滚或对比，不要把里面的文件当入口运行。
