# Steam Task Assistant

A local Tkinter + Playwright assistant for operating Steam account task workflows.

The app supports four workflow modes:

- wallet code redemption
- product key activation
- game purchase
- last-played detection

This public repository intentionally contains only the program source. It does not include any real account database, passwords, wallet codes, activation keys, screenshots, customer reports, or deployment spreadsheets.

## Safety Notice

This tool can automate browser actions against Steam pages. Use it only with accounts and codes that you are authorized to operate. Keep all task databases and screenshots local. Do not commit runtime data to GitHub.

The included `.gitignore` excludes:

- SQLite task databases
- browser profiles
- screenshots
- Excel/CSV workbooks
- deployment reports
- local diagnostics and logs

## Requirements

- Windows 10/11
- Python 3.10+
- Chromium installed by Playwright

Install dependencies:

```powershell
python -m pip install -r requirements.txt
python -m playwright install chromium
```

## Run

```powershell
python steam_task_assistant\launch_gui.py
```

On Windows, you can also double-click:

```text
run_app.bat
```

## Task Import Format

Import tasks from an Excel workbook with these columns:

```text
序号
处理状态
区域
省份
城市
店编
店名
Steam账号
Steam密码
购买游戏
游戏价格
兑换码列表
激活码列表
模式
备注
```

The `模式` value should match one of:

```text
兑换码兑换
激活码激活
游戏购买
游玩时间检测
```

Multiple wallet codes or activation keys can be separated by line breaks, commas, semicolons, or whitespace.

## Runtime Files

The app creates local runtime folders next to the source:

```text
steam_task_assistant/data/
steam_task_assistant/screenshots/
steam_task_assistant/browser_profile/
```

These folders are ignored by git and should remain private.

## Network Note

By default the Playwright Chromium instance bypasses the Windows system proxy to avoid Steam authentication requests being closed by local proxy rules. To force usage of the system proxy, set:

```powershell
$env:STEAM_ASSISTANT_USE_SYSTEM_PROXY = "1"
```

## Public Release Checklist

Before publishing changes, run:

```powershell
python -m py_compile steam_task_assistant\app.py steam_task_assistant\launch_gui.py
git status --ignored -sb
```

Confirm that no real `.xlsx`, `.sqlite3`, screenshots, account exports, wallet codes, or product keys are staged.
