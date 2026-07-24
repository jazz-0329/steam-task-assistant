# Steam Task Assistant

A local Windows Tkinter + Playwright assistant for operating authorized Steam account task workflows.

This public repository intentionally contains only the program source. It does not include any real account database, passwords, wallet codes, activation keys, screenshots, customer reports, browser profiles, or deployment spreadsheets.

## Workflow Modes

- Wallet code redemption
- Product key activation
- Game purchase
- Last-played detection
- Friend invite link pickup and automatic gift claim

## Safety Notice

This tool can automate browser actions against Steam pages. Use it only with accounts, codes, and purchases that you are authorized to operate.

Keep all task databases, account workbooks, screenshots, and generated reports local. Do not commit runtime data to GitHub.

The included `.gitignore` excludes:

- SQLite task databases
- Browser profiles
- Screenshots
- Excel/CSV/TSV workbooks
- Deployment reports
- Local diagnostics and logs

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

Debug launch, with a console window kept open:

```text
run_app_debug.bat
```

## Task Import Format

Import tasks from an Excel workbook. The app recognizes these Chinese column names:

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
提货码
提货网址
账号链接
好友链接1
好友链接2
提货状态
```

The `模式` value should match one of:

```text
兑换码兑换
激活码激活
游戏购买
游玩时间检测
好友码提货
```

Multiple wallet codes or activation keys can be separated by line breaks, commas, semicolons, or whitespace.

Friend-code pickup tasks may include pickup code, pickup URL, profile link, and two friend invite link fields. The app writes collected invite links and claim status into the local SQLite task database only.

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

Confirm that no real `.xlsx`, `.sqlite3`, screenshots, account exports, wallet codes, product keys, or browser profile files are staged.
