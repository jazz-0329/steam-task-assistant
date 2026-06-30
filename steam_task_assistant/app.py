from __future__ import annotations

import csv
import json
import os
import queue
import re
import sqlite3
import sys
import threading
import time
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from tkinter import BOTH, BOTTOM, END, LEFT, RIGHT, TOP, X, Y, BooleanVar, IntVar, StringVar, Tk, filedialog, messagebox, ttk

import openpyxl


APP_DIR = Path(__file__).resolve().parent
DATA_DIR = APP_DIR / "data"
SCREENSHOT_DIR = Path(os.environ.get("STEAM_ASSISTANT_SCREENSHOT_DIR", APP_DIR / "screenshots"))
DB_PATH = Path(os.environ.get("STEAM_ASSISTANT_DB_PATH", DATA_DIR / "tasks.sqlite3"))
AUTOMATION_LOG_PATH = APP_DIR / "自动化错误.log"

STEAM_LOGIN_URL = "https://store.steampowered.com/login/"
STEAM_REDEEM_URL = "https://store.steampowered.com/account/redeemwalletcode"
STEAM_ACTIVATE_URL = "https://store.steampowered.com/account/registerkey"
STEAM_HISTORY_URL = "https://store.steampowered.com/account/history/"
STEAM_LICENSES_URL = "https://store.steampowered.com/account/licenses/"
STEAM_CART_URL = "https://store.steampowered.com/cart/"
STEAM_SEARCH_URL = "https://store.steampowered.com/search/?term="

MODES = ("兑换码兑换", "激活码激活", "游戏购买", "游玩时间检测")
STATUSES = ("未处理", "处理中", "需要人工处理", "成功", "部分成功", "失败", "跳过")
REQUIRED_COLUMNS = {
    "序号",
    "处理状态",
    "区域",
    "省份",
    "城市",
    "店编",
    "店名",
    "Steam账号",
    "Steam密码",
    "购买游戏",
    "游戏价格",
    "兑换码列表",
    "激活码列表",
    "模式",
    "备注",
}

GAME_CATALOG = {
    "红色沙漠": {
        "official_name": "红色沙漠",
        "license_names": ["红色沙漠", "Crimson Desert"],
        "app_id": 3321460,
        "package_id": 1176092,
        "price": 268.0,
    },
    "极限竞速6": {
        "official_name": "极限竞速：地平线 6",
        "license_names": ["极限竞速：地平线 6", "Forza Horizon 6"],
        "app_id": 2483190,
        "package_id": 950527,
        "price": 298.0,
    },
    "极限竞速地平线6": {
        "official_name": "极限竞速：地平线 6",
        "license_names": ["极限竞速：地平线 6", "Forza Horizon 6"],
        "app_id": 2483190,
        "package_id": 950527,
        "price": 298.0,
    },
}


def ensure_dirs() -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    SCREENSHOT_DIR.mkdir(parents=True, exist_ok=True)


def clean_filename_part(value: object, fallback: str = "空") -> str:
    text = str(value or fallback).strip()
    text = re.sub(r'[\\/:*?"<>|\r\n\t]+', "_", text)
    text = re.sub(r"\s+", " ", text)
    return text[:80] or fallback


def split_codes(value: object) -> list[str]:
    if value is None:
        return []
    parts = re.split(r"[\n\r;；,，]+", str(value))
    return [part.strip() for part in parts if part.strip()]


def normalize_identity_text(value: object) -> str:
    text = str(value or "").lower()
    text = re.sub(r"alienware|外星人", "", text, flags=re.IGNORECASE)
    text = re.sub(r"[^0-9a-z\u4e00-\u9fff]+", "", text)
    return text


def location_token(value: object) -> str:
    return re.sub(
        r"(特别行政区|壮族自治区|回族自治区|维吾尔自治区|自治区|省|市)$",
        "",
        str(value or ""),
    ).strip()


def store_identity_core(store_name: object, province: object, city: object) -> str:
    core = normalize_identity_text(store_name)
    removable = [
        normalize_identity_text(province),
        normalize_identity_text(city),
        normalize_identity_text(location_token(province)),
        normalize_identity_text(location_token(city)),
        "专卖店",
        "旗舰店",
        "体验店",
        "门店",
        "店",
    ]
    for item in sorted({item for item in removable if item}, key=len, reverse=True):
        core = core.replace(item, "")
    return core


def normalize_game_name(value: object) -> str:
    return re.sub(r"[\s:：·_\-]+", "", str(value or "").strip()).lower()


def resolve_game_product(game_name: object) -> dict | None:
    normalized = normalize_game_name(game_name)
    for alias, product in GAME_CATALOG.items():
        if normalize_game_name(alias) == normalized:
            return product
    return None


def product_in_text(product: dict, text: str) -> bool:
    normalized_text = normalize_game_name(text)
    names = product.get("license_names") or [product["official_name"]]
    return any(normalize_game_name(name) in normalized_text for name in names)


def copy_to_clipboard(root: Tk, value: str) -> None:
    root.clipboard_clear()
    root.clipboard_append(value)
    root.update()


@dataclass
class Task:
    id: int
    row_no: int
    status: str
    region: str
    province: str
    city: str
    store_code: str
    store_name: str
    steam_account: str
    steam_password: str
    game_name: str
    game_price: str
    voucher_codes: list[str]
    activation_codes: list[str]
    mode: str
    note: str
    screenshot_dir: str
    profile_link: str
    last_played_game: str
    last_played_time: str
    last_played_days: str
    last_played_source: str


class TaskStore:
    def __init__(self, db_path: Path) -> None:
        self.conn = sqlite3.connect(db_path, timeout=30)
        self.conn.row_factory = sqlite3.Row
        self._init_schema()

    def _init_schema(self) -> None:
        self.conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS tasks (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                row_no INTEGER,
                status TEXT NOT NULL DEFAULT '未处理',
                region TEXT,
                province TEXT,
                city TEXT,
                store_code TEXT,
                store_name TEXT,
                steam_account TEXT,
                steam_password TEXT,
                game_name TEXT,
                game_price TEXT,
                voucher_codes TEXT,
                activation_codes TEXT,
                mode TEXT,
                note TEXT,
                screenshot_dir TEXT,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                updated_at TEXT DEFAULT CURRENT_TIMESTAMP
            );
            CREATE TABLE IF NOT EXISTS events (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                task_id INTEGER,
                event_time TEXT DEFAULT CURRENT_TIMESTAMP,
                event_type TEXT,
                message TEXT,
                screenshot_path TEXT
            );
            CREATE TABLE IF NOT EXISTS voucher_results (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                task_id INTEGER NOT NULL,
                code_index INTEGER NOT NULL,
                code TEXT NOT NULL,
                status TEXT NOT NULL DEFAULT '待兑换',
                message TEXT,
                screenshot_path TEXT,
                updated_at TEXT DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(task_id, code_index)
            );
            """
        )
        existing_columns = {
            row["name"] for row in self.conn.execute("PRAGMA table_info(tasks)").fetchall()
        }
        for column_name in (
            "profile_link",
            "last_played_game",
            "last_played_time",
            "last_played_days",
            "last_played_source",
        ):
            if column_name not in existing_columns:
                self.conn.execute(f"ALTER TABLE tasks ADD COLUMN {column_name} TEXT")
        self.conn.commit()

    def clear_tasks(self) -> None:
        self.conn.execute("DELETE FROM voucher_results")
        self.conn.execute("DELETE FROM events")
        self.conn.execute("DELETE FROM tasks")
        self.conn.commit()

    def import_excel(self, path: Path, replace: bool = True) -> int:
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb["任务数据库"] if "任务数据库" in wb.sheetnames else wb.active
        headers = [str(cell.value).strip() if cell.value is not None else "" for cell in ws[1]]
        missing = REQUIRED_COLUMNS - set(headers)
        if missing:
            raise ValueError("缺少标准列：" + "、".join(sorted(missing)))
        index = {name: headers.index(name) for name in headers if name}
        if replace:
            self.clear_tasks()
        count = 0
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not any(row):
                continue
            mode = str(row[index["模式"]] or "兑换码兑换").strip()
            if mode not in MODES:
                mode = "兑换码兑换"
            status = str(row[index["处理状态"]] or "未处理").strip()
            if status not in STATUSES:
                status = "未处理"
            values = {
                "row_no": row[index["序号"]] or count + 1,
                "status": status,
                "region": row[index["区域"]],
                "province": row[index["省份"]],
                "city": row[index["城市"]],
                "store_code": row[index["店编"]],
                "store_name": row[index["店名"]],
                "steam_account": row[index["Steam账号"]],
                "steam_password": row[index["Steam密码"]],
                "game_name": row[index["购买游戏"]],
                "game_price": row[index["游戏价格"]],
                "voucher_codes": "\n".join(split_codes(row[index["兑换码列表"]])),
                "activation_codes": "\n".join(split_codes(row[index["激活码列表"]])),
                "mode": mode,
                "note": row[index["备注"]],
                "screenshot_dir": "",
            }
            self.conn.execute(
                """
                INSERT INTO tasks (
                    row_no, status, region, province, city, store_code, store_name,
                    steam_account, steam_password, game_name, game_price, voucher_codes,
                    activation_codes, mode, note, screenshot_dir
                ) VALUES (
                    :row_no, :status, :region, :province, :city, :store_code, :store_name,
                    :steam_account, :steam_password, :game_name, :game_price, :voucher_codes,
                    :activation_codes, :mode, :note, :screenshot_dir
                )
                """,
                values,
            )
            count += 1
        self.conn.commit()
        return count

    def import_playtime_tasks(self, workspace_root: Path, replace: bool = True) -> int:
        v38_candidates = sorted(workspace_root.glob("*V3.8*.xlsx"))
        if not v38_candidates:
            raise FileNotFoundError("未找到 V3.8 账号清单，无法匹配账号链接。")
        v38_rows = self._read_v38_account_rows(v38_candidates[0])
        account_rows = [row for row in v38_rows if row.get("sonic_code")]
        if not account_rows:
            raise ValueError("V3.8 账号清单的《索尼克赛车交叉世界》列下没有可用激活码，无法生成检测任务。")
        by_account = {
            str(row.get("steam_account") or "").strip().lower(): row
            for row in v38_rows
            if row.get("steam_account")
        }
        by_store = {
            str(row.get("store_code") or "").strip().lower(): row
            for row in v38_rows
            if row.get("store_code")
        }

        if replace:
            ids = [
                row["id"]
                for row in self.conn.execute("SELECT id FROM tasks WHERE mode = '游玩时间检测'").fetchall()
            ]
            if ids:
                placeholders = ",".join("?" for _ in ids)
                self.conn.execute(f"DELETE FROM voucher_results WHERE task_id IN ({placeholders})", ids)
                self.conn.execute(f"DELETE FROM events WHERE task_id IN ({placeholders})", ids)
                self.conn.execute(f"DELETE FROM tasks WHERE id IN ({placeholders})", ids)

        count = 0
        by_sonic_code = {
            str(row.get("sonic_code") or "").strip().lower(): row
            for row in v38_rows
            if row.get("sonic_code")
        }

        for index, row in enumerate(account_rows, start=1):
            account = str(row.get("steam_account") or "").strip()
            if not account:
                continue
            store_code = str(row.get("store_code") or "").strip()
            sonic_code = str(row.get("sonic_code") or "").strip()
            v38_row = (
                by_account.get(account.lower())
                or by_store.get(store_code.lower())
                or by_sonic_code.get(sonic_code.lower())
                or {}
            )
            values = {
                "row_no": row.get("row_no") or index,
                "status": "未处理",
                "region": row.get("region") or v38_row.get("region"),
                "province": row.get("province") or v38_row.get("province"),
                "city": row.get("city") or v38_row.get("city"),
                "store_code": store_code or v38_row.get("store_code"),
                "store_name": row.get("store_name") or v38_row.get("store_name"),
                "steam_account": account,
                "steam_password": row.get("steam_password") or v38_row.get("steam_password"),
                "game_name": "",
                "game_price": "",
                "voucher_codes": "",
                "activation_codes": sonic_code,
                "mode": "游玩时间检测",
                "note": "V3.8索尼克赛车交叉世界激活码账号范围；检测账号上一次游玩的游戏与时间",
                "screenshot_dir": "",
                "profile_link": v38_row.get("profile_link") or "",
                "last_played_game": "",
                "last_played_time": "",
                "last_played_days": "",
                "last_played_source": "",
            }
            self.conn.execute(
                """
                INSERT INTO tasks (
                    row_no, status, region, province, city, store_code, store_name,
                    steam_account, steam_password, game_name, game_price, voucher_codes,
                    activation_codes, mode, note, screenshot_dir, profile_link,
                    last_played_game, last_played_time, last_played_days, last_played_source
                ) VALUES (
                    :row_no, :status, :region, :province, :city, :store_code, :store_name,
                    :steam_account, :steam_password, :game_name, :game_price, :voucher_codes,
                    :activation_codes, :mode, :note, :screenshot_dir, :profile_link,
                    :last_played_game, :last_played_time, :last_played_days, :last_played_source
                )
                """,
                values,
            )
            count += 1
        self.conn.commit()
        return count

    def _read_sonic_racing_rows(self, path: Path) -> list[dict[str, str]]:
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        ws = wb.worksheets[0]
        header_row = 3
        sonic_col = None
        for cell in ws[header_row]:
            text = str(cell.value or "")
            if "索尼克赛车交叉世界" in text:
                sonic_col = cell.column
                break
        if sonic_col is None:
            raise ValueError("Q1 FY27 游戏部署表中未找到《索尼克赛车交叉世界》列。")

        rows = []
        for row_index in range(header_row + 1, ws.max_row + 1):
            code = str(ws.cell(row_index, sonic_col).value or "").strip()
            if not code or code == "已部署":
                continue
            rows.append(
                {
                    "row_no": ws.cell(row_index, 1).value,
                    "region": str(ws.cell(row_index, 2).value or "").strip(),
                    "province": str(ws.cell(row_index, 3).value or "").strip(),
                    "city": str(ws.cell(row_index, 4).value or "").strip(),
                    "store_code": str(ws.cell(row_index, 6).value or "").strip(),
                    "valid_state": str(ws.cell(row_index, 7).value or "").strip(),
                    "store_name": str(ws.cell(row_index, 8).value or "").strip(),
                    "steam_account": str(ws.cell(row_index, 9).value or "").strip(),
                    "steam_password": str(ws.cell(row_index, 10).value or "").strip(),
                    "sonic_code": code,
                }
            )
        return rows

    def _read_v38_account_rows(self, path: Path) -> list[dict[str, str]]:
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        ws = wb.worksheets[0]
        rows = []
        for row_index in range(4, ws.max_row + 1):
            account = ws.cell(row_index, 15).value
            if not account:
                continue
            rows.append(
                {
                    "row_no": ws.cell(row_index, 1).value,
                    "region": str(ws.cell(row_index, 2).value or "").strip(),
                    "province": str(ws.cell(row_index, 3).value or "").strip(),
                    "city": str(ws.cell(row_index, 4).value or "").strip(),
                    "store_code": str(ws.cell(row_index, 6).value or "").strip(),
                    "valid_state": str(ws.cell(row_index, 7).value or "").strip(),
                    "store_name": str(ws.cell(row_index, 8).value or "").strip(),
                    "profile_link": str(ws.cell(row_index, 12).value or "").strip(),
                    "steam_account": str(ws.cell(row_index, 15).value or "").strip(),
                    "steam_password": str(ws.cell(row_index, 16).value or "").strip(),
                    "sonic_code": str(ws.cell(row_index, 22).value or "").strip(),
                }
            )
        return rows

    def filtered_tasks(self, mode_filter: str) -> list[Task]:
        if mode_filter == "全部模式":
            rows = self.conn.execute("SELECT * FROM tasks ORDER BY row_no, id").fetchall()
        else:
            rows = self.conn.execute("SELECT * FROM tasks WHERE mode = ? ORDER BY row_no, id", (mode_filter,)).fetchall()
        return [self._row_to_task(row) for row in rows]

    def get_task(self, task_id: int | None) -> Task | None:
        if task_id is None:
            return None
        row = self.conn.execute("SELECT * FROM tasks WHERE id = ?", (task_id,)).fetchone()
        return self._row_to_task(row) if row else None

    def next_task_after(self, task_id: int | None, mode_filter: str) -> Task | None:
        mode_clause = "" if mode_filter == "全部模式" else " AND mode = ?"
        mode_params = () if mode_filter == "全部模式" else (mode_filter,)
        if task_id is None:
            row = self.conn.execute(
                f"SELECT * FROM tasks WHERE status IN ('未处理','处理中','需要人工处理'){mode_clause} ORDER BY row_no, id LIMIT 1",
                mode_params,
            ).fetchone()
        else:
            current = self.conn.execute("SELECT row_no, id FROM tasks WHERE id = ?", (task_id,)).fetchone()
            if not current:
                return self.next_task_after(None, mode_filter)
            row = self.conn.execute(
                f"""
                SELECT * FROM tasks
                WHERE (row_no > ? OR (row_no = ? AND id > ?))
                  AND status IN ('未处理','处理中','需要人工处理')
                  {mode_clause}
                ORDER BY row_no, id LIMIT 1
                """,
                (current["row_no"], current["row_no"], current["id"], *mode_params),
            ).fetchone()
        return self._row_to_task(row) if row else None

    def pending_purchase_tasks(
        self,
        start_task_id: int | None,
        limit: int,
    ) -> list[Task]:
        tasks = [
            task
            for task in self.filtered_tasks("游戏购买")
            if task.status in {"未处理", "处理中", "需要人工处理"}
        ]
        if start_task_id is not None:
            start_index = next(
                (index for index, task in enumerate(tasks) if task.id == start_task_id),
                0,
            )
            tasks = tasks[start_index:]
        return tasks if limit == 0 else tasks[:limit]

    def update_status(self, task_id: int, status: str) -> None:
        self.conn.execute("UPDATE tasks SET status = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?", (status, task_id))
        self.conn.commit()

    def update_screenshot_dir(self, task_id: int, screenshot_dir: Path) -> None:
        self.conn.execute(
            "UPDATE tasks SET screenshot_dir = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?",
            (str(screenshot_dir), task_id),
        )
        self.conn.commit()

    def upsert_voucher_result(
        self,
        task_id: int,
        code_index: int,
        code: str,
        status: str,
        message: str = "",
        screenshot_path: str = "",
    ) -> None:
        self.conn.execute(
            """
            INSERT INTO voucher_results (
                task_id, code_index, code, status, message, screenshot_path, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
            ON CONFLICT(task_id, code_index) DO UPDATE SET
                code = excluded.code,
                status = excluded.status,
                message = excluded.message,
                screenshot_path = excluded.screenshot_path,
                updated_at = CURRENT_TIMESTAMP
            """,
            (task_id, code_index, code, status, message, screenshot_path),
        )
        self.conn.commit()

    def voucher_results(self, task_id: int) -> dict[int, sqlite3.Row]:
        rows = self.conn.execute(
            "SELECT * FROM voucher_results WHERE task_id = ? ORDER BY code_index",
            (task_id,),
        ).fetchall()
        return {int(row["code_index"]): row for row in rows}

    def pending_redeem_tasks(self, start_task_id: int | None = None, limit: int = 0) -> list[Task]:
        rows = self.conn.execute(
            """
            SELECT * FROM tasks
            WHERE mode = '兑换码兑换'
              AND status IN ('未处理','处理中','需要人工处理')
              AND COALESCE(voucher_codes, '') <> ''
            ORDER BY row_no, id
            """
        ).fetchall()
        tasks = [self._row_to_task(row) for row in rows]
        if start_task_id is not None:
            start_index = next((i for i, task in enumerate(tasks) if task.id == start_task_id), 0)
            tasks = tasks[start_index:]
        if limit > 0:
            tasks = tasks[:limit]
        return tasks

    def pending_activation_tasks(self, start_task_id: int | None = None, limit: int = 0) -> list[Task]:
        rows = self.conn.execute(
            """
            SELECT * FROM tasks
            WHERE mode = '激活码激活'
              AND status IN ('未处理','处理中','需要人工处理')
              AND COALESCE(activation_codes, '') <> ''
            ORDER BY row_no, id
            """
        ).fetchall()
        tasks = [self._row_to_task(row) for row in rows]
        if start_task_id is not None:
            start_index = next((i for i, task in enumerate(tasks) if task.id == start_task_id), 0)
            tasks = tasks[start_index:]
        if limit > 0:
            tasks = tasks[:limit]
        return tasks

    def pending_playtime_tasks(self, start_task_id: int | None = None, limit: int = 0) -> list[Task]:
        rows = self.conn.execute(
            """
            SELECT * FROM tasks
            WHERE mode = '游玩时间检测'
              AND status IN ('未处理','处理中','需要人工处理')
            ORDER BY row_no, id
            """
        ).fetchall()
        tasks = [self._row_to_task(row) for row in rows]
        if start_task_id is not None:
            start_index = next((i for i, task in enumerate(tasks) if task.id == start_task_id), 0)
            tasks = tasks[start_index:]
        if limit > 0:
            tasks = tasks[:limit]
        return tasks

    def update_playtime_result(
        self,
        task_id: int,
        status: str,
        game: str,
        played_time: str,
        days: str,
        source: str,
        note: str = "",
        screenshot_path: str = "",
    ) -> None:
        self.conn.execute(
            """
            UPDATE tasks
            SET status = ?,
                last_played_game = ?,
                last_played_time = ?,
                last_played_days = ?,
                last_played_source = ?,
                note = CASE WHEN ? <> '' THEN ? ELSE note END,
                screenshot_dir = CASE WHEN ? <> '' THEN ? ELSE screenshot_dir END,
                updated_at = CURRENT_TIMESTAMP
            WHERE id = ?
            """,
            (
                status,
                game,
                played_time,
                days,
                source,
                note,
                note,
                str(Path(screenshot_path).parent) if screenshot_path else "",
                str(Path(screenshot_path).parent) if screenshot_path else "",
                task_id,
            ),
        )
        self.conn.commit()

    def add_event(self, task_id: int, event_type: str, message: str, screenshot_path: str = "") -> None:
        self.conn.execute(
            "INSERT INTO events (task_id, event_type, message, screenshot_path) VALUES (?, ?, ?, ?)",
            (task_id, event_type, message, screenshot_path),
        )
        self.conn.commit()

    def events(self, task_id: int | None, limit: int = 250) -> list[sqlite3.Row]:
        if task_id is None:
            return []
        return self.conn.execute(
            "SELECT * FROM events WHERE task_id = ? ORDER BY id DESC LIMIT ?",
            (task_id, limit),
        ).fetchall()

    def export_csv(self, path: Path) -> None:
        rows = self.conn.execute("SELECT * FROM tasks ORDER BY row_no, id").fetchall()
        with path.open("w", encoding="utf-8-sig", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(
                [
                    "序号",
                    "处理状态",
                    "区域",
                    "省份",
                    "城市",
                    "店编",
                    "店名",
                    "Steam账号",
                    "Steam密码",
                    "购买游戏",
                    "游戏价格",
                    "兑换码列表",
                    "激活码列表",
                    "模式",
                    "备注",
                    "截图文件夹",
                    "账号链接",
                    "上次游玩游戏",
                    "上次游玩时间",
                    "距离今天",
                    "检测来源",
                ]
            )
            for row in rows:
                writer.writerow(
                    [
                        row["row_no"],
                        row["status"],
                        row["region"],
                        row["province"],
                        row["city"],
                        row["store_code"],
                        row["store_name"],
                        row["steam_account"],
                        row["steam_password"],
                        row["game_name"],
                        row["game_price"],
                        row["voucher_codes"],
                        row["activation_codes"],
                        row["mode"],
                        row["note"],
                        row["screenshot_dir"],
                        row["profile_link"],
                        row["last_played_game"],
                        row["last_played_time"],
                        row["last_played_days"],
                        row["last_played_source"],
                    ]
                )

    def _row_to_task(self, row: sqlite3.Row) -> Task:
        keys = set(row.keys())
        return Task(
            id=row["id"],
            row_no=row["row_no"],
            status=row["status"] or "未处理",
            region=row["region"] or "",
            province=row["province"] or "",
            city=row["city"] or "",
            store_code=row["store_code"] or "",
            store_name=row["store_name"] or "",
            steam_account=row["steam_account"] or "",
            steam_password=row["steam_password"] or "",
            game_name=row["game_name"] or "",
            game_price=str(row["game_price"] or ""),
            voucher_codes=split_codes(row["voucher_codes"]),
            activation_codes=split_codes(row["activation_codes"]),
            mode=row["mode"] or "兑换码兑换",
            note=row["note"] or "",
            screenshot_dir=row["screenshot_dir"] or "",
            profile_link=row["profile_link"] or "" if "profile_link" in keys else "",
            last_played_game=row["last_played_game"] or "" if "last_played_game" in keys else "",
            last_played_time=row["last_played_time"] or "" if "last_played_time" in keys else "",
            last_played_days=row["last_played_days"] or "" if "last_played_days" in keys else "",
            last_played_source=row["last_played_source"] or "" if "last_played_source" in keys else "",
        )


class BrowserWorker:
    def __init__(self, app: "SteamTaskAssistant") -> None:
        self.app = app
        self.commands: queue.Queue[tuple[str, tuple]] = queue.Queue()
        self.thread = threading.Thread(target=self._run, daemon=True)
        self.thread.start()

    def submit(self, command: str, *args) -> None:
        self.commands.put((command, args))

    def shutdown(self) -> None:
        self.commands.put(("__shutdown__", ()))

    def _run(self) -> None:
        controller = None
        while True:
            command, args = self.commands.get()
            try:
                if command == "__shutdown__":
                    if controller is not None:
                        controller.close()
                    controller = None
                    return
                if controller is None:
                    controller = SteamBrowserController(self.app)
                getattr(controller, command)(*args)
            except Exception as exc:
                try:
                    if controller is not None:
                        controller.close()
                except Exception:
                    pass
                controller = None
                self.app.ui_error(f"浏览器自动化失败：{exc}")
            finally:
                self.commands.task_done()


class SteamBrowserController:
    def __init__(self, app: "SteamTaskAssistant") -> None:
        from playwright.sync_api import sync_playwright

        self.app = app
        self.playwright = sync_playwright().start()
        browser_args = ["--start-maximized"]
        if os.environ.get("STEAM_ASSISTANT_USE_SYSTEM_PROXY", "").lower() not in {"1", "true", "yes"}:
            browser_args.extend(["--no-proxy-server", "--proxy-server=direct://", "--proxy-bypass-list=*"])
        self.context = self.playwright.chromium.launch_persistent_context(
            user_data_dir=str(APP_DIR / "browser_profile"),
            headless=False,
            viewport={"width": 1366, "height": 900},
            args=browser_args,
        )
        self.page = self.context.pages[0] if self.context.pages else self.context.new_page()
        self.page.set_default_timeout(20000)

    def close(self) -> None:
        try:
            self.context.close()
        finally:
            self.playwright.stop()

    def _ensure_page(self) -> None:
        if self.context.pages:
            for page in self.context.pages:
                if not page.is_closed():
                    self.page = page
                    return
        self.page = self.context.new_page()

    def _log(self, message: str) -> None:
        self.app.ui_log("浏览器", message)

    def _stopped(self) -> bool:
        return self.app.emergency_stopped or self.app.auto_stop_event.is_set()

    def _wait_auto_gate(self) -> bool:
        while self.app.auto_pause_event.is_set() and not self._stopped():
            time.sleep(0.2)
        return not self._stopped()

    def _goto(self, url: str) -> None:
        if self._stopped():
            return
        last_error = None
        for attempt in range(3):
            self._ensure_page()
            try:
                self.page.goto(url, wait_until="domcontentloaded")
                break
            except Exception as exc:
                last_error = exc
                message = str(exc)
                recoverable = (
                    "ERR_ABORTED" in message
                    or "is interrupted by another navigation" in message
                )
                if not recoverable or attempt == 2:
                    raise
                self._log("检测到页面仍在跳转，等待后重试")
                time.sleep(1.5)
        else:
            raise RuntimeError(f"无法打开页面：{last_error}")
        self._log(f"打开页面：{url}")

    def _first_visible(self, selectors: list[str], timeout_ms: int = 12000):
        self._ensure_page()
        deadline = time.time() + timeout_ms / 1000
        while time.time() < deadline and not self._stopped():
            for selector in selectors:
                locator = self.page.locator(selector).first
                try:
                    if locator.count() > 0 and locator.is_visible(timeout=500):
                        return locator
                except Exception:
                    continue
            time.sleep(0.2)
        return None

    def _fill(self, selectors: list[str], value: str, label: str) -> bool:
        locator = self._first_visible(selectors)
        if locator is None:
            self._log(f"没有找到{label}输入框")
            return False
        locator.click()
        locator.fill(value)
        self._log(f"已填入{label}")
        return True

    def _click(self, selectors: list[str], label: str) -> bool:
        locator = self._first_visible(selectors)
        if locator is None:
            self._log(f"没有找到{label}")
            return False
        locator.click()
        self._log(f"已点击{label}")
        return True

    def _fill_login_credentials(self, account: str, password: str) -> bool:
        self._ensure_page()
        try:
            password_locator = self.page.locator("input[type='password']").first
            password_locator.wait_for(state="visible", timeout=30000)
        except Exception:
            self._log("登录页密码框没有加载出来，已停止填写，避免误填到搜索框")
            return False

        password_box = password_locator.bounding_box()
        if not password_box:
            self._log("无法读取密码框位置，已停止填写")
            return False

        account_locator = None
        text_inputs = self.page.locator("input[type='text'], input:not([type]), input[type='email']")
        best_distance = 999999
        for index in range(text_inputs.count()):
            candidate = text_inputs.nth(index)
            try:
                if not candidate.is_visible(timeout=500):
                    continue
                box = candidate.bounding_box()
                if not box:
                    continue
                vertically_near = password_box["y"] - 160 <= box["y"] <= password_box["y"] + 10
                horizontally_near = abs(box["x"] - password_box["x"]) < 260
                if not vertically_near or not horizontally_near:
                    continue
                distance = abs((password_box["y"] - box["y"]) - 58) + abs(password_box["x"] - box["x"])
                if distance < best_distance:
                    best_distance = distance
                    account_locator = candidate
            except Exception:
                continue

        if account_locator is None:
            account_locator = self._first_visible(
                [
                    "div[class*='newlogindialog'] input[type='text']",
                    "form input[type='text']",
                    "input[autocomplete='username']",
                    "input[name='username']",
                ],
                timeout_ms=5000,
            )
        if account_locator is None:
            self._log("没有找到登录账号框，已停止填写，避免误填到搜索框")
            return False

        account_locator.click()
        account_locator.fill(account)
        password_locator.click()
        password_locator.fill(password)
        self._log("已在登录表单中填入账号和密码")
        return True

    def _submit_login(self) -> None:
        self._ensure_page()
        try:
            password_box = self.page.locator("input[type='password']").first.bounding_box()
        except Exception:
            password_box = None

        if password_box:
            candidates = self.page.locator(
                "button, div[role='button'], a[role='button'], "
                "[class*='Submit'], [class*='submit'], [class*='Button'], [class*='button']"
            )
            best = None
            best_score = 999999
            for index in range(candidates.count()):
                candidate = candidates.nth(index)
                try:
                    if not candidate.is_visible(timeout=300):
                        continue
                    box = candidate.bounding_box()
                    if not box:
                        continue
                    below_password = password_box["y"] + 20 <= box["y"] <= password_box["y"] + 220
                    horizontally_close = abs((box["x"] + box["width"] / 2) - (password_box["x"] + password_box["width"] / 2)) < 360
                    if not below_password or not horizontally_close:
                        continue
                    text = (candidate.inner_text(timeout=300) or "").strip().lower()
                    score = abs(box["y"] - (password_box["y"] + 95))
                    if "sign" in text or "登录" in text or "log" in text:
                        score -= 100
                    if score < best_score:
                        best_score = score
                        best = candidate
                except Exception:
                    continue
            if best is not None:
                try:
                    best.click(force=True)
                except Exception:
                    best.evaluate("el => el.click()")
                self._log("已点击密码框下方的登录按钮")
                return

        clicked = self._click(
            [
                "[class*='SubmitButton']",
                "[class*='submit']",
                "button[type='submit']",
                "button:has-text('Sign in')",
                "button:has-text('登录')",
                "div[role='button']:has-text('Sign in')",
                "div[role='button']:has-text('登录')",
            ],
            "登录按钮",
        )
        if clicked:
            return
        self.page.locator("input[type='password']").first.press("Enter")
        self._log("未找到登录按钮，已在密码框内回车提交")

    def _wait_logged_in(self) -> bool:
        deadline = time.time() + 60
        while time.time() < deadline and not self._stopped():
            try:
                if self.page.locator("#account_pulldown").count() > 0:
                    self._log("已确认登录成功：检测到右上角账号菜单")
                    return True
                if self.page.locator("a[href*='logout']").count() > 0:
                    self._log("已确认登录成功：检测到退出链接")
                    return True
            except Exception:
                pass
            try:
                if self.page.locator("input[type='password']").count() > 0:
                    self._log("仍在登录页，继续等待登录完成")
            except Exception:
                pass
            time.sleep(2)
        self._log("60秒内未检测到登录成功，已停止后续跳转，请检查是否需要手动点击登录或页面有提示")
        return False
    def _is_logged_in(self) -> bool:
        self._ensure_page()
        try:
            return (
                self.page.locator("#account_pulldown").count() > 0
                or self.page.locator("a[href*='logout']").count() > 0
            )
        except Exception:
            return False

    def ensure_logged_out_before_new_account(self) -> bool:
        self._goto("https://store.steampowered.com/")
        if self._stopped():
            return False
        time.sleep(1)
        if self._is_logged_in():
            self._log("检测到浏览器中已有登录账号，先退出当前账号")
            return self.logout_by_menu()
        else:
            self._log("未检测到已登录账号，继续登录当前任务账号")
            return True

    def start_redeem_account(self, task: Task) -> None:
        if not self.ensure_logged_out_before_new_account():
            return
        self._goto(STEAM_LOGIN_URL)
        if self._stopped():
            return
        if not self._fill_login_credentials(task.steam_account, task.steam_password):
            return
        self._submit_login()
        if not self._wait_logged_in():
            return
        self.capture_history_then_redeem(task)

    def capture_history_then_redeem(self, task: Task) -> None:
        self._goto(STEAM_HISTORY_URL)
        if self._stopped():
            return
        time.sleep(2)
        path = self.app.make_screenshot_path(task, "消费记录")
        self.page.screenshot(path=str(path), full_page=True)
        self.app.ui_screenshot_saved(task.id, path, "消费记录")
        time.sleep(1)
        self._goto(STEAM_REDEEM_URL)

    def fill_redeem_code(self, code: str) -> None:
        self._goto(STEAM_REDEEM_URL)
        self._fill(
            [
                "#wallet_code",
                "input[name='wallet_code']",
                "input[type='text']",
            ],
            code,
            "兑换码",
        )

    def confirm_redeem_code(self) -> bool:
        clicked = self._click(
            [
                "#validate_btn",
                "button[type='submit']",
                "button:has-text('Continue')",
                "button:has-text('继续')",
                "button:has-text('兑换')",
            ],
            "确认兑换按钮",
        )
        if not clicked:
            self._log("确认按钮未定位到，请人工检查页面")
        return clicked

    def _capture_page(self, task: Task, status_text: str) -> Path:
        path = self.app.make_screenshot_path(task, status_text)
        self.page.screenshot(path=str(path), full_page=True)
        self.app.ui_screenshot_saved(task.id, path, status_text)
        return path

    def _capture_history(self, task: Task, status_text: str) -> Path | None:
        self._goto(STEAM_HISTORY_URL)
        if self._stopped():
            return None
        time.sleep(2)
        return self._capture_page(task, status_text)

    def _playtime_url(self, task: Task, use_login: bool) -> str | None:
        if use_login:
            return "https://steamcommunity.com/my/games/?tab=recent"
        link = (task.profile_link or "").strip()
        if not link:
            return None
        link = link.replace("http://", "https://").rstrip("/")
        return f"{link}/games/?tab=recent"

    def _decode_js_string(self, value: str) -> str:
        try:
            return json.loads(f'"{value}"')
        except Exception:
            return value.replace("\\/", "/").replace('\\"', '"').replace("\\'", "'")

    def _parse_playtime_from_page(self) -> tuple[str, str, str, str]:
        self._ensure_page()
        try:
            self.page.wait_for_load_state("networkidle", timeout=12000)
        except Exception:
            pass
        time.sleep(2)
        content = self.page.content()
        candidates: list[tuple[int, str]] = []
        patterns = [
            r'"name"\s*:\s*"((?:\\.|[^"\\])*)".{0,1200}?"last_played"\s*:\s*(\d+)',
            r'"last_played"\s*:\s*(\d+).{0,1200}?"name"\s*:\s*"((?:\\.|[^"\\])*)"',
            r'data-name="([^"]+)".{0,1200}?data-last_played="(\d+)"',
            r'data-last_played="(\d+)".{0,1200}?data-name="([^"]+)"',
        ]
        for pattern in patterns:
            for match in re.finditer(pattern, content, flags=re.IGNORECASE | re.DOTALL):
                groups = match.groups()
                if groups[0].isdigit():
                    timestamp = int(groups[0])
                    game_name = self._decode_js_string(groups[1])
                else:
                    game_name = self._decode_js_string(groups[0])
                    timestamp = int(groups[1])
                if timestamp > 0 and game_name:
                    candidates.append((timestamp, game_name))
        if candidates:
            timestamp, game_name = max(candidates, key=lambda item: item[0])
            played_dt = datetime.fromtimestamp(timestamp)
            days = (datetime.now().date() - played_dt.date()).days
            return game_name, played_dt.strftime("%Y-%m-%d %H:%M:%S"), f"{days}天前", "Steam游戏列表"

        try:
            body_text = (self.page.locator("body").inner_text(timeout=3000) or "").strip()
        except Exception:
            body_text = ""
        lowered = body_text.lower()
        if "sign in" in lowered or "login" in lowered or "登录" in body_text:
            return "", "", "", "需要登录"
        if "private" in lowered or "私密" in body_text:
            return "", "", "", "资料私密"
        if "no games" in lowered or "没有游戏" in body_text:
            return "", "", "", "未显示游戏"
        return "", "", "", "页面未解析到last_played"

    def _finish_playtime_detection(self, task: Task, source_label: str) -> bool:
        path = self._capture_page(task, f"游玩时间检测-{source_label}")
        game, played_time, days, source = self._parse_playtime_from_page()
        if game and played_time:
            self.app.ui_playtime_result(
                task.id,
                "成功",
                game,
                played_time,
                days,
                source_label,
                f"检测成功：{game}，{played_time}，{days}",
                path,
            )
            return True
        message = f"未能确认最近游玩记录：{source}"
        self.app.ui_playtime_result(
            task.id,
            "需要人工处理",
            "",
            "",
            "",
            source_label,
            message,
            path,
        )
        return False

    def detect_playtime_public(self, task: Task) -> bool:
        url = self._playtime_url(task, use_login=False)
        if not url:
            self.app.ui_playtime_result(
                task.id,
                "需要人工处理",
                "",
                "",
                "",
                "公开页",
                "未匹配账号链接，无法公开检测",
            )
            return False
        self._goto(url)
        if self._stopped():
            return False
        return self._finish_playtime_detection(task, "公开页")

    def detect_playtime_login(self, task: Task) -> bool:
        if not self.ensure_logged_out_before_new_account():
            self.app.ui_playtime_result(task.id, "需要人工处理", "", "", "", "登录检测", "开始新账号前无法确认登出")
            return False
        self._goto(STEAM_LOGIN_URL)
        if not self._fill_login_credentials(task.steam_account, task.steam_password):
            self.app.ui_playtime_result(task.id, "需要人工处理", "", "", "", "登录检测", "登录表单填写失败")
            return False
        self._submit_login()
        if not self._wait_logged_in():
            self.app.ui_playtime_result(task.id, "需要人工处理", "", "", "", "登录检测", "登录结果未确认")
            return False
        if not self._verify_task_identity(task):
            self.app.ui_playtime_result(task.id, "需要人工处理", "", "", "", "登录检测", "账号省市/店名核心身份核对失败")
            return False
        self._goto(self._playtime_url(task, use_login=True) or "https://steamcommunity.com/my/games/?tab=recent")
        if self._stopped():
            return False
        ok = self._finish_playtime_detection(task, "登录检测")
        if not self.logout_by_menu():
            self.app.ui_playtime_result(task.id, "需要人工处理", "", "", "", "登录检测", "检测完成后账号登出未确认")
            return False
        return ok

    def full_auto_playtime(self, tasks: list[Task], use_login: bool) -> None:
        self.app.ui_playtime_batch_started(len(tasks), use_login)
        stopped = False
        for task_number, task in enumerate(tasks, start=1):
            if not self._wait_auto_gate():
                stopped = True
                break
            self.app.ui_playtime_task_started(task, task_number, len(tasks))
            if use_login:
                self.detect_playtime_login(task)
            else:
                self.detect_playtime_public(task)
            if self._stopped():
                stopped = True
                break
        self.app.ui_playtime_batch_finished(stopped)

    def _verify_task_identity(self, task: Task) -> bool:
        locator = self._first_visible(["#account_pulldown", ".pulldown.global_action_link"], timeout_ms=5000)
        if locator is None:
            self._log("无法读取右上角账号名称，不能执行全自动身份核对")
            return False
        try:
            display_name = (locator.inner_text(timeout=1000) or "").strip()
        except Exception:
            return False

        display_normalized = normalize_identity_text(display_name)
        expected_locations = [
            normalize_identity_text(location_token(task.city)),
            normalize_identity_text(location_token(task.province)),
        ]
        expected_locations = [item for item in expected_locations if len(item) >= 2]
        if any(item in display_normalized for item in expected_locations):
            self._log(f"账号身份核对通过（省市）：{display_name}")
            return True

        store_core = store_identity_core(task.store_name, task.province, task.city)
        if len(store_core) >= 4 and (
            store_core in display_normalized or display_normalized in store_core
        ):
            self._log(f"账号身份核对通过（店名核心“{store_core}”）：{display_name}")
            return True

        self._log(
            f"账号身份核对失败：页面名称“{display_name}”既未包含省市，"
            f"也未匹配店名核心“{store_core or '空'}”"
        )
        return False

    def _redeem_result_text(self) -> str:
        selectors = [
            "[role='dialog']",
            "[aria-modal='true']",
            "[class*='modal']",
            "[class*='Modal']",
            "[class*='popup']",
            "[class*='Popup']",
            "[class*='success']",
            "[class*='Success']",
            "#error_display",
            ".error_display",
            "#purchase_result_details",
            ".checkout_error",
            ".newmodal_content",
            "#wallet_code_form",
        ]
        parts = []
        for selector in selectors:
            locator = self.page.locator(selector).first
            try:
                if locator.count() > 0 and locator.is_visible(timeout=300):
                    text = (locator.inner_text(timeout=500) or "").strip()
                    if text:
                        parts.append(text)
            except Exception:
                continue
        if parts:
            return "\n".join(parts)
        try:
            return (self.page.locator("body").inner_text(timeout=1000) or "").strip()
        except Exception:
            return ""

    def _read_wallet_balance(self) -> float | None:
        selectors = [
            "#header_wallet_balance",
            "a[href*='/account/history']",
            ".accountData.price",
            "[class*='wallet_balance']",
            "[class*='WalletBalance']",
        ]
        balances: list[float] = []
        for selector in selectors:
            locators = self.page.locator(selector)
            for index in range(min(locators.count(), 8)):
                locator = locators.nth(index)
                try:
                    if not locator.is_visible(timeout=200):
                        continue
                    text = (locator.inner_text(timeout=300) or "").replace(",", "")
                    match = re.search(r"[¥￥]\s*(\d+(?:\.\d{1,2})?)", text)
                    if match:
                        balances.append(float(match.group(1)))
                except Exception:
                    continue
        return max(balances) if balances else None

    def _classify_redeem_result(self, text: str) -> tuple[str, str]:
        normalized = re.sub(r"\s+", " ", text).strip()
        lowered = normalized.lower()
        phrase_groups = [
            (
                "兑换成功",
                [
                    "successfully redeemed",
                    "successfully added",
                    "兑换成功",
                    "成功兑换",
                    "成功!",
                    "成功！",
                    "已添加至您的 steam 钱包",
                    "已添加到您的 steam 钱包",
                ],
            ),
            ("已使用", ["already been redeemed", "already redeemed", "previously redeemed", "already owns", "已经兑换", "已被兑换", "已使用", "已经在您的帐户中", "已经在你的账户中"]),
            ("无效", ["invalid code", "not valid", "无效", "不正确", "无法识别"]),
            (
                "地区受限",
                [
                    "not available in your region",
                    "cannot be redeemed in your country",
                    "currency mismatch",
                    "无法在您所在的地区兑换",
                    "不能在您的国家/地区兑换",
                    "货币不匹配",
                ],
            ),
        ]
        for status, phrases in phrase_groups:
            if any(phrase in lowered for phrase in phrases):
                return status, normalized[:500]
        return "结果不明确", normalized[:500]

    def _wait_redeem_result(
        self,
        balance_before: float | None,
        timeout_seconds: int = 25,
    ) -> tuple[str, str]:
        deadline = time.time() + timeout_seconds
        last_text = ""
        while time.time() < deadline and not self._stopped():
            time.sleep(0.15)
            last_text = self._redeem_result_text()
            status, message = self._classify_redeem_result(last_text)
            if status != "结果不明确":
                return status, message
            balance_after = self._read_wallet_balance()
            if (
                balance_before is not None
                and balance_after is not None
                and balance_after > balance_before + 0.01
            ):
                return (
                    "兑换成功",
                    f"钱包余额由¥{balance_before:.2f}增加至¥{balance_after:.2f}",
                )
        return "结果不明确", re.sub(r"\s+", " ", last_text).strip()[:500]

    def full_auto_redeem(self, tasks: list[Task], skip_map: dict[int, set[int]] | None = None) -> None:
        skip_map = skip_map or {}
        self.app.ui_auto_batch_started(len(tasks))
        for task_number, task in enumerate(tasks, start=1):
            if not self._wait_auto_gate():
                break
            self.app.ui_auto_task_started(task, task_number, len(tasks))
            if not self.ensure_logged_out_before_new_account():
                self.app.ui_auto_paused(task.id, "开始新账号前无法确认登出")
                break

            self._goto(STEAM_LOGIN_URL)
            if not self._fill_login_credentials(task.steam_account, task.steam_password):
                self.app.ui_auto_paused(task.id, "登录表单填写失败")
                break
            self._submit_login()
            if not self._wait_logged_in():
                self.app.ui_auto_paused(task.id, "登录结果未确认")
                break
            if not self._verify_task_identity(task):
                self.app.ui_auto_paused(task.id, "账号省市/店名核心身份核对失败")
                break

            self._capture_history(task, "全自动-初始消费记录")
            self._goto(STEAM_REDEEM_URL)
            task_has_exception = False
            for code_index, code in enumerate(task.voucher_codes):
                if code_index in skip_map.get(task.id, set()):
                    self._log(f"第{code_index + 1}个兑换码已有终态结果，断点续跑时跳过")
                    continue
                if not self._wait_auto_gate():
                    break
                self.app.ui_auto_code_started(task.id, code_index, code, len(task.voucher_codes))
                if not self._fill(
                    ["#wallet_code", "input[name='wallet_code']", "input[type='text']"],
                    code,
                    "兑换码",
                ):
                    self.app.ui_auto_paused(task.id, f"第{code_index + 1}个兑换码输入框未找到")
                    task_has_exception = True
                    break
                balance_before = self._read_wallet_balance()
                if not self.confirm_redeem_code():
                    self.app.ui_auto_paused(task.id, f"第{code_index + 1}个兑换码确认按钮未找到")
                    task_has_exception = True
                    break

                status, message = self._wait_redeem_result(balance_before)
                result_path = self._capture_page(
                    task,
                    f"全自动-{status}-码{code_index + 1}-{code[-5:]}",
                )
                self.app.ui_auto_code_result(
                    task.id,
                    code_index,
                    code,
                    status,
                    message,
                    result_path,
                )
                if status == "结果不明确":
                    task_has_exception = True
                    self.app.ui_auto_paused(task.id, f"第{code_index + 1}个兑换结果无法识别")
                    if not self._wait_auto_gate():
                        break
                elif status in {"无效", "地区受限"}:
                    task_has_exception = True

                self._capture_history(task, f"全自动-码{code_index + 1}-{status}-消费记录")
                self._goto(STEAM_REDEEM_URL)

            if self._stopped():
                break
            if not self.logout_by_menu():
                self.app.ui_auto_paused(task.id, "当前账号登出未确认")
                break
            final_status = "部分成功" if task_has_exception else "成功"
            self.app.ui_auto_task_finished(task.id, final_status)
            time.sleep(1)

        self.app.ui_auto_batch_finished(self._stopped())

    def full_auto_purchase(self, tasks: list[Task]) -> None:
        self.app.ui_purchase_batch_started(len(tasks))
        stopped = False
        for task_number, task in enumerate(tasks, start=1):
            if not self._wait_auto_gate():
                stopped = True
                break
            self.app.ui_purchase_batch_task_started(
                task,
                task_number,
                len(tasks),
            )
            if not self.start_purchase_account(task):
                self.app.ui_purchase_batch_paused(task.id, "账号登录或身份核对失败")
                stopped = True
                break
            if not self.clear_purchase_cart(task, require_confirmation=True):
                self.app.ui_purchase_batch_paused(task.id, "购物车清空未完成")
                stopped = True
                break

            if not self.check_purchase_license(task):
                if not self.logout_by_menu():
                    self.app.ui_purchase_batch_paused(
                        task.id,
                        "资产已存在，但账号登出未确认",
                    )
                    stopped = True
                    break
                self.app.ui_purchase_batch_already_deployed(task.id)
                continue

            if not self.add_standard_game_to_cart(task):
                self.app.ui_purchase_batch_paused(task.id, "加入购物车失败")
                stopped = True
                break
            if not self.prepare_purchase_checkout(task):
                self.app.ui_purchase_batch_paused(task.id, "进入结算页失败")
                stopped = True
                break
            if not self.app.request_purchase_confirmation(task):
                self.app.ui_purchase_batch_paused(task.id, "用户未确认本次购买")
                stopped = True
                break
            if not self.confirm_purchase_checkout(task):
                self.app.ui_purchase_batch_paused(task.id, "购买结果未确认")
                stopped = True
                break
            if not self.verify_purchase_asset_and_logout(task):
                self.app.ui_purchase_batch_paused(task.id, "资产复核或登出失败")
                stopped = True
                break
            time.sleep(1)

        self.app.ui_purchase_batch_finished(stopped or self._stopped())

    def start_activation_account(self, task: Task) -> bool:
        if not self.ensure_logged_out_before_new_account():
            return False
        self._goto(STEAM_LOGIN_URL)
        if not self._fill_login_credentials(task.steam_account, task.steam_password):
            return False
        self._submit_login()
        if not self._wait_logged_in():
            return False
        if not self._verify_task_identity(task):
            self.app.ui_activation_issue(task.id, "账号省市/店名核心身份核对失败")
            return False
        self._goto(STEAM_ACTIVATE_URL)
        self.app.ui_activation_step(task.id, "登录完成", "已进入Steam产品激活页面")
        return True

    def fill_activation_code(self, code: str) -> None:
        self._goto(STEAM_ACTIVATE_URL)
        self._fill(
            [
                "input[name='product_key']",
                "input[type='text']",
            ],
            code,
            "激活码",
        )

    def _activation_product(self, task: Task) -> dict | None:
        return resolve_game_product(task.game_name)

    def _activation_result_text(self) -> str:
        selectors = [
            "[role='dialog']",
            "[aria-modal='true']",
            "[class*='modal']",
            "[class*='Modal']",
            "[class*='popup']",
            "[class*='Popup']",
            "[class*='success']",
            "[class*='Success']",
            "#error_display",
            ".error_display",
            ".newmodal_content",
            "#registerkey_form",
            "#game_area_purchase",
        ]
        parts = []
        for selector in selectors:
            locator = self.page.locator(selector).first
            try:
                if locator.count() > 0 and locator.is_visible(timeout=300):
                    text = (locator.inner_text(timeout=500) or "").strip()
                    if text:
                        parts.append(text)
            except Exception:
                continue
        if parts:
            return "\n".join(parts)
        try:
            return (self.page.locator("body").inner_text(timeout=1000) or "").strip()
        except Exception:
            return ""

    def _classify_activation_result(self, text: str) -> tuple[str, str]:
        normalized = re.sub(r"\s+", " ", text).strip()
        lowered = normalized.lower()
        phrase_groups = [
            (
                "激活成功",
                [
                    "activation successful",
                    "successfully activated",
                    "has been activated",
                    "激活成功",
                    "成功激活",
                    "已成功激活",
                    "已添加到您的 steam 帐户",
                    "已添加至您的 steam 帐户",
                    "已添加到你的 steam 帐户",
                ],
            ),
            (
                "已拥有",
                [
                    "already in your steam library",
                    "already registered to this account",
                    "already owns",
                    "already have",
                    "已经在您的帐户中",
                    "已经在你的账户中",
                    "已在您的帐户中",
                    "已经拥有",
                    "已拥有",
                ],
            ),
            (
                "已被使用",
                [
                    "duplicate product code",
                    "already been activated by another steam account",
                    "already been redeemed",
                    "already redeemed",
                    "already used",
                    "已由另一个 steam 帐户激活",
                    "已被另一个 steam 帐户激活",
                    "已被兑换",
                    "已被激活",
                    "已使用",
                ],
            ),
            ("无效", ["invalid product code", "invalid code", "not valid", "无效", "不正确", "无法识别"]),
            (
                "地区受限",
                [
                    "not available in your region",
                    "cannot be redeemed in your country",
                    "无法在您所在的地区",
                    "不能在您的国家",
                    "地区",
                ],
            ),
        ]
        for status, phrases in phrase_groups:
            if any(phrase in lowered for phrase in phrases):
                return status, normalized[:500]
        return "结果不明确", normalized[:500]

    def _wait_activation_result(self, timeout_seconds: int = 35) -> tuple[str, str]:
        deadline = time.time() + timeout_seconds
        last_text = ""
        while time.time() < deadline and not self._stopped():
            time.sleep(0.3)
            last_text = self._activation_result_text()
            status, message = self._classify_activation_result(last_text)
            if status != "结果不明确":
                return status, message
        return "结果不明确", re.sub(r"\s+", " ", last_text).strip()[:500]

    def _click_activation_submit(self) -> bool:
        checkbox = self._first_visible(
            [
                "#accept_ssa",
                "input[name='accept_ssa']",
                "input[type='checkbox'][name*='accept']",
                "input[type='checkbox']",
            ],
            timeout_ms=2500,
        )
        if checkbox is not None:
            try:
                if not checkbox.is_checked():
                    checkbox.check()
                    self._log("已勾选Steam订户协议确认")
            except Exception:
                try:
                    checkbox.click(force=True)
                except Exception:
                    pass
        clicked = self._click(
            [
                "#register_btn",
                "button[type='submit']",
                "input[type='submit']",
                "button:has-text('继续')",
                "button:has-text('下一步')",
                "button:has-text('Continue')",
                "button:has-text('激活')",
                "button:has-text('Activate')",
                "a:has-text('继续')",
                "a:has-text('Continue')",
            ],
            "确认激活按钮",
        )
        if clicked:
            return True
        try:
            self.page.locator("input[name='product_key'], input[type='text']").first.press("Enter")
            self._log("未找到确认激活按钮，已在激活码输入框内回车提交")
            return True
        except Exception:
            return False

    def activate_current_code(self, task: Task, code_index: int, code: str) -> bool:
        self._goto(STEAM_ACTIVATE_URL)
        if not self._fill(
            ["input[name='product_key']", "#product_key", "input[type='text']"],
            code,
            "激活码",
        ):
            self.app.ui_activation_issue(task.id, f"第{code_index + 1}个激活码输入框未找到")
            return False
        if not self._click_activation_submit():
            self.app.ui_activation_issue(task.id, f"第{code_index + 1}个激活码确认按钮未找到")
            return False
        status, message = self._wait_activation_result()
        path = self._capture_page(task, f"激活码激活-{status}-码{code_index + 1}-{code[-5:]}")
        self.app.ui_activation_code_result(task.id, code_index, code, status, message, path)
        return status in {"激活成功", "已拥有"}

    def check_activation_license(self, task: Task) -> bool:
        product = self._activation_product(task)
        if product is None:
            self.app.ui_activation_step(task.id, "资产检查", "当前任务未配置可核对的游戏产品标识")
            return False
        self._goto(STEAM_LICENSES_URL)
        time.sleep(2)
        body_text = (self.page.locator("body").inner_text(timeout=5000) or "").strip()
        owned = product_in_text(product, body_text)
        if owned:
            path = self._capture_page(task, "激活码激活-资产已确认")
            self.app.ui_activation_step(
                task.id,
                "资产确认",
                f"许可证与产品激活列表中已存在《{product['official_name']}》",
                path,
            )
            return True
        self.app.ui_activation_step(
            task.id,
            "资产检查",
            f"许可证与产品激活列表中未找到《{product['official_name']}》",
        )
        return False

    def finish_activation_and_logout(self, task: Task) -> bool:
        owned = self.check_activation_license(task)
        if not owned:
            self.app.ui_activation_issue(task.id, "激活后资产清单未确认目标游戏")
            return False
        if not self.logout_by_menu():
            self.app.ui_activation_issue(task.id, "激活后登出未确认")
            return False
        self.app.ui_activation_finished(task.id)
        return True

    def full_auto_activation(self, tasks: list[Task], skip_map: dict[int, set[int]] | None = None) -> None:
        skip_map = skip_map or {}
        self.app.ui_activation_batch_started(len(tasks))
        stopped = False
        for task_number, task in enumerate(tasks, start=1):
            if not self._wait_auto_gate():
                stopped = True
                break
            self.app.ui_activation_task_started(task, task_number, len(tasks))
            if not self.start_activation_account(task):
                self.app.ui_activation_batch_paused(task.id, "账号登录或身份核对失败")
                stopped = True
                break

            if self.check_activation_license(task):
                if not self.logout_by_menu():
                    self.app.ui_activation_batch_paused(task.id, "资产已存在，但账号登出未确认")
                    stopped = True
                    break
                self.app.ui_activation_finished(task.id)
                continue

            task_has_exception = False
            for code_index, code in enumerate(task.activation_codes):
                if code_index in skip_map.get(task.id, set()):
                    self._log(f"第{code_index + 1}个激活码已有终态结果，断点续跑时跳过")
                    continue
                if not self._wait_auto_gate():
                    stopped = True
                    task_has_exception = True
                    break
                self.app.ui_activation_code_started(task.id, code_index, code, len(task.activation_codes))
                success = self.activate_current_code(task, code_index, code)
                if not success:
                    task_has_exception = True
                    self.app.ui_activation_batch_paused(task.id, f"第{code_index + 1}个激活码结果未达成成功")
                    break
                if self.check_activation_license(task):
                    break

            if stopped or self._stopped():
                break
            if task_has_exception:
                break
            if not self.logout_by_menu():
                self.app.ui_activation_batch_paused(task.id, "当前账号登出未确认")
                stopped = True
                break
            self.app.ui_activation_finished(task.id)
            time.sleep(1)

        self.app.ui_activation_batch_finished(stopped or self._stopped())

    def _purchase_product(self, task: Task) -> dict:
        product = resolve_game_product(task.game_name)
        if product is None:
            raise RuntimeError(f"未配置游戏产品标识：{task.game_name}")
        expected_price = float(task.game_price or 0)
        if expected_price and abs(expected_price - product["price"]) > 0.01:
            raise RuntimeError(
                f"任务价格¥{expected_price:.2f}与Steam标准版配置"
                f"¥{product['price']:.2f}不一致"
            )
        return product

    def _verify_live_product_catalog(self, product: dict) -> dict:
        api_url = (
            "https://store.steampowered.com/api/appdetails"
            f"?appids={product['app_id']}&cc=CN&l=schinese"
        )
        response = self.context.request.get(api_url, timeout=20000)
        if not response.ok:
            raise RuntimeError(f"Steam产品目录请求失败：HTTP {response.status}")
        payload = response.json().get(str(product["app_id"]), {})
        if not payload.get("success"):
            raise RuntimeError("Steam产品目录未返回有效产品")
        data = payload.get("data") or {}
        live_name = str(data.get("name") or "").strip()
        live_price = float((data.get("price_overview") or {}).get("final", 0)) / 100
        package_options = [
            sub
            for group in data.get("package_groups") or []
            for sub in group.get("subs") or []
        ]
        standard = next(
            (
                sub
                for sub in package_options
                if int(sub.get("packageid", 0)) == product["package_id"]
            ),
            None,
        )
        if normalize_game_name(live_name) != normalize_game_name(product["official_name"]):
            raise RuntimeError(
                f"Steam正式名称变化：预期“{product['official_name']}”，实际“{live_name}”"
            )
        if abs(live_price - product["price"]) > 0.01:
            raise RuntimeError(
                f"Steam当前价格¥{live_price:.2f}与任务价格¥{product['price']:.2f}不一致"
            )
        if standard is None:
            raise RuntimeError(
                f"Steam目录未找到标准版Package {product['package_id']}"
            )
        standard_price = float(standard.get("price_in_cents_with_discount", 0)) / 100
        if abs(standard_price - product["price"]) > 0.01:
            raise RuntimeError(
                f"标准版Package价格¥{standard_price:.2f}与任务价格不一致"
            )
        return {
            "name": live_name,
            "price": live_price,
            "package_id": int(standard["packageid"]),
            "option_text": standard.get("option_text") or "",
        }

    def _steam_session_id(self) -> str:
        for cookie in self.context.cookies("https://store.steampowered.com"):
            if cookie.get("name") == "sessionid":
                return str(cookie.get("value") or "")
        try:
            return str(self.page.evaluate("() => window.g_sessionID || ''") or "")
        except Exception:
            return ""

    def start_purchase_account(self, task: Task) -> bool:
        if not self.ensure_logged_out_before_new_account():
            return False
        self._goto(STEAM_LOGIN_URL)
        if not self._fill_login_credentials(task.steam_account, task.steam_password):
            return False
        self._submit_login()
        if not self._wait_logged_in():
            return False
        if not self._verify_task_identity(task):
            self.app.ui_purchase_issue(task.id, "账号省市/店名核心身份核对失败")
            return False
        product = self._purchase_product(task)
        self.app.ui_purchase_step(
            task.id,
            "登录完成",
            f"目标：{product['official_name']}；App {product['app_id']}；"
            f"标准版Package {product['package_id']}",
        )
        return True

    def clear_purchase_cart(self, task: Task, require_confirmation: bool = False) -> bool:
        self._purchase_product(task)
        self._goto(STEAM_CART_URL)
        time.sleep(2)
        removed = 0

        body_text = (self.page.locator("body").inner_text(timeout=5000) or "").strip()
        cart_count_match = re.search(
            r"(?:购物车|Cart)\s*[\(\[]?\s*(\d+)",
            body_text,
            flags=re.IGNORECASE,
        )
        initial_count = int(cart_count_match.group(1)) if cart_count_match else 0
        if initial_count > 0 and require_confirmation:
            if not self.app.request_cart_clear_confirmation(task, initial_count):
                self.app.ui_purchase_issue(
                    task.id,
                    f"购物车已有{initial_count}项内容，用户未确认删除",
                )
                return False

        remove_all = self._first_visible(
            [
                "text=\"移除所有项目\"",
                "text=\"Remove all items\"",
            ],
            timeout_ms=2500,
        )
        if remove_all is not None:
            remove_all.click()
            removed = initial_count
            time.sleep(1)
            confirm = self._first_visible(
                [
                    "button:has-text('确定')",
                    "button:has-text('确认')",
                    "button:has-text('移除')",
                    "button:has-text('Remove')",
                ],
                timeout_ms=1200,
            )
            if confirm is not None:
                confirm.click()
                time.sleep(1)

        remove_selectors = [
            "[data-featuretarget='removeitem']",
            "[data-featuretarget='remove-item']",
            "button[aria-label*='移除']",
            "button:has-text('移除')",
            "a:has-text('移除')",
            "text=\"移除\"",
            "button:has-text('Remove')",
            "a:has-text('Remove')",
            "text=\"Remove\"",
            ".remove_link",
        ]
        while removed < 30 and not self._stopped():
            remove_button = self._first_visible(remove_selectors, timeout_ms=1800)
            if remove_button is None:
                break
            remove_button.click()
            removed += 1
            time.sleep(1)

        self._goto(STEAM_CART_URL)
        time.sleep(2)
        body_text = (self.page.locator("body").inner_text(timeout=5000) or "").strip()
        cart_count_match = re.search(
            r"(?:购物车|Cart)\s*[\(\[]?\s*(\d+)",
            body_text,
            flags=re.IGNORECASE,
        )
        remaining_count = int(cart_count_match.group(1)) if cart_count_match else 0
        empty_markers = (
            "您的购物车是空的",
            "您的购物车为空",
            "购物车中没有商品",
            "Your cart is empty",
        )
        cart_is_empty = remaining_count == 0 and (
            any(marker.lower() in body_text.lower() for marker in empty_markers)
            or self._first_visible(remove_selectors, timeout_ms=1000) is None
        )
        if not cart_is_empty:
            self.app.ui_purchase_issue(
                task.id,
                f"购物车仍有{remaining_count or '未知数量'}项内容，未继续后续购买步骤",
            )
            return False
        path = self._capture_page(task, "游戏购买-购物车已清空")
        self.app.ui_purchase_step(
            task.id,
            "购物车清空",
            f"已移除{max(removed, initial_count)}项购物车内容，当前计数为0",
            path,
        )
        return True

    def check_purchase_license(self, task: Task) -> bool:
        product = self._purchase_product(task)
        self._goto(STEAM_LICENSES_URL)
        time.sleep(2)
        body_text = (self.page.locator("body").inner_text(timeout=5000) or "").strip()
        owned = product_in_text(product, body_text)
        if owned:
            self.app.ui_purchase_already_deployed(
                task.id,
                f"资产列表已存在《{product['official_name']}》，已标记为完成部署",
            )
            return False
        self.app.ui_purchase_step(
            task.id,
            "资产检查",
            f"许可证与产品激活列表中未找到《{product['official_name']}》",
        )
        return True

    def open_standard_game_page(self, task: Task) -> bool:
        product = self._purchase_product(task)
        live = None
        catalog_error = ""
        try:
            live = self._verify_live_product_catalog(product)
        except Exception as exc:
            catalog_error = str(exc)
        self._goto(
            f"https://store.steampowered.com/app/{product['app_id']}/"
            "?cc=CN&l=schinese"
        )
        time.sleep(2)
        if live is None:
            body_text = (
                self.page.locator("body").inner_text(timeout=8000) or ""
            ).strip()
            page_html = self.page.content()
            normalized_body = normalize_game_name(body_text)
            if normalize_game_name(product["official_name"]) not in normalized_body:
                self.app.ui_purchase_issue(
                    task.id,
                    f"商店页未识别到《{product['official_name']}》；"
                    f"目录接口错误：{catalog_error}",
                )
                return False
            price_text = body_text.replace(",", "").replace(" ", "")
            if f"{product['price']:.2f}" not in price_text:
                self.app.ui_purchase_issue(
                    task.id,
                    f"商店页未识别到标准版价格¥{product['price']:.2f}；"
                    f"目录接口错误：{catalog_error}",
                )
                return False
            package_id = str(product["package_id"])
            if package_id not in page_html:
                self.app.ui_purchase_issue(
                    task.id,
                    f"商店页未识别到标准版Package {package_id}；"
                    f"目录接口错误：{catalog_error}",
                )
                return False
            verification_message = (
                f"目录接口暂时不可用，已改用商店页现场核对："
                f"{product['official_name']}；App {product['app_id']}；"
                f"标准版Package {product['package_id']}；¥{product['price']:.2f}"
            )
        else:
            verification_message = (
                f"官方目录已核对：{live['name']}；App {product['app_id']}；"
                f"标准版Package {live['package_id']}；¥{live['price']:.2f}"
            )
        self.app.ui_purchase_step(
            task.id,
            "标准版商店页",
            verification_message,
        )
        return True

    def add_standard_game_to_cart(self, task: Task) -> bool:
        product = self._purchase_product(task)
        package_id = product["package_id"]
        try:
            session_id = self._steam_session_id()
            if not session_id:
                raise RuntimeError("未取得Steam会话ID")
            self._goto(STEAM_CART_URL)
            self.page.evaluate(
                """({ sessionId, packageId }) => {
                    const form = document.createElement("form");
                    form.method = "POST";
                    form.action = "https://store.steampowered.com/cart/";
                    const fields = {
                        snr: "1_5_9__403",
                        originating_snr: "",
                        action: "add_to_cart",
                        sessionid: sessionId,
                        subid: String(packageId),
                    };
                    for (const [name, value] of Object.entries(fields)) {
                        const input = document.createElement("input");
                        input.type = "hidden";
                        input.name = name;
                        input.value = value;
                        form.appendChild(input);
                    }
                    document.body.appendChild(form);
                    form.submit();
                }""",
                {"sessionId": session_id, "packageId": package_id},
            )
            self.page.wait_for_load_state("domcontentloaded", timeout=20000)
            self.app.ui_purchase_step(
                task.id,
                "标准版提交",
                f"已在Steam页面内直接提交标准版Package {package_id}，"
                "无需经过商店详情页",
            )
        except Exception as exc:
            self.app.ui_purchase_issue(
                task.id,
                f"标准版Package {package_id}加入购物车失败：{exc}",
            )
            return False
        try:
            self.page.wait_for_url("**/cart/**", timeout=15000)
        except Exception:
            self._goto(STEAM_CART_URL)
        time.sleep(2)
        return self.verify_purchase_cart(task)

    def verify_purchase_cart(self, task: Task) -> bool:
        product = self._purchase_product(task)
        self._goto(STEAM_CART_URL)
        time.sleep(2)
        body_text = (self.page.locator("body").inner_text(timeout=5000) or "").strip()
        if normalize_game_name(product["official_name"]) not in normalize_game_name(body_text):
            self.app.ui_purchase_issue(
                task.id,
                f"购物车中未找到《{product['official_name']}》",
            )
            return False
        if f"{product['price']:.2f}" not in body_text.replace(",", ""):
            self.app.ui_purchase_issue(
                task.id,
                f"购物车价格与标准版¥{product['price']:.2f}不一致",
            )
            return False
        remove_candidates = self.page.locator(
            "[data-featuretarget='removeitem'], "
            "button[aria-label*='移除'], .remove_link"
        )
        visible_remove_count = 0
        for index in range(min(remove_candidates.count(), 30)):
            try:
                if remove_candidates.nth(index).is_visible(timeout=200):
                    visible_remove_count += 1
            except Exception:
                continue
        header_count_match = re.search(
            r"(?:购物车|Cart)\s*[\(\[]?\s*(\d+)",
            body_text,
            flags=re.IGNORECASE,
        )
        header_count = int(header_count_match.group(1)) if header_count_match else 0
        if header_count > 1:
            self.app.ui_purchase_issue(
                task.id,
                f"购物车标题显示{header_count}项内容，必须先清空后重试",
            )
            return False
        if visible_remove_count > 1:
            self.app.ui_purchase_issue(
                task.id,
                f"购物车检测到{visible_remove_count}项内容，必须先清空后重试",
            )
            return False
        self.app.ui_purchase_step(
            task.id,
            "购物车核对",
            f"购物车仅保留《{product['official_name']}》标准版，"
            f"价格¥{product['price']:.2f}",
        )
        return True

    def prepare_purchase_checkout(self, task: Task) -> bool:
        product = self._purchase_product(task)
        if not self.verify_purchase_cart(task):
            return False
        clicked_checkout = self._click(
            [
                "#checkout_btn",
                "#btn_checkout",
                "[id*='checkout']:has-text('跳转至支付')",
                "[class*='checkout']:has-text('跳转至支付')",
                "text=\"跳转至支付\"",
                "button:has-text('跳转至支付')",
                "a:has-text('跳转至支付')",
                "[id*='checkout']:has-text('Proceed to checkout')",
                "[class*='checkout']:has-text('Proceed to checkout')",
                "text=\"Proceed to checkout\"",
                "button:has-text('Proceed to checkout')",
                "a:has-text('Proceed to checkout')",
                "#btn_purchase_self",
                "button:has-text('为自己购买')",
                "a:has-text('为自己购买')",
                "button:has-text('Purchase for myself')",
                "a:has-text('Purchase for myself')",
            ],
            "跳转至支付",
        )
        if not clicked_checkout:
            try:
                clicked_checkout = bool(
                    self.page.evaluate(
                        """() => {
                            const labels = [
                                "跳转至支付",
                                "Proceed to checkout",
                                "为自己购买",
                                "Purchase for myself",
                            ];
                            const isVisible = (el) => {
                                const style = window.getComputedStyle(el);
                                const rect = el.getBoundingClientRect();
                                return style
                                    && style.display !== "none"
                                    && style.visibility !== "hidden"
                                    && rect.width > 20
                                    && rect.height > 10;
                            };
                            const nodes = Array.from(
                                document.querySelectorAll("button,a,div,span")
                            );
                            const candidates = nodes
                                .filter((el) => isVisible(el))
                                .filter((el) => {
                                    const text = (el.innerText || el.textContent || "")
                                        .replace(/\\s+/g, " ")
                                        .trim();
                                    return labels.some((label) => text === label || text.includes(label));
                                })
                                .sort((a, b) => {
                                    const ar = a.getBoundingClientRect();
                                    const br = b.getBoundingClientRect();
                                    return (ar.width * ar.height) - (br.width * br.height);
                                });
                            if (!candidates.length) {
                                return false;
                            }
                            candidates[0].click();
                            return true;
                        }"""
                    )
                )
                if clicked_checkout:
                    self._log("已通过页面文本点击跳转至支付")
            except Exception as exc:
                self._log(f"页面文本点击跳转至支付失败：{exc}")
        if not clicked_checkout:
            self.app.ui_purchase_issue(task.id, "未找到“跳转至支付”按钮")
            return False
        time.sleep(3)
        body_text = (self.page.locator("body").inner_text(timeout=5000) or "").strip()
        if not any(
            marker.lower() in body_text.lower()
            for marker in ["Steam 钱包", "Steam钱包", "Steam Wallet"]
        ):
            self.app.ui_purchase_issue(
                task.id,
                "结算页未确认使用Steam钱包，请人工检查付款方式",
            )
            return False
        self.app.ui_purchase_step(
            task.id,
            "结算准备",
            f"已进入《{product['official_name']}》结算页并检测到Steam钱包；"
            "尚未提交购买",
        )
        return True

    def confirm_purchase_checkout(self, task: Task) -> bool:
        product = self._purchase_product(task)
        body_text = (self.page.locator("body").inner_text(timeout=5000) or "").strip()
        if not any(
            marker.lower() in body_text.lower()
            for marker in ["Steam 钱包", "Steam钱包", "Steam Wallet"]
        ):
            self.app.ui_purchase_issue(
                task.id,
                "当前页面未确认Steam钱包付款，已阻止提交",
            )
            return False
        checkbox = self._first_visible(
            [
                "#accept_ssa",
                "input[type='checkbox'][name*='agree']",
                "input[type='checkbox']",
            ],
            timeout_ms=2500,
        )
        if checkbox is not None:
            try:
                if not checkbox.is_checked():
                    checkbox.check()
                    self._log("已勾选Steam订户协议确认")
            except Exception:
                pass
        if not self._click(
            [
                "#purchase_button",
                "button:has-text('购买')",
                "a:has-text('购买')",
                "button:has-text('Purchase')",
            ],
            "最终购买按钮",
        ):
            self.app.ui_purchase_issue(task.id, "未找到最终购买按钮")
            return False
        deadline = time.time() + 40
        success = False
        while time.time() < deadline and not self._stopped():
            time.sleep(1)
            result_text = (
                self.page.locator("body").inner_text(timeout=3000) or ""
            ).strip()
            if any(
                marker.lower() in result_text.lower()
                for marker in [
                    "感谢您的购买",
                    "谢谢惠顾",
                    "购买完成",
                    "购买成功",
                    "您的购物收据",
                    "确认代码",
                    "Thank you for your purchase",
                    "Purchase complete",
                ]
            ):
                success = True
                break
        if success:
            self.app.ui_purchase_success(
                task.id,
                f"《{product['official_name']}》购买成功",
            )
            return True
        else:
            self.app.ui_purchase_issue(
                task.id,
                "提交后未识别到明确成功文本，请检查保留现场",
            )
            return False

    def verify_purchase_asset_and_logout(self, task: Task) -> bool:
        product = self._purchase_product(task)
        self._goto(STEAM_LICENSES_URL)
        time.sleep(3)
        body_text = (self.page.locator("body").inner_text(timeout=5000) or "").strip()
        owned = product_in_text(product, body_text)
        if not owned:
            self.app.ui_purchase_issue(
                task.id,
                f"购买后资产清单中未找到《{product['official_name']}》",
            )
            return False
        path = self._capture_page(task, "游戏购买-部署完成-资产已确认")
        self.app.ui_purchase_step(
            task.id,
            "部署确认",
            f"资产清单已确认存在《{product['official_name']}》",
            path,
        )
        if not self.logout_by_menu():
            self.app.ui_purchase_issue(task.id, "购买后登出未确认")
            return False
        self.app.ui_purchase_finished(task.id)
        return True

    def start_login_only(self, task: Task) -> None:
        if not self.ensure_logged_out_before_new_account():
            return
        self._goto(STEAM_LOGIN_URL)
        if not self._fill_login_credentials(task.steam_account, task.steam_password):
            return
        self._submit_login()
        self._wait_logged_in()

    def logout_by_menu(self) -> bool:
        opened = self._click(["#account_pulldown", ".pulldown.global_action_link"], "右上角账号菜单")
        time.sleep(1)
        logout = self._click(["a[href*='logout']", "text=Logout", "text=Sign out", "text=退出"], "退出账户")
        if not opened or not logout:
            self._log("未能自动完成菜单退出，请人工点击右上角账号菜单退出")
            return False
        deadline = time.time() + 20
        while time.time() < deadline and not self._stopped():
            time.sleep(0.5)
            if not self._is_logged_in():
                try:
                    self.page.wait_for_load_state("domcontentloaded", timeout=10000)
                except Exception:
                    pass
                time.sleep(1)
                self._log("已确认退出当前账号")
                return True
        self._log("点击退出后仍检测到登录状态")
        return False


class SteamTaskAssistant:
    def __init__(self, root: Tk) -> None:
        self.root = root
        self.store = TaskStore(DB_PATH)
        self.current_task_id: int | None = None
        self.current_code_index = 0
        self.screenshot_root = SCREENSHOT_DIR
        self.emergency_stopped = False
        self.auto_pause_event = threading.Event()
        self.auto_stop_event = threading.Event()
        self.auto_running = False
        self.suppress_selection_event = False
        self.filter_mode_var = StringVar(value="兑换码兑换")
        self.auto_limit_var = IntVar(value=3)
        self.auto_progress_var = StringVar(value="全自动：未启动")
        self.purchase_progress_var = StringVar(value="购买流程：未开始")
        self.activation_progress_var = StringVar(value="激活流程：未开始")
        self.playtime_progress_var = StringVar(value="游玩检测：未开始")
        self.current_var = StringVar(value="当前账号：无")
        self.next_var = StringVar(value="下一个账号：无")
        self.status_var = StringVar(value="就绪")
        self.browser_worker = BrowserWorker(self)
        self._build_ui()
        self.refresh_tasks()
        self.root.protocol("WM_DELETE_WINDOW", self.close_app)

    def _build_ui(self) -> None:
        self.root.title("Steam账号兑换/激活/购买/游玩检测助手")
        self.root.geometry("1360x820")
        self.root.minsize(1180, 720)

        toolbar = ttk.Frame(self.root, padding=8)
        toolbar.pack(side=TOP, fill=X)
        ttk.Button(toolbar, text="导入标准表格", command=self.import_excel).pack(side=LEFT, padx=4)
        ttk.Button(toolbar, text="选择截图目录", command=self.select_screenshot_dir).pack(side=LEFT, padx=4)
        ttk.Button(toolbar, text="导出结果CSV", command=self.export_csv).pack(side=LEFT, padx=4)
        ttk.Separator(toolbar, orient="vertical").pack(side=LEFT, fill=Y, padx=8)
        ttk.Label(toolbar, text="工作模式").pack(side=LEFT, padx=(0, 4))
        mode_box = ttk.Combobox(
            toolbar,
            textvariable=self.filter_mode_var,
            values=("兑换码兑换", "激活码激活", "游戏购买", "游玩时间检测", "全部模式"),
            state="readonly",
            width=12,
        )
        mode_box.pack(side=LEFT, padx=4)
        mode_box.bind("<<ComboboxSelected>>", lambda _e: self.on_mode_changed())
        ttk.Separator(toolbar, orient="vertical").pack(side=LEFT, fill=Y, padx=8)
        ttk.Button(toolbar, text="急停", command=self.emergency_stop).pack(side=LEFT, padx=4)
        ttk.Button(toolbar, text="解除急停", command=self.release_emergency_stop).pack(side=LEFT, padx=4)

        main = ttk.PanedWindow(self.root, orient="horizontal")
        main.pack(fill=BOTH, expand=True)

        left = ttk.Frame(main, padding=8)
        main.add(left, weight=2)
        self.task_tree = ttk.Treeview(
            left,
            columns=("row", "status", "store_code", "store_name", "account", "mode", "game"),
            show="headings",
            height=26,
        )
        columns = {
            "row": ("序号", 54),
            "status": ("状态", 96),
            "store_code": ("店编", 96),
            "store_name": ("店名", 240),
            "account": ("Steam账号", 160),
            "mode": ("模式", 96),
            "game": ("游戏", 120),
        }
        for key, (title, width) in columns.items():
            self.task_tree.heading(key, text=title)
            self.task_tree.column(key, width=width, anchor="w")
        self.task_tree.pack(fill=BOTH, expand=True)
        self.task_tree.bind("<<TreeviewSelect>>", self.on_task_selected)

        right = ttk.Frame(main, padding=8)
        main.add(right, weight=3)

        cards = ttk.Frame(right)
        cards.pack(fill=X)
        current_box = ttk.LabelFrame(cards, text="当前账号", padding=8)
        current_box.pack(side=LEFT, fill=BOTH, expand=True, padx=(0, 8))
        next_box = ttk.LabelFrame(cards, text="下一个账号", padding=8)
        next_box.pack(side=RIGHT, fill=BOTH, expand=True)
        ttk.Label(current_box, textvariable=self.current_var, justify=LEFT, wraplength=520).pack(anchor="w", fill=X)
        ttk.Label(next_box, textvariable=self.next_var, justify=LEFT, wraplength=420).pack(anchor="w", fill=X)

        self.full_auto_box = ttk.LabelFrame(right, text="全自动兑换（监护运行）", padding=8)
        self.full_auto_box.pack(fill=X, pady=8)
        ttk.Button(self.full_auto_box, text="开始全自动", command=self.start_full_auto_redeem).pack(side=LEFT, padx=4, pady=4)
        ttk.Button(self.full_auto_box, text="暂停", command=self.pause_full_auto).pack(side=LEFT, padx=4, pady=4)
        ttk.Button(self.full_auto_box, text="继续", command=self.resume_full_auto).pack(side=LEFT, padx=4, pady=4)
        ttk.Button(self.full_auto_box, text="停止并保留现场", command=self.stop_full_auto).pack(side=LEFT, padx=4, pady=4)
        ttk.Label(self.full_auto_box, text="本次账号数").pack(side=LEFT, padx=(16, 4))
        ttk.Spinbox(self.full_auto_box, from_=0, to=999, textvariable=self.auto_limit_var, width=5).pack(side=LEFT, padx=4)
        ttk.Label(self.full_auto_box, text="0=全部").pack(side=LEFT, padx=4)
        ttk.Label(self.full_auto_box, textvariable=self.auto_progress_var).pack(side=LEFT, padx=16)

        self.redeem_auto_box = ttk.LabelFrame(
            right,
            text="兑换/激活半自动操作",
            padding=8,
        )
        self.redeem_auto_box.pack(fill=X, pady=8)
        ttk.Button(self.redeem_auto_box, text="启动当前账号", command=self.start_current_task).pack(side=LEFT, padx=4, pady=4)
        ttk.Button(self.redeem_auto_box, text="填入当前兑换码/激活码", command=self.fill_current_code).pack(side=LEFT, padx=4, pady=4)
        ttk.Button(self.redeem_auto_box, text="确认兑换码", command=self.confirm_redeem_code).pack(side=LEFT, padx=4, pady=4)
        ttk.Button(self.redeem_auto_box, text="消费记录截图并返回兑换页", command=self.history_screenshot_back).pack(side=LEFT, padx=4, pady=4)
        ttk.Button(self.redeem_auto_box, text="退出当前账号", command=self.logout_current_account).pack(side=LEFT, padx=4, pady=4)
        ttk.Button(self.redeem_auto_box, text="完成当前账号", command=self.complete_current_account).pack(side=LEFT, padx=4, pady=4)
        ttk.Button(self.redeem_auto_box, text="标记异常", command=self.mark_exception).pack(side=LEFT, padx=4, pady=4)
        ttk.Button(self.redeem_auto_box, text="跳过当前账号", command=self.skip_current_account).pack(side=LEFT, padx=4, pady=4)

        self.activation_full_auto_box = ttk.LabelFrame(right, text="全自动激活（监护运行）", padding=8)
        ttk.Button(self.activation_full_auto_box, text="开始全自动激活", command=self.start_full_auto_activation).pack(side=LEFT, padx=4, pady=4)
        ttk.Button(self.activation_full_auto_box, text="暂停", command=self.pause_full_auto).pack(side=LEFT, padx=4, pady=4)
        ttk.Button(self.activation_full_auto_box, text="继续", command=self.resume_full_auto_activation).pack(side=LEFT, padx=4, pady=4)
        ttk.Button(self.activation_full_auto_box, text="停止并保留现场", command=self.stop_full_auto).pack(side=LEFT, padx=4, pady=4)
        ttk.Label(self.activation_full_auto_box, text="本次账号数").pack(side=LEFT, padx=(16, 4))
        ttk.Spinbox(self.activation_full_auto_box, from_=0, to=999, textvariable=self.auto_limit_var, width=5).pack(side=LEFT, padx=4)
        ttk.Label(self.activation_full_auto_box, text="0=全部").pack(side=LEFT, padx=4)
        ttk.Label(self.activation_full_auto_box, textvariable=self.activation_progress_var).pack(side=LEFT, padx=16)

        self.activation_box = ttk.LabelFrame(right, text="激活码半自动流程", padding=8)
        activation_row_one = ttk.Frame(self.activation_box)
        activation_row_one.pack(fill=X)
        ttk.Button(activation_row_one, text="1 一键登录并打开激活页", command=self.start_current_task).pack(side=LEFT, padx=4, pady=4)
        ttk.Button(activation_row_one, text="2 填入当前激活码", command=self.fill_current_code).pack(side=LEFT, padx=4, pady=4)
        ttk.Button(activation_row_one, text="3 确认激活并截图", command=self.confirm_activation_code).pack(side=LEFT, padx=4, pady=4)
        ttk.Button(activation_row_one, text="4 验证资产并登出", command=self.activation_finish_logout).pack(side=LEFT, padx=4, pady=4)
        activation_row_two = ttk.Frame(self.activation_box)
        activation_row_two.pack(fill=X)
        ttk.Button(activation_row_two, text="退出当前账号", command=self.logout_current_account).pack(side=LEFT, padx=4, pady=4)
        ttk.Button(activation_row_two, text="完成当前账号", command=self.complete_current_account).pack(side=LEFT, padx=4, pady=4)
        ttk.Button(activation_row_two, text="标记异常", command=self.mark_exception).pack(side=LEFT, padx=4, pady=4)
        ttk.Button(activation_row_two, text="跳过当前账号", command=self.skip_current_account).pack(side=LEFT, padx=4, pady=4)
        ttk.Label(activation_row_two, textvariable=self.activation_progress_var).pack(side=LEFT, padx=16)

        self.purchase_box = ttk.LabelFrame(right, text="游戏购买半自动流程", padding=8)
        purchase_row_one = ttk.Frame(self.purchase_box)
        purchase_row_one.pack(fill=X)
        ttk.Button(purchase_row_one, text="1 一键登录", command=self.purchase_login).pack(side=LEFT, padx=4, pady=4)
        ttk.Button(purchase_row_one, text="2 清空购物车", command=self.purchase_clear_cart).pack(side=LEFT, padx=4, pady=4)
        ttk.Button(purchase_row_one, text="3 检查数字资产", command=self.purchase_check_license).pack(side=LEFT, padx=4, pady=4)
        ttk.Button(purchase_row_one, text="4 提交游戏到购物车", command=self.purchase_add_cart).pack(side=LEFT, padx=4, pady=4)
        ttk.Button(purchase_row_one, text="5 跳转支付流程", command=self.purchase_prepare_checkout).pack(side=LEFT, padx=4, pady=4)
        purchase_row_two = ttk.Frame(self.purchase_box)
        purchase_row_two.pack(fill=X)
        ttk.Button(purchase_row_two, text="6 购买并截图", command=self.purchase_confirm_checkout).pack(side=LEFT, padx=4, pady=4)
        ttk.Button(purchase_row_two, text="7 验证资产并登出", command=self.purchase_finish_logout).pack(side=LEFT, padx=4, pady=4)
        ttk.Button(purchase_row_two, text="标记异常", command=self.mark_exception).pack(side=LEFT, padx=4, pady=4)
        ttk.Button(purchase_row_two, text="跳过当前账号", command=self.skip_current_account).pack(side=LEFT, padx=4, pady=4)
        ttk.Label(purchase_row_two, textvariable=self.purchase_progress_var).pack(side=LEFT, padx=16)
        purchase_row_three = ttk.Frame(self.purchase_box)
        purchase_row_three.pack(fill=X)
        ttk.Button(purchase_row_three, text="批量购买到确认点", command=self.start_full_auto_purchase).pack(side=LEFT, padx=4, pady=4)
        ttk.Button(purchase_row_three, text="暂停", command=self.pause_full_auto).pack(side=LEFT, padx=4, pady=4)
        ttk.Button(purchase_row_three, text="继续", command=self.resume_full_auto_purchase).pack(side=LEFT, padx=4, pady=4)
        ttk.Button(purchase_row_three, text="停止并保留现场", command=self.stop_full_auto).pack(side=LEFT, padx=4, pady=4)
        ttk.Label(purchase_row_three, text="本次账号数").pack(side=LEFT, padx=(16, 4))
        ttk.Spinbox(purchase_row_three, from_=0, to=999, textvariable=self.auto_limit_var, width=5).pack(side=LEFT, padx=4)
        self.purchase_box.pack(fill=X, pady=8)

        self.playtime_box = ttk.LabelFrame(right, text="游玩时间检测", padding=8)
        playtime_row_one = ttk.Frame(self.playtime_box)
        playtime_row_one.pack(fill=X)
        ttk.Button(playtime_row_one, text="导入/刷新索尼克检测任务", command=self.import_playtime_tasks).pack(side=LEFT, padx=4, pady=4)
        ttk.Button(playtime_row_one, text="公开页检测当前账号", command=self.detect_current_playtime_public).pack(side=LEFT, padx=4, pady=4)
        ttk.Button(playtime_row_one, text="登录检测当前账号", command=self.detect_current_playtime_login).pack(side=LEFT, padx=4, pady=4)
        ttk.Button(playtime_row_one, text="批量登录检测", command=self.start_full_auto_playtime_login).pack(side=LEFT, padx=4, pady=4)
        ttk.Button(playtime_row_one, text="批量公开检测", command=self.start_full_auto_playtime_public).pack(side=LEFT, padx=4, pady=4)
        playtime_row_two = ttk.Frame(self.playtime_box)
        playtime_row_two.pack(fill=X)
        ttk.Button(playtime_row_two, text="暂停", command=self.pause_full_auto).pack(side=LEFT, padx=4, pady=4)
        ttk.Button(playtime_row_two, text="继续", command=self.resume_full_auto_playtime).pack(side=LEFT, padx=4, pady=4)
        ttk.Button(playtime_row_two, text="停止并保留现场", command=self.stop_full_auto).pack(side=LEFT, padx=4, pady=4)
        ttk.Label(playtime_row_two, text="本次账号数").pack(side=LEFT, padx=(16, 4))
        ttk.Spinbox(playtime_row_two, from_=0, to=999, textvariable=self.auto_limit_var, width=5).pack(side=LEFT, padx=4)
        ttk.Label(playtime_row_two, text="0=全部").pack(side=LEFT, padx=4)
        ttk.Label(playtime_row_two, textvariable=self.playtime_progress_var).pack(side=LEFT, padx=16)

        self.manual_box = ttk.LabelFrame(right, text="应急手动复制（仅在半自动失败时使用）", padding=8)
        self.manual_box.pack(fill=X, pady=8)
        ttk.Button(self.manual_box, text="复制账号", command=self.copy_account).pack(side=LEFT, padx=4, pady=4)
        ttk.Button(self.manual_box, text="复制密码", command=self.copy_password).pack(side=LEFT, padx=4, pady=4)
        ttk.Button(self.manual_box, text="复制当前码", command=self.copy_current_code).pack(side=LEFT, padx=4, pady=4)

        self.code_box = ttk.LabelFrame(right, text="当前任务号码", padding=8)
        self.code_box.pack(fill=BOTH, expand=True, pady=8)
        controls = ttk.Frame(self.code_box)
        controls.pack(fill=X)
        ttk.Button(controls, text="上一个码", command=lambda: self.shift_code(-1)).pack(side=LEFT, padx=4)
        ttk.Button(controls, text="下一个码", command=lambda: self.shift_code(1)).pack(side=LEFT, padx=4)
        self.code_tree = ttk.Treeview(self.code_box, columns=("idx", "kind", "code", "result"), show="headings", height=8)
        self.code_tree.heading("idx", text="序号")
        self.code_tree.heading("kind", text="类型")
        self.code_tree.heading("code", text="号码")
        self.code_tree.heading("result", text="自动结果")
        self.code_tree.column("idx", width=60, anchor="center")
        self.code_tree.column("kind", width=100, anchor="center")
        self.code_tree.column("code", width=320, anchor="w")
        self.code_tree.column("result", width=120, anchor="center")
        self.code_tree.pack(fill=BOTH, expand=True, pady=(6, 0))
        self.code_tree.bind("<<TreeviewSelect>>", self.on_code_selected)

        self.log_box = ttk.LabelFrame(right, text="日志", padding=8)
        self.log_box.pack(fill=BOTH, expand=True)
        self.log_tree = ttk.Treeview(self.log_box, columns=("time", "type", "message"), show="headings", height=8)
        self.log_tree.heading("time", text="时间")
        self.log_tree.heading("type", text="类型")
        self.log_tree.heading("message", text="内容")
        self.log_tree.column("time", width=150, anchor="w")
        self.log_tree.column("type", width=90, anchor="w")
        self.log_tree.column("message", width=820, anchor="w")
        self.log_tree.pack(fill=BOTH, expand=True)

        ttk.Label(self.root, textvariable=self.status_var, relief="sunken", anchor="w", padding=4).pack(side=BOTTOM, fill=X)
        self.update_mode_controls()

    def import_excel(self) -> None:
        path = filedialog.askopenfilename(title="选择标准任务表格", filetypes=[("Excel 工作簿", "*.xlsx"), ("所有文件", "*.*")])
        if not path:
            return
        try:
            count = self.store.import_excel(Path(path), replace=True)
        except Exception as exc:
            messagebox.showerror("导入失败", str(exc))
            return
        self.current_task_id = None
        self.refresh_tasks()
        self.status_var.set(f"已导入 {count} 条任务")

    def select_screenshot_dir(self) -> None:
        path = filedialog.askdirectory(title="选择截图保存目录")
        if path:
            self.screenshot_root = Path(path)
            self.screenshot_root.mkdir(parents=True, exist_ok=True)
            self.status_var.set(f"截图目录：{self.screenshot_root}")

    def export_csv(self) -> None:
        path = filedialog.asksaveasfilename(title="导出结果CSV", defaultextension=".csv", filetypes=[("CSV", "*.csv")])
        if not path:
            return
        self.store.export_csv(Path(path))
        self.status_var.set(f"已导出：{path}")

    def import_playtime_tasks(self) -> None:
        try:
            count = self.store.import_playtime_tasks(APP_DIR.parent, replace=True)
        except Exception as exc:
            messagebox.showerror("导入游玩检测任务失败", str(exc))
            return
        self.filter_mode_var.set("游玩时间检测")
        self.current_task_id = None
        self.update_mode_controls()
        self.refresh_tasks()
        self.status_var.set(f"已导入/刷新 {count} 个索尼克赛车交叉世界账号检测任务")

    def current_playtime_task(self) -> Task | None:
        task = self.current_task()
        if task is None:
            return None
        if task.mode != "游玩时间检测":
            messagebox.showwarning("模式不正确", "请先切换到“游玩时间检测”模式。")
            return None
        return task

    def detect_current_playtime_public(self) -> None:
        if not self.ensure_not_stopped():
            return
        task = self.current_playtime_task()
        if task is None:
            return
        self.store.update_status(task.id, "处理中")
        self.refresh_tasks()
        self.browser_worker.submit("detect_playtime_public", task)

    def detect_current_playtime_login(self) -> None:
        if not self.ensure_not_stopped():
            return
        task = self.current_playtime_task()
        if task is None:
            return
        self.store.update_status(task.id, "处理中")
        self.refresh_tasks()
        self.browser_worker.submit("detect_playtime_login", task)

    def start_full_auto_playtime_public(self) -> None:
        self.start_full_auto_playtime(use_login=False)

    def start_full_auto_playtime_login(self) -> None:
        self.start_full_auto_playtime(use_login=True)

    def start_full_auto_playtime(self, use_login: bool) -> None:
        if not self.ensure_not_stopped() or self.auto_running:
            return
        if self.filter_mode_var.get() != "游玩时间检测":
            messagebox.showwarning("模式不正确", "游玩检测只会在“游玩时间检测”模式中运行。")
            return
        limit = max(0, int(self.auto_limit_var.get() or 0))
        tasks = self.store.pending_playtime_tasks(self.current_task_id, limit)
        if not tasks:
            messagebox.showinfo("没有任务", "从当前账号开始没有待检测的游玩时间任务。")
            return
        mode_name = "登录检测" if use_login else "公开页检测"
        if not messagebox.askyesno(
            f"开始批量{mode_name}",
            f"将从当前账号开始连续检测 {len(tasks)} 个账号。\n"
            "程序只读取上次游玩的游戏和时间，不会执行购买、兑换或激活操作。\n\n是否开始？",
        ):
            return
        self.auto_pause_event.clear()
        self.auto_stop_event.clear()
        self.auto_running = True
        self.browser_worker.submit("full_auto_playtime", tasks, use_login)

    def resume_full_auto_playtime(self) -> None:
        if self.emergency_stopped:
            return
        if self.auto_running:
            self.auto_pause_event.clear()
            self.playtime_progress_var.set("游玩检测：继续运行")
            self.log("游玩检测", "已继续运行")
            return
        if self.filter_mode_var.get() != "游玩时间检测":
            messagebox.showwarning("模式不正确", "继续检测只适用于“游玩时间检测”模式。")
            return
        self.start_full_auto_playtime(use_login=True)

    def refresh_tasks(self) -> None:
        for item in self.task_tree.get_children():
            self.task_tree.delete(item)
        tasks = self.store.filtered_tasks(self.filter_mode_var.get())
        for task in tasks:
            game_display = task.last_played_game if task.mode == "游玩时间检测" and task.last_played_game else task.game_name
            if task.mode == "游玩时间检测" and task.last_played_days:
                game_display = f"{game_display}（{task.last_played_days}）" if game_display else task.last_played_days
            self.task_tree.insert(
                "",
                END,
                iid=str(task.id),
                values=(task.row_no, task.status, task.store_code, task.store_name, task.steam_account, task.mode, game_display),
            )
        visible = {task.id for task in tasks}
        if self.current_task_id not in visible:
            first = self.store.next_task_after(None, self.filter_mode_var.get())
            self.set_current(first.id if first else None)
        else:
            self.set_current(self.current_task_id)

    def on_mode_changed(self) -> None:
        self.current_task_id = None
        self.update_mode_controls()
        self.refresh_tasks()
        self.status_var.set(f"已切换工作模式：{self.filter_mode_var.get()}")

    def update_mode_controls(self) -> None:
        mode = self.filter_mode_var.get()
        self.full_auto_box.pack_forget()
        self.redeem_auto_box.pack_forget()
        self.activation_full_auto_box.pack_forget()
        self.activation_box.pack_forget()
        self.purchase_box.pack_forget()
        self.playtime_box.pack_forget()
        self.code_box.pack_forget()

        if mode == "兑换码兑换":
            self.full_auto_box.pack(fill=X, pady=8, before=self.manual_box)
            self.redeem_auto_box.pack(fill=X, pady=8, before=self.manual_box)
            self.code_box.pack(fill=BOTH, expand=True, pady=8, before=self.log_box)
        elif mode == "激活码激活":
            self.activation_full_auto_box.pack(fill=X, pady=8, before=self.manual_box)
            self.activation_box.pack(fill=X, pady=8, before=self.manual_box)
            self.code_box.pack(fill=BOTH, expand=True, pady=8, before=self.log_box)
        elif mode == "游戏购买":
            self.purchase_box.pack(fill=X, pady=8, before=self.manual_box)
        elif mode == "游玩时间检测":
            self.playtime_box.pack(fill=X, pady=8, before=self.manual_box)
        else:
            self.status_var.set("请选择具体工作模式后再执行自动化操作")

    def on_task_selected(self, _event: object) -> None:
        if self.suppress_selection_event:
            return
        selection = self.task_tree.selection()
        if selection:
            self.set_current(int(selection[0]))

    def set_current(self, task_id: int | None) -> None:
        self.current_task_id = task_id
        self.current_code_index = 0
        task = self.current_task()
        if task is None:
            self.current_var.set("当前账号：无")
            self.next_var.set("下一个账号：无")
            self.fill_codes(None)
            self.fill_logs(None)
            return
        self.current_var.set(self.format_task(task, show_password=True))
        next_task = self.store.next_task_after(task.id, self.filter_mode_var.get())
        self.next_var.set(self.format_task(next_task, show_password=False) if next_task else "下一个账号：无")
        self.fill_codes(task)
        self.fill_logs(task.id)
        target = str(task.id)
        if target in self.task_tree.get_children() and self.task_tree.selection() != (target,):
            self.suppress_selection_event = True
            try:
                self.task_tree.selection_set(target)
                self.task_tree.focus(target)
            finally:
                self.suppress_selection_event = False

    def current_task(self) -> Task | None:
        return self.store.get_task(self.current_task_id)

    def format_task(self, task: Task | None, show_password: bool) -> str:
        if task is None:
            return "无"
        password = task.steam_password if show_password else "********"
        product = resolve_game_product(task.game_name)
        product_line = ""
        if task.mode == "游戏购买" and product is not None:
            product_line = (
                f"Steam正式名称：{product['official_name']}\n"
                f"App ID：{product['app_id']}    标准版Package：{product['package_id']}\n"
            )
        playtime_line = ""
        if task.mode == "游玩时间检测":
            playtime_line = (
                f"账号链接：{task.profile_link or '未匹配'}\n"
                f"上次游玩游戏：{task.last_played_game or '未检测'}\n"
                f"上次游玩时间：{task.last_played_time or '未检测'}\n"
                f"距离今天：{task.last_played_days or '未检测'}\n"
                f"检测来源：{task.last_played_source or '未检测'}\n"
            )
        return (
            f"店编：{task.store_code}\n"
            f"店名：{task.store_name}\n"
            f"省份/城市：{task.province} / {task.city}\n"
            f"Steam账号：{task.steam_account}\n"
            f"Steam密码：{password}\n"
            f"模式：{task.mode}\n"
            f"购买游戏：{task.game_name} {task.game_price}\n"
            f"{product_line}"
            f"{playtime_line}"
            f"兑换码数量：{len(task.voucher_codes)}    激活码数量：{len(task.activation_codes)}\n"
            f"状态：{task.status}\n"
            f"备注：{task.note}"
        )

    def fill_codes(self, task: Task | None) -> None:
        for item in self.code_tree.get_children():
            self.code_tree.delete(item)
        if task is None:
            return
        results = self.store.voucher_results(task.id) if task.mode in {"兑换码兑换", "激活码激活"} else {}
        if task.mode == "兑换码兑换":
            rows = [("兑换码", code) for code in task.voucher_codes]
        elif task.mode == "激活码激活":
            rows = [("激活码", code) for code in task.activation_codes]
        elif task.mode == "游玩时间检测":
            rows = [("账号链接", task.profile_link or "未匹配账号链接")]
        else:
            rows = [("购买游戏", task.game_name)] if task.game_name else []
        for idx, (kind, code) in enumerate(rows, start=1):
            result = results.get(idx - 1)
            result_status = result["status"] if result else "待处理"
            self.code_tree.insert("", END, iid=str(idx - 1), values=(idx, kind, code, result_status))
        if rows:
            self.code_tree.selection_set("0")
            self.code_tree.focus("0")

    def fill_logs(self, task_id: int | None) -> None:
        for item in self.log_tree.get_children():
            self.log_tree.delete(item)
        for event in reversed(self.store.events(task_id)):
            self.log_tree.insert("", END, values=(event["event_time"], event["event_type"], event["message"]))

    def current_codes(self) -> list[str]:
        task = self.current_task()
        if task is None:
            return []
        if task.mode == "兑换码兑换":
            return task.voucher_codes
        if task.mode == "激活码激活":
            return task.activation_codes
        return [task.game_name] if task.game_name else []

    def on_code_selected(self, _event: object) -> None:
        selection = self.code_tree.selection()
        if selection:
            self.current_code_index = int(selection[0])

    def shift_code(self, delta: int) -> None:
        codes = self.current_codes()
        if not codes:
            return
        self.current_code_index = max(0, min(len(codes) - 1, self.current_code_index + delta))
        iid = str(self.current_code_index)
        self.code_tree.selection_set(iid)
        self.code_tree.focus(iid)

    def ensure_not_stopped(self) -> bool:
        if self.emergency_stopped:
            messagebox.showwarning("急停已启用", "当前处于急停状态。请先解除急停再继续。")
            return False
        return True

    def emergency_stop(self) -> None:
        self.emergency_stopped = True
        self.auto_stop_event.set()
        self.log("急停", "用户触发急停，已阻止后续自动化动作")

    def release_emergency_stop(self) -> None:
        self.emergency_stopped = False
        self.auto_stop_event.clear()
        self.log("急停", "用户解除急停")

    def start_full_auto_redeem(self) -> None:
        if not self.ensure_not_stopped() or self.auto_running:
            return
        if self.filter_mode_var.get() != "兑换码兑换":
            messagebox.showwarning("模式不正确", "全自动兑换只会在“兑换码兑换”模式中运行。")
            return
        limit = max(0, int(self.auto_limit_var.get() or 0))
        tasks = self.store.pending_redeem_tasks(self.current_task_id, limit)
        if not tasks:
            messagebox.showinfo("没有任务", "从当前账号开始没有待处理的兑换任务。")
            return
        if not messagebox.askyesno(
            "开始全自动兑换",
            f"将从当前账号开始连续处理 {len(tasks)} 个账号。\n"
            "登录身份不符、结果不明确或登出失败时会自动暂停。\n\n是否开始？",
        ):
            return
        self.auto_pause_event.clear()
        self.auto_stop_event.clear()
        self.auto_running = True
        self.browser_worker.submit("full_auto_redeem", tasks, self.auto_skip_map(tasks))

    def auto_skip_map(self, tasks: list[Task]) -> dict[int, set[int]]:
        terminal_statuses = {"兑换成功", "已使用", "激活成功", "已拥有", "已被使用", "无效", "地区受限"}
        return {
            task.id: {
                code_index
                for code_index, row in self.store.voucher_results(task.id).items()
                if row["status"] in terminal_statuses
            }
            for task in tasks
        }

    def start_full_auto_activation(self) -> None:
        if not self.ensure_not_stopped() or self.auto_running:
            return
        if self.filter_mode_var.get() != "激活码激活":
            messagebox.showwarning("模式不正确", "全自动激活只会在“激活码激活”模式中运行。")
            return
        limit = max(0, int(self.auto_limit_var.get() or 0))
        tasks = self.store.pending_activation_tasks(self.current_task_id, limit)
        if not tasks:
            messagebox.showinfo("没有任务", "从当前账号开始没有待处理的激活任务。")
            return
        if not messagebox.askyesno(
            "开始全自动激活",
            f"将从当前账号开始连续处理 {len(tasks)} 个账号。\n"
            "程序会自动登录、核对身份、提交激活码、截图成功结果、验证资产并登出。\n\n"
            "激活码会真实消耗；结果不明确、无效、已被其他账号使用或登出失败时会自动暂停。\n\n是否开始？",
            default=messagebox.NO,
        ):
            return
        self.auto_pause_event.clear()
        self.auto_stop_event.clear()
        self.auto_running = True
        self.browser_worker.submit("full_auto_activation", tasks, self.auto_skip_map(tasks))

    def pause_full_auto(self) -> None:
        if self.auto_running:
            self.auto_pause_event.set()
            if self.filter_mode_var.get() == "游戏购买":
                self.purchase_progress_var.set("购买流程：批量已暂停")
                self.log("批量购买", "已请求暂停，将在当前安全步骤停下")
            elif self.filter_mode_var.get() == "激活码激活":
                self.activation_progress_var.set("激活流程：批量已暂停")
                self.log("全自动激活", "已请求暂停，将在当前安全步骤停下")
            elif self.filter_mode_var.get() == "游玩时间检测":
                self.playtime_progress_var.set("游玩检测：批量已暂停")
                self.log("游玩检测", "已请求暂停，将在当前安全步骤停下")
            else:
                self.auto_progress_var.set("全自动：已暂停")
                self.log("全自动", "已请求暂停，将在当前安全步骤停下")

    def resume_full_auto(self) -> None:
        if self.emergency_stopped:
            return
        if self.auto_running:
            self.auto_pause_event.clear()
            self.auto_progress_var.set("全自动：继续运行")
            self.log("全自动", "已继续运行")
            return
        if self.filter_mode_var.get() != "兑换码兑换":
            messagebox.showwarning("模式不正确", "断点继续只适用于“兑换码兑换”模式。")
            return
        tasks = self.store.pending_redeem_tasks(self.current_task_id, 1)
        if not tasks:
            messagebox.showinfo("没有任务", "从当前账号开始没有可继续的兑换任务。")
            return
        self.auto_pause_event.clear()
        self.auto_stop_event.clear()
        self.auto_running = True
        self.auto_progress_var.set("全自动：从断点重新启动")
        self.log("全自动", "批次已结束，现从数据库断点重新启动")
        self.browser_worker.submit("full_auto_redeem", tasks, self.auto_skip_map(tasks))

    def stop_full_auto(self) -> None:
        if self.auto_running:
            self.auto_stop_event.set()
            self.auto_pause_event.clear()
            if self.filter_mode_var.get() == "游戏购买":
                self.purchase_progress_var.set("购买流程：批量正在停止")
                self.log("批量购买", "已请求停止，浏览器现场将保留")
            elif self.filter_mode_var.get() == "激活码激活":
                self.activation_progress_var.set("激活流程：批量正在停止")
                self.log("全自动激活", "已请求停止，浏览器现场将保留")
            elif self.filter_mode_var.get() == "游玩时间检测":
                self.playtime_progress_var.set("游玩检测：批量正在停止")
                self.log("游玩检测", "已请求停止，浏览器现场将保留")
            else:
                self.auto_progress_var.set("全自动：正在停止")
                self.log("全自动", "已请求停止，浏览器现场将保留")

    def close_app(self) -> None:
        self.auto_stop_event.set()
        self.auto_pause_event.clear()
        self.browser_worker.shutdown()
        self.root.after(700, self.root.destroy)

    def copy_account(self) -> None:
        task = self.current_task()
        if task:
            copy_to_clipboard(self.root, task.steam_account)
            self.log("应急复制", "已复制账号")

    def copy_password(self) -> None:
        task = self.current_task()
        if task:
            copy_to_clipboard(self.root, task.steam_password)
            self.log("应急复制", "已复制密码")

    def copy_current_code(self) -> None:
        codes = self.current_codes()
        if not codes:
            return
        copy_to_clipboard(self.root, codes[self.current_code_index])
        self.log("应急复制", f"已复制当前码：第 {self.current_code_index + 1} 个")

    def current_purchase_task(self) -> Task | None:
        if not self.ensure_not_stopped():
            return None
        task = self.current_task()
        if task is None:
            return None
        if task.mode != "游戏购买":
            messagebox.showwarning("模式不正确", "当前任务不是游戏购买任务。")
            return None
        product = resolve_game_product(task.game_name)
        if product is None:
            messagebox.showerror("游戏未配置", f"无法识别游戏：{task.game_name}")
            return None
        return task

    def request_cart_clear_confirmation(self, task: Task, item_count: int) -> bool:
        event = threading.Event()
        result = {"confirmed": False}

        def ask() -> None:
            result["confirmed"] = messagebox.askyesno(
                "确认清空购物车",
                f"即将删除当前Steam购物车内的 {item_count} 项内容：\n\n"
                f"店编：{task.store_code}\n"
                f"店名：{task.store_name}\n"
                f"账号：{task.steam_account}\n\n"
                "点击“是”后程序将清空购物车并继续批量购买流程，是否继续？",
                default=messagebox.NO,
            )
            if result["confirmed"]:
                self.store.add_event(
                    task.id,
                    "清空确认",
                    f"用户确认删除购物车内{item_count}项内容",
                )
            event.set()

        self.root.after(0, ask)
        while not event.wait(0.2):
            if self.emergency_stopped or self.auto_stop_event.is_set():
                return False
        return bool(result["confirmed"])

    def request_purchase_confirmation(self, task: Task) -> bool:
        product = resolve_game_product(task.game_name)
        if product is None:
            return False
        event = threading.Event()
        result = {"confirmed": False}

        def ask() -> None:
            result["confirmed"] = messagebox.askyesno(
                "确认本次购买",
                f"即将使用Steam钱包完成真实购买：\n\n"
                f"店编：{task.store_code}\n"
                f"店名：{task.store_name}\n"
                f"账号：{task.steam_account}\n"
                f"游戏：{product['official_name']} 标准版\n"
                f"价格：¥{product['price']:.2f}\n\n"
                "点击“是”后程序将提交最终购买，是否继续？",
                default=messagebox.NO,
            )
            if result["confirmed"]:
                self.store.add_event(
                    task.id,
                    "支付确认",
                    "用户确认使用Steam钱包提交最终购买",
                )
            event.set()

        self.root.after(0, ask)
        while not event.wait(0.2):
            if self.emergency_stopped or self.auto_stop_event.is_set():
                return False
        return bool(result["confirmed"])

    def submit_purchase_command(self, command: str) -> None:
        task = self.current_purchase_task()
        if task is None:
            return
        self.store.update_status(task.id, "处理中")
        self.refresh_tasks()
        self.browser_worker.submit(command, task)

    def purchase_login(self) -> None:
        self.submit_purchase_command("start_purchase_account")

    def purchase_clear_cart(self) -> None:
        self.submit_purchase_command("clear_purchase_cart")

    def purchase_check_license(self) -> None:
        self.submit_purchase_command("check_purchase_license")

    def purchase_open_standard(self) -> None:
        self.submit_purchase_command("open_standard_game_page")

    def purchase_add_cart(self) -> None:
        self.submit_purchase_command("add_standard_game_to_cart")

    def purchase_verify_cart(self) -> None:
        self.submit_purchase_command("verify_purchase_cart")

    def purchase_prepare_checkout(self) -> None:
        self.submit_purchase_command("prepare_purchase_checkout")

    def purchase_confirm_checkout(self) -> None:
        task = self.current_purchase_task()
        if task is None:
            return
        product = resolve_game_product(task.game_name)
        if product is None:
            return
        confirmed = messagebox.askyesno(
            "确认支付",
            f"即将使用Steam钱包完成真实购买：\n\n"
            f"店编：{task.store_code}\n"
            f"账号：{task.steam_account}\n"
            f"游戏：{product['official_name']} 标准版\n"
            f"价格：¥{product['price']:.2f}\n\n"
            "点击“是”后程序将提交最终购买，是否继续？",
        )
        if not confirmed:
            self.log("游戏购买", "用户取消最终支付")
            return
        self.store.add_event(task.id, "支付确认", "用户确认使用Steam钱包提交最终购买")
        self.browser_worker.submit("confirm_purchase_checkout", task)

    def purchase_finish_logout(self) -> None:
        self.submit_purchase_command("verify_purchase_asset_and_logout")

    def start_full_auto_purchase(self) -> None:
        if not self.ensure_not_stopped() or self.auto_running:
            return
        if self.filter_mode_var.get() != "游戏购买":
            messagebox.showwarning("模式不正确", "批量购买只会在“游戏购买”模式中运行。")
            return
        limit = max(0, int(self.auto_limit_var.get() or 0))
        tasks = self.store.pending_purchase_tasks(self.current_task_id, limit)
        if not tasks:
            messagebox.showinfo("没有任务", "从当前账号开始没有待处理的购买任务。")
            return
        if not messagebox.askyesno(
            "开始批量购买",
            f"将从当前账号开始处理 {len(tasks)} 个购买任务。\n\n"
            "程序会自动登录、清空购物车、检查资产、加入购物车并进入结算页；"
            "这些准备步骤不会再弹出中间确认。\n\n"
            "到达每个账号的最终扣款页时，会显示该店编、店名、账号、游戏和价格，"
            "并只在这一笔真实购买前确认一次。\n\n"
            "此处仅确认启动批量准备流程，是否开始？",
            default=messagebox.NO,
        ):
            return
        self.auto_pause_event.clear()
        self.auto_stop_event.clear()
        self.auto_running = True
        self.browser_worker.submit("full_auto_purchase", tasks)

    def resume_full_auto_purchase(self) -> None:
        if self.emergency_stopped:
            return
        if self.auto_running:
            self.auto_pause_event.clear()
            self.purchase_progress_var.set("购买流程：批量继续运行")
            self.log("批量购买", "已继续运行")
            return
        if self.filter_mode_var.get() != "游戏购买":
            messagebox.showwarning("模式不正确", "批量购买继续只适用于“游戏购买”模式。")
            return
        self.start_full_auto_purchase()

    def resume_full_auto_activation(self) -> None:
        if self.emergency_stopped:
            return
        if self.auto_running:
            self.auto_pause_event.clear()
            self.activation_progress_var.set("激活流程：批量继续运行")
            self.log("全自动激活", "已继续运行")
            return
        if self.filter_mode_var.get() != "激活码激活":
            messagebox.showwarning("模式不正确", "批量激活继续只适用于“激活码激活”模式。")
            return
        self.start_full_auto_activation()

    def start_current_task(self) -> None:
        if not self.ensure_not_stopped():
            return
        task = self.current_task()
        if task is None:
            return
        self.store.update_status(task.id, "处理中")
        self.refresh_tasks()
        if task.mode == "兑换码兑换":
            self.browser_worker.submit("start_redeem_account", task)
        elif task.mode == "激活码激活":
            self.browser_worker.submit("start_activation_account", task)
        elif task.mode == "游戏购买":
            self.browser_worker.submit("start_purchase_account", task)
        elif task.mode == "游玩时间检测":
            self.browser_worker.submit("detect_playtime_login", task)
        self.log("任务", "已启动当前账号半自动流程")

    def fill_current_code(self) -> None:
        if not self.ensure_not_stopped():
            return
        task = self.current_task()
        codes = self.current_codes()
        if task is None or not codes:
            return
        code = codes[self.current_code_index]
        if task.mode == "激活码激活":
            self.browser_worker.submit("fill_activation_code", code)
        else:
            self.browser_worker.submit("fill_redeem_code", code)

    def confirm_redeem_code(self) -> None:
        if self.ensure_not_stopped():
            self.browser_worker.submit("confirm_redeem_code")

    def confirm_activation_code(self) -> None:
        if not self.ensure_not_stopped():
            return
        task = self.current_task()
        codes = self.current_codes()
        if task is None or task.mode != "激活码激活" or not codes:
            return
        code = codes[self.current_code_index]
        self.store.update_status(task.id, "处理中")
        self.refresh_tasks()
        self.browser_worker.submit("activate_current_code", task, self.current_code_index, code)

    def activation_finish_logout(self) -> None:
        if not self.ensure_not_stopped():
            return
        task = self.current_task()
        if task is None or task.mode != "激活码激活":
            return
        self.browser_worker.submit("finish_activation_and_logout", task)

    def history_screenshot_back(self) -> None:
        if not self.ensure_not_stopped():
            return
        task = self.current_task()
        if task:
            self.browser_worker.submit("capture_history_then_redeem", task)

    def logout_current_account(self) -> None:
        if self.ensure_not_stopped():
            self.browser_worker.submit("logout_by_menu")

    def complete_current_account(self) -> None:
        task = self.current_task()
        if task:
            self.store.update_status(task.id, "成功")
            self.log("状态", "当前账号已标记成功")
            self.move_to_next()

    def mark_exception(self) -> None:
        task = self.current_task()
        if task:
            self.store.update_status(task.id, "需要人工处理")
            self.log("状态", "当前账号已标记需要人工处理")
            self.refresh_tasks()

    def skip_current_account(self) -> None:
        task = self.current_task()
        if task:
            self.store.update_status(task.id, "跳过")
            self.log("状态", "当前账号已跳过")
            self.move_to_next()

    def move_to_next(self) -> None:
        task = self.current_task()
        next_task = self.store.next_task_after(task.id if task else None, self.filter_mode_var.get())
        self.set_current(next_task.id if next_task else None)

    def make_screenshot_path(self, task: Task, status_text: str) -> Path:
        folder = self.screenshot_root / "-".join(
            [clean_filename_part(task.store_code), clean_filename_part(task.store_name), clean_filename_part(task.steam_account)]
        )
        folder.mkdir(parents=True, exist_ok=True)
        timestamp = time.strftime("%Y%m%d-%H%M%S")
        filename = "-".join(
            [
                clean_filename_part(task.store_code),
                clean_filename_part(task.store_name),
                clean_filename_part(task.steam_account),
                clean_filename_part(status_text),
                timestamp,
            ]
        ) + ".png"
        return folder / filename

    def ui_log(self, event_type: str, message: str) -> None:
        self.root.after(0, lambda: self.log(event_type, message))

    def ui_error(self, message: str) -> None:
        timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
        with AUTOMATION_LOG_PATH.open("a", encoding="utf-8") as f:
            f.write(f"[{timestamp}] {message}\n")
        def update() -> None:
            if self.auto_running:
                self.auto_running = False
                self.auto_pause_event.clear()
                self.auto_stop_event.clear()
                if self.filter_mode_var.get() == "游戏购买":
                    self.purchase_progress_var.set("购买流程：自动化异常，已停止")
                elif self.filter_mode_var.get() == "激活码激活":
                    self.activation_progress_var.set("激活流程：自动化异常，已停止")
                elif self.filter_mode_var.get() == "游玩时间检测":
                    self.playtime_progress_var.set("游玩检测：自动化异常，已停止")
                else:
                    self.auto_progress_var.set("全自动：自动化异常，已停止")
            messagebox.showerror("自动化错误", message)

        self.root.after(0, update)

    def ui_screenshot_saved(self, task_id: int, path: Path, status_text: str) -> None:
        def update() -> None:
            self.store.update_screenshot_dir(task_id, path.parent)
            self.store.add_event(task_id, "截图", f"{status_text}：{path}", str(path))
            self.fill_logs(task_id)
            self.status_var.set(f"已截图：{path}")

        self.root.after(0, update)

    def ui_purchase_step(
        self,
        task_id: int,
        step: str,
        message: str,
        screenshot_path: Path | None = None,
    ) -> None:
        def update() -> None:
            self.store.add_event(
                task_id,
                "游戏购买",
                f"{step}：{message}",
                str(screenshot_path or ""),
            )
            self.purchase_progress_var.set(f"购买流程：{step}")
            self.status_var.set(message)
            if self.current_task_id == task_id:
                self.fill_logs(task_id)

        self.root.after(0, update)

    def ui_purchase_issue(
        self,
        task_id: int,
        message: str,
        screenshot_path: Path | None = None,
    ) -> None:
        def update() -> None:
            self.store.update_status(task_id, "需要人工处理")
            self.store.add_event(
                task_id,
                "购买暂停",
                message,
                str(screenshot_path or ""),
            )
            self.purchase_progress_var.set(f"购买流程：已暂停，{message}")
            self.status_var.set(message)
            self.refresh_tasks()

        self.root.after(0, update)

    def ui_purchase_success(
        self,
        task_id: int,
        message: str,
    ) -> None:
        def update() -> None:
            self.store.update_status(task_id, "处理中")
            self.store.add_event(task_id, "购买成功", message)
            self.purchase_progress_var.set("购买流程：购买成功，等待登出")
            self.status_var.set(message)
            self.refresh_tasks()

        self.root.after(0, update)

    def ui_purchase_already_deployed(
        self,
        task_id: int,
        message: str,
    ) -> None:
        def update() -> None:
            self.store.update_status(task_id, "成功")
            self.store.add_event(
                task_id,
                "游戏购买",
                message,
            )
            self.purchase_progress_var.set("购买流程：资产已存在，完成部署")
            self.status_var.set(message)
            self.refresh_tasks()

        self.root.after(0, update)

    def ui_purchase_finished(self, task_id: int) -> None:
        def update() -> None:
            self.store.update_status(task_id, "成功")
            self.store.add_event(task_id, "游戏购买", "购买完成并确认登出")
            self.purchase_progress_var.set("购买流程：完成并已登出")
            if self.current_task_id == task_id:
                next_task = self.store.next_task_after(task_id, "游戏购买")
                self.set_current(next_task.id if next_task else None)
            self.refresh_tasks()

        self.root.after(0, update)

    def ui_purchase_batch_started(self, task_count: int) -> None:
        def update() -> None:
            self.auto_running = True
            self.purchase_progress_var.set(f"购买流程：批量准备处理 {task_count} 个账号")
            self.status_var.set("批量购买已启动")

        self.root.after(0, update)

    def ui_purchase_batch_task_started(
        self,
        task: Task,
        task_number: int,
        task_count: int,
    ) -> None:
        def update() -> None:
            self.set_current(task.id)
            self.store.update_status(task.id, "处理中")
            self.store.add_event(
                task.id,
                "批量购买",
                f"开始账号 {task_number}/{task_count}",
            )
            self.purchase_progress_var.set(
                f"购买流程：批量 {task_number}/{task_count}，{task.store_code}"
            )
            self.refresh_tasks()

        self.root.after(0, update)

    def ui_purchase_batch_paused(self, task_id: int, reason: str) -> None:
        self.auto_pause_event.set()

        def update() -> None:
            self.store.update_status(task_id, "需要人工处理")
            self.store.add_event(task_id, "批量购买暂停", reason)
            self.purchase_progress_var.set(f"购买流程：批量已暂停，{reason}")
            self.status_var.set(reason)
            self.refresh_tasks()

        self.root.after(0, update)

    def ui_purchase_batch_already_deployed(self, task_id: int) -> None:
        def update() -> None:
            self.store.add_event(task_id, "批量购买", "资产已存在，跳过购买并已登出")
            self.refresh_tasks()

        self.root.after(0, update)

    def ui_purchase_batch_finished(self, stopped: bool) -> None:
        def update() -> None:
            self.auto_running = False
            self.auto_pause_event.clear()
            self.auto_stop_event.clear()
            text = "购买流程：批量已停止" if stopped else "购买流程：批量完成"
            self.purchase_progress_var.set(text)
            self.status_var.set(text)
            self.refresh_tasks()

        self.root.after(0, update)

    def ui_playtime_batch_started(self, task_count: int, use_login: bool) -> None:
        def update() -> None:
            self.auto_running = True
            mode_name = "登录检测" if use_login else "公开页检测"
            self.playtime_progress_var.set(f"游玩检测：{mode_name}准备处理 {task_count} 个账号")
            self.status_var.set("批量游玩检测已启动")

        self.root.after(0, update)

    def ui_playtime_task_started(self, task: Task, task_number: int, task_count: int) -> None:
        def update() -> None:
            self.set_current(task.id)
            self.store.update_status(task.id, "处理中")
            self.store.add_event(task.id, "游玩检测", f"开始账号 {task_number}/{task_count}")
            self.playtime_progress_var.set(
                f"游玩检测：账号 {task_number}/{task_count}，{task.store_code}"
            )
            self.refresh_tasks()

        self.root.after(0, update)

    def ui_playtime_result(
        self,
        task_id: int,
        status: str,
        game: str,
        played_time: str,
        days: str,
        source: str,
        message: str,
        screenshot_path: Path | None = None,
    ) -> None:
        def update() -> None:
            self.store.update_playtime_result(
                task_id,
                status,
                game,
                played_time,
                days,
                source,
                message,
                str(screenshot_path or ""),
            )
            self.store.add_event(task_id, "游玩检测", message, str(screenshot_path or ""))
            self.playtime_progress_var.set(f"游玩检测：{message}")
            self.status_var.set(message)
            self.refresh_tasks()
            if self.current_task_id == task_id:
                self.fill_logs(task_id)

        self.root.after(0, update)

    def ui_playtime_batch_finished(self, stopped: bool) -> None:
        def update() -> None:
            self.auto_running = False
            self.auto_pause_event.clear()
            self.auto_stop_event.clear()
            text = "游玩检测：批量已停止" if stopped else "游玩检测：批量完成"
            self.playtime_progress_var.set(text)
            self.status_var.set(text)
            self.refresh_tasks()

        self.root.after(0, update)

    def ui_activation_step(
        self,
        task_id: int,
        step: str,
        message: str,
        screenshot_path: Path | None = None,
    ) -> None:
        def update() -> None:
            self.store.add_event(
                task_id,
                "激活码激活",
                f"{step}：{message}",
                str(screenshot_path or ""),
            )
            self.activation_progress_var.set(f"激活流程：{step}")
            self.status_var.set(message)
            if self.current_task_id == task_id:
                self.fill_logs(task_id)

        self.root.after(0, update)

    def ui_activation_issue(
        self,
        task_id: int,
        message: str,
        screenshot_path: Path | None = None,
    ) -> None:
        def update() -> None:
            self.store.update_status(task_id, "需要人工处理")
            self.store.add_event(
                task_id,
                "激活暂停",
                message,
                str(screenshot_path or ""),
            )
            self.activation_progress_var.set(f"激活流程：已暂停，{message}")
            self.status_var.set(message)
            self.refresh_tasks()

        self.root.after(0, update)

    def ui_activation_code_started(self, task_id: int, code_index: int, code: str, code_count: int) -> None:
        def update() -> None:
            self.current_code_index = code_index
            self.store.upsert_voucher_result(task_id, code_index, code, "激活中")
            self.store.add_event(task_id, "全自动激活", f"开始激活码 {code_index + 1}/{code_count}")
            if self.current_task_id == task_id:
                self.fill_codes(self.current_task())
                iid = str(code_index)
                if iid in self.code_tree.get_children():
                    self.code_tree.selection_set(iid)
                    self.code_tree.focus(iid)
            self.activation_progress_var.set(f"激活流程：激活码 {code_index + 1}/{code_count}")

        self.root.after(0, update)

    def ui_activation_code_result(
        self,
        task_id: int,
        code_index: int,
        code: str,
        status: str,
        message: str,
        screenshot_path: Path,
    ) -> None:
        def update() -> None:
            self.store.upsert_voucher_result(
                task_id,
                code_index,
                code,
                status,
                message,
                str(screenshot_path),
            )
            self.store.add_event(
                task_id,
                "激活结果",
                f"第{code_index + 1}个：{status}；{message[:180]}",
                str(screenshot_path),
            )
            if self.current_task_id == task_id:
                self.fill_codes(self.current_task())
                self.fill_logs(task_id)
            self.activation_progress_var.set(f"激活流程：第{code_index + 1}个码 {status}")

        self.root.after(0, update)

    def ui_activation_task_started(self, task: Task, task_number: int, task_count: int) -> None:
        def update() -> None:
            self.set_current(task.id)
            self.store.update_status(task.id, "处理中")
            self.store.add_event(task.id, "全自动激活", f"开始账号 {task_number}/{task_count}")
            self.activation_progress_var.set(
                f"激活流程：账号 {task_number}/{task_count}，{task.store_code}"
            )
            self.refresh_tasks()

        self.root.after(0, update)

    def ui_activation_batch_started(self, task_count: int) -> None:
        def update() -> None:
            self.auto_running = True
            self.activation_progress_var.set(f"激活流程：准备处理 {task_count} 个账号")
            self.status_var.set("全自动激活已启动")

        self.root.after(0, update)

    def ui_activation_batch_paused(self, task_id: int, reason: str) -> None:
        self.auto_pause_event.set()

        def update() -> None:
            self.store.update_status(task_id, "需要人工处理")
            self.store.add_event(task_id, "全自动激活暂停", reason)
            self.activation_progress_var.set(f"激活流程：批量已暂停，{reason}")
            self.status_var.set(reason)
            self.refresh_tasks()

        self.root.after(0, update)

    def ui_activation_finished(self, task_id: int) -> None:
        def update() -> None:
            self.store.update_status(task_id, "成功")
            self.store.add_event(task_id, "激活码激活", "激活完成并确认登出")
            self.activation_progress_var.set("激活流程：完成并已登出")
            if self.current_task_id == task_id:
                next_task = self.store.next_task_after(task_id, "激活码激活")
                self.set_current(next_task.id if next_task else None)
            self.refresh_tasks()

        self.root.after(0, update)

    def ui_activation_batch_finished(self, stopped: bool) -> None:
        def update() -> None:
            self.auto_running = False
            self.auto_pause_event.clear()
            self.auto_stop_event.clear()
            text = "激活流程：批量已停止" if stopped else "激活流程：批量完成"
            self.activation_progress_var.set(text)
            self.status_var.set(text)
            self.refresh_tasks()

        self.root.after(0, update)

    def ui_auto_batch_started(self, task_count: int) -> None:
        def update() -> None:
            self.auto_running = True
            self.auto_progress_var.set(f"全自动：准备处理 {task_count} 个账号")
            self.status_var.set("全自动兑换已启动")

        self.root.after(0, update)

    def ui_auto_task_started(self, task: Task, task_number: int, task_count: int) -> None:
        def update() -> None:
            self.set_current(task.id)
            self.store.update_status(task.id, "处理中")
            self.store.add_event(task.id, "全自动", f"开始账号 {task_number}/{task_count}")
            self.auto_progress_var.set(
                f"全自动：账号 {task_number}/{task_count}，{task.store_code}"
            )
            self.refresh_tasks()

        self.root.after(0, update)

    def ui_auto_code_started(self, task_id: int, code_index: int, code: str, code_count: int) -> None:
        def update() -> None:
            self.current_code_index = code_index
            self.store.upsert_voucher_result(task_id, code_index, code, "兑换中")
            self.store.add_event(task_id, "全自动", f"开始兑换码 {code_index + 1}/{code_count}")
            if self.current_task_id == task_id:
                self.fill_codes(self.current_task())
                iid = str(code_index)
                if iid in self.code_tree.get_children():
                    self.code_tree.selection_set(iid)
                    self.code_tree.focus(iid)
            self.auto_progress_var.set(f"全自动：兑换码 {code_index + 1}/{code_count}")

        self.root.after(0, update)

    def ui_auto_code_result(
        self,
        task_id: int,
        code_index: int,
        code: str,
        status: str,
        message: str,
        screenshot_path: Path,
    ) -> None:
        def update() -> None:
            self.store.upsert_voucher_result(
                task_id,
                code_index,
                code,
                status,
                message,
                str(screenshot_path),
            )
            self.store.add_event(
                task_id,
                "兑换结果",
                f"第{code_index + 1}个：{status}；{message[:180]}",
                str(screenshot_path),
            )
            if self.current_task_id == task_id:
                self.fill_codes(self.current_task())
                self.fill_logs(task_id)
            self.auto_progress_var.set(f"全自动：第{code_index + 1}个码 {status}")

        self.root.after(0, update)

    def ui_auto_paused(self, task_id: int, reason: str) -> None:
        self.auto_pause_event.set()

        def update() -> None:
            self.store.update_status(task_id, "需要人工处理")
            self.store.add_event(task_id, "全自动暂停", reason)
            self.auto_progress_var.set(f"全自动：已暂停，{reason}")
            self.status_var.set(reason)
            self.refresh_tasks()

        self.root.after(0, update)

    def ui_auto_task_finished(self, task_id: int, status: str) -> None:
        def update() -> None:
            self.store.update_status(task_id, status)
            self.store.add_event(task_id, "全自动", f"账号处理完成：{status}")
            self.refresh_tasks()

        self.root.after(0, update)

    def ui_auto_batch_finished(self, stopped: bool) -> None:
        def update() -> None:
            self.auto_running = False
            self.auto_pause_event.clear()
            self.auto_stop_event.clear()
            text = "全自动：已停止" if stopped else "全自动：批次完成"
            self.auto_progress_var.set(text)
            self.status_var.set(text)
            self.refresh_tasks()

        self.root.after(0, update)

    def log(self, event_type: str, message: str) -> None:
        task = self.current_task()
        if task:
            self.store.add_event(task.id, event_type, message)
            self.fill_logs(task.id)
        self.status_var.set(message)


def main() -> None:
    ensure_dirs()
    root = Tk()
    try:
        ttk.Style().theme_use("clam")
    except Exception:
        pass
    SteamTaskAssistant(root)
    root.mainloop()


if __name__ == "__main__":
    main()
