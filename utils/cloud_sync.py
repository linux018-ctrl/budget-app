"""
雲端同步模組 - 管理 Google Sheets 連結設定 & 同步狀態
"""
import os
import json
from datetime import datetime
from typing import Optional

# 設定檔路徑（與 app 同目錄）
CONFIG_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CONFIG_FILE = os.path.join(CONFIG_DIR, "cloud_config.json")


def load_cloud_config() -> dict:
    """
    載入雲端設定
    回傳: {
        "google_sheets_url": "https://...",
        "last_sync_time": "2026-06-15 10:30:00",
        "auto_refresh": True
    }
    """
    default = {
        "google_sheets_url": "",
        "last_sync_time": "",
        "auto_refresh": False
    }
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                saved = json.load(f)
                default.update(saved)
        except (json.JSONDecodeError, IOError):
            pass
    return default


def save_cloud_config(config: dict):
    """儲存雲端設定到 config 檔"""
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
    except IOError:
        pass


def update_last_sync_time():
    """更新最後同步時間"""
    config = load_cloud_config()
    config["last_sync_time"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    save_cloud_config(config)
    return config["last_sync_time"]


def validate_google_sheets_url(url: str) -> tuple:
    """
    驗證 Google Sheets URL 是否有效
    回傳: (is_valid: bool, message: str)
    """
    url = url.strip()
    if not url:
        return False, "請輸入 Google Sheets URL"

    if "docs.google.com/spreadsheets" not in url:
        return False, "URL 格式不正確，請貼上 Google Sheets 的連結"

    return True, "✅ URL 格式正確"
