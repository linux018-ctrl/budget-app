"""
資料管理模組 - 處理預算與記帳資料的 CRUD 操作
"""
import json
import os
import uuid
from datetime import datetime, date
from typing import Optional


DATA_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), "data")
BUDGET_FILE = os.path.join(DATA_DIR, "budget.json")
RECORDS_FILE = os.path.join(DATA_DIR, "records.json")

# 預設支出類別
DEFAULT_EXPENSE_CATEGORIES = [
    "🍜 飲食", "🏠 居住", "🚗 交通", "📱 通訊",
    "🎮 娛樂", "👕 服飾", "💊 醫療", "📚 教育",
    "🛒 日用品", "💝 人情", "🐾 寵物", "📦 其他支出"
]

# 預設收入類別
DEFAULT_INCOME_CATEGORIES = [
    "💼 薪資", "💰 獎金", "📈 投資收入", "🏠 租金收入",
    "🎁 副業收入", "📦 其他收入"
]


def _ensure_data_dir():
    """確保 data 目錄存在"""
    os.makedirs(DATA_DIR, exist_ok=True)


def _load_json(filepath: str) -> dict:
    """載入 JSON 檔案"""
    _ensure_data_dir()
    if os.path.exists(filepath):
        with open(filepath, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def _save_json(filepath: str, data: dict):
    """儲存 JSON 檔案"""
    _ensure_data_dir()
    with open(filepath, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


# ─── 預算管理 ───────────────────────────────────────────

def get_budget(year: int, month: int) -> dict:
    """
    取得某年月的預算設定
    回傳格式:
    {
        "income_target": 50000,
        "categories": {
            "🍜 飲食": 8000,
            "🏠 居住": 15000,
            ...
        }
    }
    """
    budgets = _load_json(BUDGET_FILE)
    key = f"{year}-{month:02d}"
    if key in budgets:
        return budgets[key]
    return {"income_target": 0, "categories": {}}


def save_budget(year: int, month: int, income_target: float, categories: dict):
    """儲存某年月的預算設定"""
    budgets = _load_json(BUDGET_FILE)
    key = f"{year}-{month:02d}"
    budgets[key] = {
        "income_target": income_target,
        "categories": categories
    }
    _save_json(BUDGET_FILE, budgets)


def copy_budget_from(src_year: int, src_month: int, dst_year: int, dst_month: int):
    """從某月複製預算到另一個月"""
    budget = get_budget(src_year, src_month)
    if budget["income_target"] > 0 or budget["categories"]:
        save_budget(dst_year, dst_month, budget["income_target"], budget["categories"])
        return True
    return False


def get_all_budget_months() -> list:
    """取得所有已設定預算的月份"""
    budgets = _load_json(BUDGET_FILE)
    return sorted(budgets.keys())


# ─── 記帳紀錄管理 ─────────────────────────────────────────

def get_records(year: int, month: int) -> list:
    """
    取得某年月的所有記帳紀錄
    每筆紀錄格式:
    {
        "id": "uuid",
        "date": "2026-03-15",
        "type": "expense" | "income",
        "category": "🍜 飲食",
        "amount": 150,
        "note": "午餐 - 牛肉麵"
    }
    """
    all_records = _load_json(RECORDS_FILE)
    key = f"{year}-{month:02d}"
    return all_records.get(key, [])


def add_record(year: int, month: int, record_date: str, record_type: str,
               category: str, amount: float, note: str = "") -> str:
    """新增一筆記帳紀錄，回傳紀錄 ID"""
    all_records = _load_json(RECORDS_FILE)
    key = f"{year}-{month:02d}"
    if key not in all_records:
        all_records[key] = []

    record_id = str(uuid.uuid4())[:8]
    all_records[key].append({
        "id": record_id,
        "date": record_date,
        "type": record_type,
        "category": category,
        "amount": amount,
        "note": note,
        "created_at": datetime.now().isoformat()
    })
    _save_json(RECORDS_FILE, all_records)
    return record_id


def update_record(year: int, month: int, record_id: str,
                  record_date: str, record_type: str,
                  category: str, amount: float, note: str = ""):
    """更新一筆記帳紀錄"""
    all_records = _load_json(RECORDS_FILE)
    key = f"{year}-{month:02d}"
    records = all_records.get(key, [])
    for rec in records:
        if rec["id"] == record_id:
            rec["date"] = record_date
            rec["type"] = record_type
            rec["category"] = category
            rec["amount"] = amount
            rec["note"] = note
            rec["updated_at"] = datetime.now().isoformat()
            break
    _save_json(RECORDS_FILE, all_records)


def delete_record(year: int, month: int, record_id: str):
    """刪除一筆記帳紀錄"""
    all_records = _load_json(RECORDS_FILE)
    key = f"{year}-{month:02d}"
    records = all_records.get(key, [])
    all_records[key] = [r for r in records if r["id"] != record_id]
    _save_json(RECORDS_FILE, all_records)


def get_all_record_months() -> list:
    """取得所有有紀錄的月份"""
    all_records = _load_json(RECORDS_FILE)
    return sorted(all_records.keys())


# ─── 分析計算 ─────────────────────────────────────────────

def get_monthly_summary(year: int, month: int) -> dict:
    """
    取得某年月的收支摘要
    回傳:
    {
        "total_income": 總收入,
        "total_expense": 總支出,
        "balance": 結餘（收入-支出）,
        "budget_total": 預算總額,
        "budget_remaining": 預算剩餘,
        "is_over_budget": 是否透支,
        "expense_by_category": { 類別: 金額 },
        "income_by_category": { 類別: 金額 },
        "budget_vs_actual": { 類別: {"budget": 預算, "actual": 實際, "diff": 差額} },
        "daily_expenses": { 日期: 金額 }
    }
    """
    records = get_records(year, month)
    budget = get_budget(year, month)

    total_income = sum(r["amount"] for r in records if r["type"] == "income")
    total_expense = sum(r["amount"] for r in records if r["type"] == "expense")

    # 各類別支出統計
    expense_by_cat = {}
    for r in records:
        if r["type"] == "expense":
            cat = r["category"]
            expense_by_cat[cat] = expense_by_cat.get(cat, 0) + r["amount"]

    # 各類別收入統計
    income_by_cat = {}
    for r in records:
        if r["type"] == "income":
            cat = r["category"]
            income_by_cat[cat] = income_by_cat.get(cat, 0) + r["amount"]

    # 預算 vs 實際
    budget_total = sum(budget["categories"].values()) if budget["categories"] else 0
    budget_vs_actual = {}
    for cat, budgeted in budget.get("categories", {}).items():
        actual = expense_by_cat.get(cat, 0)
        budget_vs_actual[cat] = {
            "budget": budgeted,
            "actual": actual,
            "diff": budgeted - actual,
            "usage_pct": (actual / budgeted * 100) if budgeted > 0 else 0
        }

    # 每日支出統計
    daily_expenses = {}
    for r in records:
        if r["type"] == "expense":
            d = r["date"]
            daily_expenses[d] = daily_expenses.get(d, 0) + r["amount"]

    return {
        "total_income": total_income,
        "total_expense": total_expense,
        "balance": total_income - total_expense,
        "income_target": budget.get("income_target", 0),
        "budget_total": budget_total,
        "budget_remaining": budget_total - total_expense,
        "is_over_budget": total_expense > budget_total if budget_total > 0 else False,
        "expense_by_category": expense_by_cat,
        "income_by_category": income_by_cat,
        "budget_vs_actual": budget_vs_actual,
        "daily_expenses": daily_expenses,
        "record_count": len(records)
    }


def get_yearly_summary(year: int) -> list:
    """取得整年各月的收支摘要"""
    summaries = []
    for month in range(1, 13):
        summary = get_monthly_summary(year, month)
        summary["month"] = month
        summaries.append(summary)
    return summaries
