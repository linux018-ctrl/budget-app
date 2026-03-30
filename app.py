"""
💰 Alan's 每月收支管理 App
使用 Streamlit 建構，整合自訂預算表與 cwmoney 記帳紀錄
"""
import streamlit as st
import pandas as pd
from datetime import datetime, date
import calendar

import json
import os

from utils.excel_importer import (
    load_budget_from_excel, load_cwmoney_records,
    get_cwmoney_monthly_summary, get_project_icons, get_main_category_icons,
    find_cwmoney_files, get_latest_cwmoney_file,
    load_cwmoney_from_uploaded_file, load_cwmoney_from_google_sheets,
    load_budget_from_uploaded_file, parse_cwmoney_dataframe,
    BUDGET_EXCEL_PATH, DATA_DIR, IS_CLOUD
)
from utils.charts import (
    create_budget_vs_actual_chart, create_expense_pie_chart,
    create_daily_expense_chart, create_budget_usage_gauges,
    create_expense_type_pie, create_sub_category_treemap,
    create_yearly_income_expense_chart, create_yearly_savings_chart,
    create_yearly_cumulative_chart
)
from utils.cloud_sync import (
    load_cloud_config, save_cloud_config,
    update_last_sync_time, validate_google_sheets_url
)
from utils.drive_sync import get_latest_csv_dataframe, get_latest_xml_dataframe

# ─── 頁面設定 ──────────────────────────────────────────────
st.set_page_config(
    page_title="💰 Alan's 收支管理",
    page_icon="💰",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ─── 自訂 CSS ──────────────────────────────────────────────
st.markdown("""
<style>
    .metric-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 20px;
        border-radius: 12px;
        color: white;
        text-align: center;
        box-shadow: 0 4px 15px rgba(0,0,0,0.1);
    }
    .metric-card.income {
        background: linear-gradient(135deg, #11998e 0%, #38ef7d 100%);
    }
    .metric-card.expense {
        background: linear-gradient(135deg, #eb3349 0%, #f45c43 100%);
    }
    .metric-card.savings {
        background: linear-gradient(135deg, #4facfe 0%, #00f2fe 100%);
    }
    .metric-card.balance-positive {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    }
    .metric-card.balance-negative {
        background: linear-gradient(135deg, #c0392b 0%, #e74c3c 100%);
    }
    .metric-card h3 { margin: 0; font-size: 14px; opacity: 0.9; }
    .metric-card h1 { margin: 5px 0 0 0; font-size: 26px; }
    /* ── 專案層級：圓環卡片 ── */
    .project-card {
        background: #ffffff;
        border-radius: 14px;
        padding: 18px 14px;
        text-align: center;
        box-shadow: 0 2px 10px rgba(0,0,0,0.07);
        border: 1px solid #eee;
        margin-bottom: 8px;
    }
    .project-card .ring {
        width: 90px; height: 90px;
        border-radius: 50%;
        margin: 0 auto 10px;
        display: flex; align-items: center; justify-content: center;
        font-size: 22px; font-weight: bold; color: white;
    }
    .project-card .title { font-size: 15px; font-weight: 600; margin-bottom: 4px; }
    .project-card .detail { font-size: 12px; color: #636e72; }
    .project-card .detail .over { color: #e74c3c; font-weight: 600; }
    .project-card .detail .under { color: #27ae60; font-weight: 600; }

    /* ── 主分類層級：表格行 ── */
    .main-cat-table { width: 100%; border-collapse: separate; border-spacing: 0 6px; }
    .main-cat-table th {
        text-align: left; font-size: 12px; color: #636e72;
        padding: 4px 8px; border-bottom: 2px solid #dfe6e9;
    }
    .main-cat-table td { padding: 8px; font-size: 13px; vertical-align: middle; }
    .main-cat-row {
        background: #f8f9fa;
        border-radius: 8px;
    }
    .main-cat-row td:first-child { border-radius: 8px 0 0 8px; font-weight: 600; }
    .main-cat-row td:last-child { border-radius: 0 8px 8px 0; }
    .mini-bar {
        background: #ecf0f1;
        border-radius: 6px;
        overflow: hidden;
        height: 18px;
        min-width: 100px;
    }
    .mini-bar-fill {
        height: 100%;
        border-radius: 6px;
        display: flex; align-items: center; justify-content: center;
        color: white; font-size: 11px; font-weight: bold;
        min-width: 28px;
        transition: width 0.4s ease;
    }
</style>
""", unsafe_allow_html=True)


# ─── 載入資料（快取）─────────────────────────────────────────
@st.cache_data
def load_budget():
    """載入預算表（快取）"""
    if os.path.exists(BUDGET_EXCEL_PATH):
        return load_budget_from_excel()
    return None

@st.cache_data
def load_records_local(filepath: str, year: int, month: int):
    """載入本機 cwmoney 某月記帳紀錄（快取）"""
    if filepath and os.path.exists(filepath):
        return load_cwmoney_records(filepath=filepath, year=year, month=month)
    return []

@st.cache_data
def load_all_records_local(filepath: str):
    """載入本機 cwmoney 全部記帳紀錄（用於計算累積結餘）"""
    if filepath and os.path.exists(filepath):
        return load_cwmoney_records(filepath=filepath)
    return []

@st.cache_data(ttl=300)  # Google Sheets 快取 5 分鐘
def load_records_google(sheet_url: str, year: int, month: int):
    """從 Google Sheets 載入某月紀錄（快取 5 分鐘）"""
    try:
        return load_cwmoney_from_google_sheets(sheet_url, year=year, month=month)
    except Exception:
        return []

@st.cache_data(ttl=300)
def load_all_records_google(sheet_url: str):
    """從 Google Sheets 載入全部紀錄"""
    try:
        return load_cwmoney_from_google_sheets(sheet_url)
    except Exception:
        return []

def calc_cumulative_balance(all_records: list, up_to_year: int, up_to_month: int) -> float:
    """計算從最早紀錄到指定年月（含）的累積結餘"""
    total = 0.0
    for r in all_records:
        y, m = int(r["date"][:4]), int(r["date"][5:7])
        if (y, m) <= (up_to_year, up_to_month):
            if r["type"] == "收入":
                total += r["amount"]
            elif r["type"] == "支出":
                total -= r["amount"]
    return total

budget_data = load_budget()

# ─── 側邊欄 ──────────────────────────────────────────────
with st.sidebar:
    st.title("💰 Alan's 收支管理")
    st.markdown("---")

    today = date.today()
    col1, col2 = st.columns(2)
    with col1:
        selected_year = st.selectbox("📆 年份", range(today.year - 2, today.year + 2),
                                     index=2, key="year")
    with col2:
        selected_month = st.selectbox("📅 月份", range(1, 13),
                                      index=today.month - 1, key="month",
                                      format_func=lambda x: f"{x} 月")

    st.markdown("---")

    # ── 資料來源選擇 ──
    st.markdown("#### 📂 資料來源")
    data_source = st.radio(
        "選擇記帳資料來源",
        ["📁 本機檔案", "📤 上傳檔案", "☁️ Google Sheets", "🔄 Google Drive 自動同步"],
        index=3,
        key="data_source",
        horizontal=True,
        help="選擇如何取得 cwmoney 記帳資料"
    )

    records = []
    all_records = []
    source_status = ""

    # ── 本機檔案模式 ──
    if data_source == "📁 本機檔案":
        cwmoney_files = find_cwmoney_files()
        if cwmoney_files:
            file_options = {os.path.basename(f): f for f in cwmoney_files}
            selected_cwmoney_name = st.selectbox(
                "📝 記帳檔案",
                options=list(file_options.keys()),
                index=0,
                help="自動偵測資料夾內所有 cwmoney 匯出檔，最新的排最前面"
            )
            selected_cwmoney_path = file_options[selected_cwmoney_name]
            st.caption(f"✅ 共找到 {len(cwmoney_files)} 個記帳檔案")
            source_status = f"📁 {selected_cwmoney_name}"

            records = load_records_local(selected_cwmoney_path, selected_year, selected_month)
            all_records = load_all_records_local(selected_cwmoney_path)
        else:
            st.warning(f"⚠️ 在以下資料夾找不到 cwmoney 匯出檔：\n`{DATA_DIR}`")
            st.caption("請將 cwmoney 匯出的 xlsx 檔放入上述資料夾")

    # ── 上傳檔案模式 ──
    elif data_source == "📤 上傳檔案":
        uploaded_file = st.file_uploader(
            "拖放或選擇 cwmoney 匯出檔",
            type=["xlsx", "xls", "csv"],
            key="cwmoney_upload",
            help="支援 .xlsx、.xls、.csv 格式"
        )
        if uploaded_file:
            file_bytes = uploaded_file.read()
            records = load_cwmoney_from_uploaded_file(
                file_bytes, uploaded_file.name,
                year=selected_year, month=selected_month
            )
            all_records = load_cwmoney_from_uploaded_file(
                file_bytes, uploaded_file.name
            )
            source_status = f"📤 {uploaded_file.name}"
            st.success(f"✅ 已讀取 {len(all_records)} 筆紀錄")
        else:
            st.info("📎 請上傳 cwmoney 匯出的 Excel 或 CSV 檔案")
            st.caption("手機 cwmoney → 匯出 → 傳到電腦 → 拖放至此")

    # ── Google Sheets 模式 ──
    elif data_source == "☁️ Google Sheets":
        cloud_config = load_cloud_config()
        saved_url = cloud_config.get("google_sheets_url", "")

        sheet_url = st.text_input(
            "🔗 Google Sheets 連結",
            value=saved_url,
            placeholder="https://docs.google.com/spreadsheets/d/...",
            help="貼上已發佈的 Google Sheets CSV 連結"
        )

        # 設定說明（可摺疊）
        with st.expander("📖 如何設定 Google Sheets 同步？"):
            st.markdown("""
            **步驟：**
            1. 將 cwmoney 匯出的資料貼到 Google Sheets
            2. 在 Google Sheets 中：**檔案** → **共用** → **發佈到網路**
            3. 選擇要發佈的工作表 → 格式選 **CSV**
            4. 點「發佈」並複製連結
            5. 將連結貼到上方輸入框

            **💡 小技巧：**
            - 之後只要更新 Google Sheets 上的資料，App 就會自動讀取最新版
            - 也可以直接貼一般的 Google Sheets 分享連結
            - 資料快取 5 分鐘，點「重新同步」可立即更新
            """)

        if sheet_url:
            is_valid, msg = validate_google_sheets_url(sheet_url)
            if is_valid:
                # 儲存 URL
                if sheet_url != saved_url:
                    cloud_config["google_sheets_url"] = sheet_url
                    save_cloud_config(cloud_config)

                col_sync1, col_sync2 = st.columns([2, 1])
                with col_sync2:
                    force_refresh = st.button("🔄 同步", key="gs_sync")

                if force_refresh:
                    load_records_google.clear()
                    load_all_records_google.clear()
                    st.rerun()

                try:
                    with st.spinner("☁️ 正在從 Google Sheets 載入..."):
                        records = load_records_google(sheet_url, selected_year, selected_month)
                        all_records = load_all_records_google(sheet_url)
                    sync_time = update_last_sync_time()
                    source_status = "☁️ Google Sheets"
                    with col_sync1:
                        st.caption(f"✅ 已同步 {len(all_records)} 筆 | {sync_time}")
                except Exception as e:
                    st.error(f"❌ 無法連線到 Google Sheets：{e}")
                    st.caption("請確認連結正確且已發佈到網路")
            else:
                st.warning(msg)
        else:
            st.info("☁️ 請輸入 Google Sheets 連結以啟用線上同步")


    # ── Google Drive 自動同步模式 ──
    elif data_source == "🔄 Google Drive 自動同步":
        st.markdown("#### Google Drive 自動同步設定")
        st.info("1. 本地端：上傳 credentials.json 並輸入 Folder ID，會自動記住。\n2. 雲端部署：請將 Service Account JSON 內容貼到 .streamlit/secrets.toml 的 gdrive_credentials 欄位，folder_id 也寫入 gdrive_folder_id 欄位，App 會自動偵測。\n\n如兩者皆有，優先使用 secrets。")

        # 1. 先檢查 st.secrets（雲端部署推薦）
        cred_bytes = None
        folder_id = ""
        has_secrets = False
        try:
            has_secrets = "gdrive_credentials" in st.secrets and "gdrive_folder_id" in st.secrets
        except Exception:
            pass
        if has_secrets:
            import io
            cred_json = st.secrets["gdrive_credentials"]
            if isinstance(cred_json, str):
                cred_bytes = cred_json.encode("utf-8")
            else:
                # dict 轉 json string
                cred_bytes = json.dumps(dict(cred_json)).encode("utf-8")
            folder_id = st.secrets["gdrive_folder_id"]
            st.success("已自動載入雲端 secrets 設定，不需手動上傳/輸入")
        else:
            # 2. fallback 本地 config/上傳
            DRIVE_CONFIG_PATH = os.path.join(os.path.dirname(__file__), "drive_config.json")
            drive_config = {"folder_id": "", "credentials_path": "credentials.json"}
            if os.path.exists(DRIVE_CONFIG_PATH):
                try:
                    with open(DRIVE_CONFIG_PATH, "r", encoding="utf-8") as f:
                        drive_config = json.load(f)
                except Exception:
                    pass
            cred_file = st.file_uploader("上傳 credentials.json", type=["json"], key="drive_cred")
            folder_id = st.text_input("Google Drive Folder ID", value=drive_config.get("folder_id", ""), key="drive_folder_id", help="Google Drive 資料夾網址中 /folders/ 後面的那串字母")
            cred_path = drive_config.get("credentials_path", "credentials.json")
            if cred_file:
                cred_bytes = cred_file.read()
                with open(os.path.join(os.path.dirname(__file__), cred_path), "wb") as f:
                    f.write(cred_bytes)
            elif os.path.exists(os.path.join(os.path.dirname(__file__), cred_path)):
                with open(os.path.join(os.path.dirname(__file__), cred_path), "rb") as f:
                    cred_bytes = f.read()
            if (cred_file or folder_id) and folder_id:
                try:
                    with open(DRIVE_CONFIG_PATH, "w", encoding="utf-8") as f:
                        json.dump({"folder_id": folder_id, "credentials_path": cred_path}, f, ensure_ascii=False, indent=2)
                except Exception:
                    pass

        drive_status = ""
        if cred_bytes and folder_id:
            try:
                with st.spinner("🔄 正在從 Google Drive 讀取最新檔案..."):
                    # 先嘗試 CSV，找不到再 fallback XML
                    try:
                        df, file_name = get_latest_csv_dataframe(cred_bytes, folder_id)
                        file_type = 'csv'
                    except Exception:
                        # XML：不傳 year/month，取得全部資料
                        df, file_name = get_latest_xml_dataframe(cred_bytes, folder_id)
                        file_type = 'xml'
                all_records = parse_cwmoney_dataframe(df)
                # 篩選當月紀錄
                records = [r for r in all_records
                           if int(r['date'][:4]) == selected_year
                           and int(r['date'][5:7]) == selected_month]
                drive_status = f"✅ 已載入 {file_name}（{file_type.upper()}），共 {len(all_records)} 筆紀錄（本月 {len(records)} 筆）"
                source_status = f"🔄 Google Drive: {file_name}"
                st.success(drive_status)
            except Exception as e:
                st.error(f"❌ Google Drive 讀取失敗：{e}")
        else:
            st.info("請上傳 credentials.json 並輸入 Folder ID，或已自動載入本地/雲端設定")

    if budget_data:
        budget_name = os.path.basename(BUDGET_EXCEL_PATH) if os.path.exists(BUDGET_EXCEL_PATH) else "已上傳的預算表"
        st.caption(f"📊 預算表：{budget_name}")

    st.markdown("---")
    st.markdown(f"### 📌 目前檢視")
    st.markdown(f"## {selected_year} 年 {selected_month} 月")
    if source_status:
        st.caption(f"資料來源：{source_status}")

    # 載入當月紀錄
    if records and budget_data:
        summary = get_cwmoney_monthly_summary(records, budget_data)

        # 計算累積結餘
        cumulative_balance = calc_cumulative_balance(
            all_records, selected_year, selected_month
        )

        st.markdown("---")
        st.markdown("#### ⚡ 快速摘要")
        st.metric("收入", f"${summary['total_income']:,.0f}")
        st.metric("支出", f"${summary['total_expense']:,.0f}")
        st.metric("📅 單月結餘", f"${summary['balance']:,.0f}",
                  delta=f"{'盈餘' if summary['balance'] >= 0 else '透支'}")
        st.metric("📊 累積結餘", f"${cumulative_balance:,.0f}",
                  delta=f"{'正值' if cumulative_balance >= 0 else '負值'}")
        if summary["is_over_budget"]:
            st.error("⚠️ 本月已超出預算！")
    else:
        summary = None
        if not budget_data:
            st.warning("⚠️ 找不到預算表 Excel")
        if not records:
            st.info("本月尚無 cwmoney 紀錄")

    st.markdown("---")
    if st.button("🔄 重新掃描檔案並載入"):
        st.cache_data.clear()
        st.rerun()


# ─── 主要頁面標籤 ──────────────────────────────────────────
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📊 月報總覽", "📝 記帳明細", "💰 預算總表", "📈 分析報表", "📅 年度總覽"
])


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# TAB 1: 月報總覽
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
with tab1:
    st.header(f"📊 {selected_year}年{selected_month}月 收支報告")

    if not summary:
        st.info("📭 尚無資料，請確認預算表與 cwmoney 記帳檔案路徑正確。")
    else:
        # ── 頂部指標卡片 ──
        c1, c2, c3, c4, c5 = st.columns(5)

        with c1:
            st.markdown(f"""
            <div class="metric-card income">
                <h3>💰 總收入</h3>
                <h1>${summary['total_income']:,.0f}</h1>
            </div>
            """, unsafe_allow_html=True)

        with c2:
            st.markdown(f"""
            <div class="metric-card expense">
                <h3>💸 總支出</h3>
                <h1>${summary['total_expense']:,.0f}</h1>
            </div>
            """, unsafe_allow_html=True)

        with c3:
            bal_class = "balance-positive" if summary['balance'] >= 0 else "balance-negative"
            bal_icon = "📈" if summary['balance'] >= 0 else "📉"
            st.markdown(f"""
            <div class="metric-card {bal_class}">
                <h3>{bal_icon} 結餘</h3>
                <h1>${summary['balance']:,.0f}</h1>
            </div>
            """, unsafe_allow_html=True)

        with c4:
            remaining = summary['budget_remaining']
            if remaining >= 0:
                rem_class = "income"
                rem_label = "🎯 預算剩餘"
                rem_value = f"${remaining:,.0f}"
            else:
                rem_class = "expense"
                rem_label = "🚨 預算超支"
                rem_value = f"-${abs(remaining):,.0f}"
            st.markdown(f"""
            <div class="metric-card {rem_class}">
                <h3>{rem_label}</h3>
                <h1>{rem_value}</h1>
            </div>
            """, unsafe_allow_html=True)

        with c5:
            savings_budget = summary['total_savings_budget']
            actual_savings = summary['actual_savings_expense']
            savings_rate = (actual_savings / savings_budget * 100) if savings_budget > 0 else 0
            if savings_rate >= 100:
                sav_class = "income"
                sav_icon = "🏆"
                sav_label = "儲蓄達成率"
            elif savings_rate >= 60:
                sav_class = "savings"
                sav_icon = "🏦"
                sav_label = "儲蓄達成率"
            else:
                sav_class = "expense"
                sav_icon = "⚠️"
                sav_label = "儲蓄達成率"
            st.markdown(f"""
            <div class="metric-card {sav_class}">
                <h3>{sav_icon} {sav_label}</h3>
                <h1>{savings_rate:.0f}%</h1>
                <p style="margin:4px 0 0 0;font-size:13px;opacity:0.85;">
                    ${actual_savings:,.0f} / ${savings_budget:,.0f}
                </p>
            </div>
            """, unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # ── 透支警告 ──
        if summary["is_over_budget"]:
            over_amount = abs(summary["budget_remaining"])
            st.error(f"🚨 **本月已超出預算！** 超支 **${over_amount:,.0f}** 元（不含儲蓄預算）")

        # ── 預算使用率（按專案層級）── 圓環卡片風格 ──
        st.subheader("🎯 各專案預算使用率")
        gauges = create_budget_usage_gauges(summary["budget_vs_actual_by_project"])

        if gauges:
            cols_per_row = 4
            for i in range(0, len(gauges), cols_per_row):
                cols = st.columns(cols_per_row)
                for j, col in enumerate(cols):
                    if i + j < len(gauges):
                        g = gauges[i + j]
                        with col:
                            diff_class = "under" if g['diff'] >= 0 else "over"
                            diff_label = "剩餘" if g['diff'] >= 0 else "超支"
                            # 用 conic-gradient 畫圓環
                            pct_ring = min(g['pct'], 100)
                            ring_bg = f"background: conic-gradient({g['color']} {pct_ring * 3.6}deg, #ecf0f1 0deg);"
                            st.markdown(f"""
                            <div class="project-card">
                                <div class="ring" style="{ring_bg}">
                                    <span style="background:#fff;width:62px;height:62px;border-radius:50%;display:flex;align-items:center;justify-content:center;color:{g['color']};font-size:18px;">
                                        {g['pct']:.0f}%
                                    </span>
                                </div>
                                <div class="title">{g['status']} {g['category']}</div>
                                <div class="detail">
                                    已花 <b>${g['actual']:,.0f}</b> / ${g['budget']:,.0f}<br>
                                    <span class="{diff_class}">{diff_label} ${abs(g['diff']):,.0f}</span>
                                </div>
                            </div>
                            """, unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # ── 主分類預算使用率 ── 表格行＋迷你進度條 ──
        st.subheader("📋 各主分類預算使用率")
        gauges_main = create_budget_usage_gauges(summary["budget_vs_actual_by_main"])

        if gauges_main:
            table_html = '<table class="main-cat-table">'
            table_html += '<tr><th>分類</th><th>已花 / 預算</th><th>使用率</th><th>差額</th></tr>'
            for g in gauges_main:
                pct_display = min(g['pct'], 100)
                diff_color = "#27ae60" if g['diff'] >= 0 else "#e74c3c"
                diff_label = "剩餘" if g['diff'] >= 0 else "超支"
                table_html += f'''
                <tr class="main-cat-row">
                    <td>{g['status']} {g['category']}</td>
                    <td>${g['actual']:,.0f} <span style="color:#aaa">/</span> ${g['budget']:,.0f}</td>
                    <td>
                        <div class="mini-bar">
                            <div class="mini-bar-fill" style="width:{max(pct_display, 5)}%;background:{g['color']}">
                                {g['pct']:.0f}%
                            </div>
                        </div>
                    </td>
                    <td style="color:{diff_color};font-weight:600;">{diff_label} ${abs(g['diff']):,.0f}</td>
                </tr>'''
            table_html += '</table>'
            st.markdown(table_html, unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # ── 圖表區 ──
        col_left, col_right = st.columns(2)

        with col_left:
            fig_pie = create_expense_pie_chart(
                summary["expense_by_project"], title="🍩 按專案（食住行育樂）佔比"
            )
            st.plotly_chart(fig_pie, width="stretch")

        with col_right:
            fig_pie_main = create_expense_pie_chart(
                summary["expense_by_main_category"], title="🍩 按主分類佔比"
            )
            st.plotly_chart(fig_pie_main, width="stretch")

        # 每日支出走勢
        fig_daily = create_daily_expense_chart(
            summary["daily_expenses"], selected_year, selected_month,
            budget_total=summary["budget_total_no_savings"]
        )
        st.plotly_chart(fig_daily, width="stretch")

        # 預算 vs 實際（主分類）
        fig_bar = create_budget_vs_actual_chart(
            summary["budget_vs_actual_by_main"],
            title="📊 預算 vs 實際支出（按主分類）"
        )
        st.plotly_chart(fig_bar, width="stretch")


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# TAB 2: 記帳明細
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
with tab2:
    st.header(f"📝 {selected_year}年{selected_month}月 記帳明細")
    st.caption(f"資料來源：{source_status if source_status else 'cwmoney 匯出檔'}")

    if not records:
        st.info("📭 本月尚無 cwmoney 紀錄")
    else:
        # 篩選
        filter_col1, filter_col2, filter_col3 = st.columns([1, 1, 1])
        with filter_col1:
            filter_type = st.multiselect("📌 類型篩選", ["支出", "收入"],
                                          default=["支出", "收入"], key="filter_type")
        with filter_col2:
            all_projects = sorted(set(r["project"] for r in records if r["project"]))
            filter_project = st.multiselect("📁 專案篩選", all_projects,
                                             default=all_projects, key="filter_proj")
        with filter_col3:
            sort_by = st.selectbox("🔃 排序", ["日期 ↓", "日期 ↑", "金額 ↓", "金額 ↑"])

        # 篩選
        filtered = [r for r in records
                    if r["type"] in filter_type
                    and r["project"] in filter_project]

        # 排序
        if sort_by == "日期 ↓":
            filtered = sorted(filtered, key=lambda x: x["date"], reverse=True)
        elif sort_by == "日期 ↑":
            filtered = sorted(filtered, key=lambda x: x["date"])
        elif sort_by == "金額 ↓":
            filtered = sorted(filtered, key=lambda x: x["amount"], reverse=True)
        else:
            filtered = sorted(filtered, key=lambda x: x["amount"])

        # 表格顯示
        if filtered:
            proj_icons = get_project_icons()

            df_display = pd.DataFrame([{
                "日期": r["date"],
                "類型": r["type"],
                "專案": f"{proj_icons.get(r['project'], '📌')} {r['project']}",
                "主分類": r["main_category"],
                "子分類": r["sub_category"],
                "帳戶": r["account"],
                "金額": r["amount"],
                "備註": r["note"][:50] if r["note"] else ""
            } for r in filtered])

            st.dataframe(df_display, width="stretch", hide_index=True,
                        column_config={
                            "金額": st.column_config.NumberColumn(format="$%d"),
                        })

            # 統計
            total_filtered_expense = sum(r["amount"] for r in filtered if r["type"] == "支出")
            total_filtered_income = sum(r["amount"] for r in filtered if r["type"] == "收入")

            stat1, stat2, stat3 = st.columns(3)
            with stat1:
                st.metric("📊 篩選筆數", f"{len(filtered)} 筆")
            with stat2:
                st.metric("💸 篩選支出小計", f"${total_filtered_expense:,.0f}")
            with stat3:
                st.metric("💰 篩選收入小計", f"${total_filtered_income:,.0f}")

            # 匯出 CSV
            if st.button("📥 匯出篩選結果為 CSV"):
                csv = df_display.to_csv(index=False)
                st.download_button("💾 下載 CSV 檔", csv,
                                 f"records_{selected_year}_{selected_month:02d}.csv",
                                 "text/csv")
        else:
            st.info("沒有符合篩選條件的紀錄")


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# TAB 3: 預算總表
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
with tab3:
    st.header("💰 每月預算總表")
    _budget_label = os.path.basename(BUDGET_EXCEL_PATH) if os.path.exists(BUDGET_EXCEL_PATH) else "已上傳的預算表"
    st.caption(f"資料來源：{_budget_label}")

    if not budget_data:
        st.warning("⚠️ 找不到預算表檔案，請在側邊欄上傳或確認路徑正確")
    else:
        proj_icons = get_project_icons()

        # ── 預算總計 ──
        total_budget = sum(c["budget"] for c in budget_data["categories"])
        total_no_savings = sum(c["budget"] for c in budget_data["categories"]
                               if c["expense_type"] != "儲蓄支出")
        total_savings = sum(c["budget"] for c in budget_data["categories"]
                            if c["expense_type"] == "儲蓄支出")
        total_fixed = sum(c["budget"] for c in budget_data["categories"]
                          if c["expense_type"] == "固定支出")
        total_variable = sum(c["budget"] for c in budget_data["categories"]
                             if c["expense_type"] == "變動支出")

        bc1, bc2, bc3, bc4 = st.columns(4)
        with bc1:
            st.metric("💵 預算總計", f"${total_budget:,.0f}")
        with bc2:
            st.metric("📌 固定支出", f"${total_fixed:,.0f}")
        with bc3:
            st.metric("🔄 變動支出", f"${total_variable:,.0f}")
        with bc4:
            st.metric("🏦 儲蓄支出", f"${total_savings:,.0f}")

        st.markdown("---")

        # ── 按專案分組顯示 ──
        structure = budget_data["structure"]
        for proj, mains in structure.items():
            icon = proj_icons.get(proj, "📌")
            proj_total = sum(
                sub_data["budget"]
                for main_subs in mains.values()
                for sub_data in main_subs.values()
            )

            with st.expander(f"{icon} **{proj}** — 預算合計 ${proj_total:,.0f}", expanded=False):
                for main_cat, subs in mains.items():
                    main_total = sum(s["budget"] for s in subs.values())
                    st.markdown(f"#### {main_cat}（${main_total:,.0f}）")

                    table_rows = []
                    for sub_name, sub_info in subs.items():
                        table_rows.append({
                            "子分類": sub_name,
                            "預算金額": sub_info["budget"],
                            "支出類型": sub_info["type"]
                        })

                    df_sub = pd.DataFrame(table_rows)
                    st.dataframe(df_sub, width="stretch", hide_index=True,
                                column_config={
                                    "預算金額": st.column_config.NumberColumn(format="$%d"),
                                })

        # ── 子分類對應表 ──
        if budget_data.get("sub_category_mapping"):
            st.markdown("---")
            st.subheader("🔗 子分類對應表")
            st.caption("cwmoney 的子分類如何對應到預算表的子分類")

            mapping_rows = []
            for budget_sub, items in budget_data["sub_category_mapping"].items():
                for item in items:
                    mapping_rows.append({
                        "預算子分類": budget_sub,
                        "對應 cwmoney 子分類": item
                    })

            if mapping_rows:
                df_mapping = pd.DataFrame(mapping_rows)
                st.dataframe(df_mapping, width="stretch", hide_index=True)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# TAB 4: 分析報表
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
with tab4:
    st.header(f"📈 {selected_year}年{selected_month}月 深入分析")

    if not summary:
        st.info("📭 尚無資料可分析")
    else:
        # ── 支出類型分析 ──
        col_left, col_right = st.columns(2)
        with col_left:
            fig_type = create_expense_type_pie(summary["expense_type_breakdown"])
            st.plotly_chart(fig_type, width="stretch")

        with col_right:
            # 預算 vs 實際（專案層級）
            fig_proj = create_budget_vs_actual_chart(
                summary["budget_vs_actual_by_project"],
                title="📊 預算 vs 實際（按專案）"
            )
            st.plotly_chart(fig_proj, width="stretch")

        # ── Treemap 支出分類總覽 ──
        if budget_data:
            fig_tree = create_sub_category_treemap(
                budget_data["categories"],
                summary["expense_by_sub_category"]
            )
            st.plotly_chart(fig_tree, width="stretch")

        # ── 子分類明細（有預算的項目）──
        st.subheader("📋 子分類預算執行明細")

        detail_rows = []
        for sub, data in summary["budget_vs_actual_by_sub"].items():
            if data["budget"] > 0 or data["actual"] > 0:
                pct = data["pct"]
                if pct <= 70:
                    status = "✅ 正常"
                elif pct <= 100:
                    status = "⚠️ 注意"
                else:
                    status = "🔴 超支"

                detail_rows.append({
                    "專案": data.get("project", ""),
                    "主分類": data.get("main_category", ""),
                    "子分類": sub,
                    "類型": data.get("expense_type", ""),
                    "預算": data["budget"],
                    "實際": data["actual"],
                    "差額": data["diff"],
                    "使用率": f"{pct:.0f}%",
                    "狀態": status
                })

        if detail_rows:
            df_detail = pd.DataFrame(detail_rows)
            # 排序：超支的排前面
            df_detail = df_detail.sort_values("差額", ascending=True)

            st.dataframe(df_detail, width="stretch", hide_index=True,
                        column_config={
                            "預算": st.column_config.NumberColumn(format="$%d"),
                            "實際": st.column_config.NumberColumn(format="$%d"),
                            "差額": st.column_config.NumberColumn(format="$%d"),
                        })

            # 超支項目提醒
            over_items = [r for r in detail_rows if r["差額"] < 0]
            if over_items:
                st.warning(f"⚠️ 有 **{len(over_items)}** 個子分類已超出預算：")
                for item in over_items:
                    st.markdown(
                        f"- **{item['子分類']}**（{item['專案']}）："
                        f"預算 ${item['預算']:,.0f}，"
                        f"實際 ${item['實際']:,.0f}，"
                        f"超支 ${abs(item['差額']):,.0f}"
                    )
        else:
            st.info("尚無預算執行資料")

        # ── 帳戶支出分布 ──
        st.subheader("🏦 帳戶支出分布")
        expense_records = [r for r in records if r["type"] == "支出"]
        if expense_records:
            account_totals = {}
            for r in expense_records:
                acc = r["account"]
                account_totals[acc] = account_totals.get(acc, 0) + r["amount"]

            fig_account = create_expense_pie_chart(account_totals, title="🏦 各帳戶支出佔比")
            st.plotly_chart(fig_account, width="stretch")


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# TAB 5: 年度總覽
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
with tab5:
    st.header(f"📅 {selected_year} 年度總覽")

    if not all_records or not budget_data:
        st.info("📭 尚無資料可供年度分析")
    else:
        # ── 計算各月份摘要 ──
        savings_budget = summary['total_savings_budget'] if summary else 0
        budget_no_savings = summary['budget_total_no_savings'] if summary else 0

        monthly_data = []
        for m in range(1, 13):
            m_records = [r for r in all_records
                         if int(r['date'][:4]) == selected_year
                         and int(r['date'][5:7]) == m]
            if not m_records:
                monthly_data.append({
                    "month": m, "income": 0, "expense_total": 0,
                    "expense": 0, "balance": 0, "savings": 0,
                    "savings_rate": 0, "budget_remaining": 0,
                    "record_count": 0
                })
                continue

            m_summary = get_cwmoney_monthly_summary(m_records, budget_data)
            monthly_data.append({
                "month": m,
                "income": m_summary['total_income'],
                "expense_total": m_summary['total_expense'],
                "expense": m_summary['expense_no_savings'],
                "balance": m_summary['total_income'] - m_summary['expense_no_savings'],
                "savings": m_summary['actual_savings_expense'],
                "savings_rate": (m_summary['actual_savings_expense'] / savings_budget * 100) if savings_budget > 0 else 0,
                "budget_remaining": m_summary['budget_remaining'],
                "record_count": m_summary['record_count']
            })

        # 只顯示有資料的月份
        active_months = [d for d in monthly_data if d['record_count'] > 0]

        if not active_months:
            st.info(f"📭 {selected_year} 年尚無記帳紀錄")
        else:
            # ── 年度累計卡片 ──
            yr_income = sum(d['income'] for d in active_months)
            yr_expense = sum(d['expense'] for d in active_months)
            yr_expense_total = sum(d['expense_total'] for d in active_months)
            yr_savings = sum(d['savings'] for d in active_months)
            yr_balance = yr_income - yr_expense
            yr_savings_target = savings_budget * len(active_months)
            yr_savings_rate = (yr_savings / yr_savings_target * 100) if yr_savings_target > 0 else 0

            yc1, yc2, yc3, yc4, yc5 = st.columns(5)
            with yc1:
                st.markdown(f"""
                <div class="metric-card income">
                    <h3>💵 年度總收入</h3>
                    <h1>${yr_income:,.0f}</h1>
                    <p style="margin:4px 0 0 0;font-size:12px;opacity:0.8;">共 {len(active_months)} 個月</p>
                </div>
                """, unsafe_allow_html=True)
            with yc2:
                st.markdown(f"""
                <div class="metric-card expense">
                    <h3>💸 年度總支出</h3>
                    <h1>${yr_expense:,.0f}</h1>
                    <p style="margin:4px 0 0 0;font-size:12px;opacity:0.8;">不含儲蓄</p>
                </div>
                """, unsafe_allow_html=True)
            with yc3:
                yr_bal_class = "income" if yr_balance >= 0 else "expense"
                yr_bal_icon = "📈" if yr_balance >= 0 else "📉"
                st.markdown(f"""
                <div class="metric-card {yr_bal_class}">
                    <h3>{yr_bal_icon} 年度結餘</h3>
                    <h1>${yr_balance:,.0f}</h1>
                    <p style="margin:4px 0 0 0;font-size:12px;opacity:0.8;">收入-支出(不含儲蓄)</p>
                </div>
                """, unsafe_allow_html=True)
            with yc4:
                st.markdown(f"""
                <div class="metric-card savings">
                    <h3>🏦 年度總儲蓄</h3>
                    <h1>${yr_savings:,.0f}</h1>
                    <p style="margin:4px 0 0 0;font-size:12px;opacity:0.8;">目標 ${yr_savings_target:,.0f}</p>
                </div>
                """, unsafe_allow_html=True)
            with yc5:
                if yr_savings_rate >= 100:
                    yr_sav_class = "income"
                    yr_sav_icon = "🏆"
                elif yr_savings_rate >= 60:
                    yr_sav_class = "savings"
                    yr_sav_icon = "🏦"
                else:
                    yr_sav_class = "expense"
                    yr_sav_icon = "⚠️"
                st.markdown(f"""
                <div class="metric-card {yr_sav_class}">
                    <h3>{yr_sav_icon} 年度儲蓄達成率</h3>
                    <h1>{yr_savings_rate:.0f}%</h1>
                    <p style="margin:4px 0 0 0;font-size:12px;opacity:0.8;">平均 ${yr_savings / len(active_months):,.0f}/月</p>
                </div>
                """, unsafe_allow_html=True)

            st.markdown("<br>", unsafe_allow_html=True)

            # ── 年度趨勢圖表 ──
            col_chart1, col_chart2 = st.columns(2)
            with col_chart1:
                fig_trend = create_yearly_income_expense_chart(active_months)
                st.plotly_chart(fig_trend, use_container_width=True)
            with col_chart2:
                fig_savings = create_yearly_savings_chart(active_months, savings_budget)
                st.plotly_chart(fig_savings, use_container_width=True)

            # ── 年度累積趨勢 ──
            fig_cum = create_yearly_cumulative_chart(active_months)
            st.plotly_chart(fig_cum, use_container_width=True)

            # ── 月度明細表 ──
            st.subheader("📋 月度收支明細")
            table_rows = []
            for d in active_months:
                budget_status = "✅ 未超支" if d['budget_remaining'] >= 0 else "🔴 超支"
                table_rows.append({
                    "月份": f"{d['month']}月",
                    "收入": d['income'],
                    "支出(不含儲蓄)": d['expense'],
                    "結餘": d['balance'],
                    "儲蓄": d['savings'],
                    "儲蓄達成率": f"{d['savings_rate']:.0f}%",
                    "預算剩餘": d['budget_remaining'],
                    "狀態": budget_status
                })

            # 合計列
            table_rows.append({
                "月份": "📊 合計",
                "收入": yr_income,
                "支出(不含儲蓄)": yr_expense,
                "結餘": yr_balance,
                "儲蓄": yr_savings,
                "儲蓄達成率": f"{yr_savings_rate:.0f}%",
                "預算剩餘": sum(d['budget_remaining'] for d in active_months),
                "狀態": ""
            })

            df_yearly = pd.DataFrame(table_rows)
            st.dataframe(df_yearly, width="stretch", hide_index=True,
                        column_config={
                            "收入": st.column_config.NumberColumn(format="$%d"),
                            "支出(不含儲蓄)": st.column_config.NumberColumn(format="$%d"),
                            "結餘": st.column_config.NumberColumn(format="$%d"),
                            "儲蓄": st.column_config.NumberColumn(format="$%d"),
                            "預算剩餘": st.column_config.NumberColumn(format="$%d"),
                        })

            # ── 月平均統計 ──
            n = len(active_months)
            st.subheader("📊 月平均統計")
            avg1, avg2, avg3, avg4 = st.columns(4)
            with avg1:
                st.metric("平均月收入", f"${yr_income / n:,.0f}")
            with avg2:
                st.metric("平均月支出", f"${yr_expense / n:,.0f}")
            with avg3:
                st.metric("平均月結餘", f"${yr_balance / n:,.0f}")
            with avg4:
                st.metric("平均月儲蓄", f"${yr_savings / n:,.0f}")


# ─── Footer ──────────────────────────────────────────────
st.markdown("---")
st.markdown(
    "<div style='text-align:center;color:gray;font-size:12px'>"
    "💰 Alan's 每月收支管理 App | 資料來源：自訂預算表 + cwmoney"
    "</div>",
    unsafe_allow_html=True
)
