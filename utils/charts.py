"""
圖表模組 - 支援三層分類結構（專案→主分類→子分類）的圖表視覺化
"""
import plotly.graph_objects as go
import plotly.express as px
import calendar


def create_budget_vs_actual_chart(budget_vs_actual: dict, title: str = "📊 預算 vs 實際支出") -> go.Figure:
    """建立預算 vs 實際支出的水平長條圖"""
    if not budget_vs_actual:
        fig = go.Figure()
        fig.add_annotation(text="尚無預算資料", xref="paper", yref="paper",
                          x=0.5, y=0.5, showarrow=False, font=dict(size=18, color="gray"))
        fig.update_layout(height=300)
        return fig

    # 只顯示有預算或有實際支出的類別
    filtered = {k: v for k, v in budget_vs_actual.items()
                if v["budget"] > 0 or v["actual"] > 0}
    if not filtered:
        fig = go.Figure()
        fig.add_annotation(text="尚無資料", xref="paper", yref="paper",
                          x=0.5, y=0.5, showarrow=False, font=dict(size=18, color="gray"))
        fig.update_layout(height=300)
        return fig

    categories = list(filtered.keys())
    budgets = [v["budget"] for v in filtered.values()]
    actuals = [v["actual"] for v in filtered.values()]

    bar_colors = ["#e74c3c" if a > b else "#2ecc71" for a, b in zip(actuals, budgets)]

    fig = go.Figure()
    fig.add_trace(go.Bar(
        y=categories, x=budgets, name="預算",
        orientation="h", marker_color="#74b9ff",
        text=[f"${b:,.0f}" for b in budgets], textposition="auto"
    ))
    fig.add_trace(go.Bar(
        y=categories, x=actuals, name="實際支出",
        orientation="h", marker_color=bar_colors,
        text=[f"${a:,.0f}" for a in actuals], textposition="auto"
    ))

    fig.update_layout(
        title=title,
        barmode="group",
        height=max(350, len(categories) * 55),
        xaxis_title="金額 (NT$)",
        font=dict(size=13),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        margin=dict(l=10, r=10, t=60, b=40)
    )
    return fig


def create_expense_pie_chart(expense_by_category: dict, title: str = "🍩 支出分類佔比") -> go.Figure:
    """建立支出分類圓餅圖"""
    if not expense_by_category:
        fig = go.Figure()
        fig.add_annotation(text="尚無支出紀錄", xref="paper", yref="paper",
                          x=0.5, y=0.5, showarrow=False, font=dict(size=18, color="gray"))
        fig.update_layout(height=350)
        return fig

    filtered = {k: v for k, v in expense_by_category.items() if v > 0}
    if not filtered:
        fig = go.Figure()
        fig.add_annotation(text="尚無支出紀錄", xref="paper", yref="paper",
                          x=0.5, y=0.5, showarrow=False, font=dict(size=18, color="gray"))
        fig.update_layout(height=350)
        return fig

    categories = list(filtered.keys())
    values = list(filtered.values())
    colors = px.colors.qualitative.Set3[:len(categories)]

    fig = go.Figure(data=[go.Pie(
        labels=categories, values=values,
        hole=0.4,
        textinfo="label+percent",
        textposition="outside",
        marker=dict(colors=colors),
        hovertemplate="<b>%{label}</b><br>金額: $%{value:,.0f}<br>佔比: %{percent}<extra></extra>"
    )])

    total = sum(values)
    fig.update_layout(
        title=title,
        height=450,
        font=dict(size=12),
        annotations=[dict(text=f"${total:,.0f}", x=0.5, y=0.5, font_size=18, showarrow=False)],
        margin=dict(l=10, r=10, t=60, b=10)
    )
    return fig


def create_daily_expense_chart(daily_expenses: dict, year: int, month: int,
                                budget_total: float = 0) -> go.Figure:
    """建立每日支出走勢圖（含預算線）"""
    days_in_month = calendar.monthrange(year, month)[1]
    all_dates = [f"{year}-{month:02d}-{d:02d}" for d in range(1, days_in_month + 1)]
    amounts = [daily_expenses.get(d, 0) for d in all_dates]

    cumulative = []
    total = 0
    for a in amounts:
        total += a
        cumulative.append(total)

    day_labels = [str(i) for i in range(1, days_in_month + 1)]

    fig = go.Figure()

    fig.add_trace(go.Bar(
        x=day_labels, y=amounts, name="每日支出",
        marker_color="#74b9ff",
        hovertemplate="日期: %{x}日<br>支出: $%{y:,.0f}<extra></extra>"
    ))

    fig.add_trace(go.Scatter(
        x=day_labels, y=cumulative, name="累積支出",
        mode="lines+markers",
        line=dict(color="#e17055", width=2.5),
        marker=dict(size=4),
        yaxis="y2",
        hovertemplate="日期: %{x}日<br>累積: $%{y:,.0f}<extra></extra>"
    ))

    if budget_total > 0:
        fig.add_trace(go.Scatter(
            x=day_labels,
            y=[budget_total] * len(day_labels),
            name=f"預算上限 ${budget_total:,.0f}",
            mode="lines",
            line=dict(color="#d63031", width=2, dash="dash"),
            yaxis="y2",
            hovertemplate="預算上限: $%{y:,.0f}<extra></extra>"
        ))

    fig.update_layout(
        title=f"📅 {year}年{month}月 每日支出走勢",
        xaxis_title="日期",
        yaxis=dict(title="每日支出 (NT$)", side="left"),
        yaxis2=dict(title="累積支出 (NT$)", side="right", overlaying="y"),
        height=380,
        font=dict(size=12),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        margin=dict(l=10, r=10, t=60, b=40)
    )
    return fig


def create_budget_usage_gauges(budget_vs_actual: dict) -> list:
    """建立各類別預算使用率資料（只顯示有預算的項目）"""
    gauges = []
    for cat, data in budget_vs_actual.items():
        if data["budget"] <= 0 and data["actual"] <= 0:
            continue
        pct = data["pct"]
        if pct <= 70:
            color = "#2ecc71"
            status = "✅"
        elif pct <= 100:
            color = "#f39c12"
            status = "⚠️"
        else:
            color = "#e74c3c"
            status = "🔴"
        gauges.append({
            "category": cat,
            "budget": data["budget"],
            "actual": data["actual"],
            "diff": data["diff"],
            "pct": pct,
            "color": color,
            "status": status
        })
    return gauges


def create_expense_type_pie(expense_type_breakdown: dict) -> go.Figure:
    """建立固定/變動/儲蓄支出佔比圖"""
    filtered = {k: v for k, v in expense_type_breakdown.items() if v > 0}
    if not filtered:
        fig = go.Figure()
        fig.add_annotation(text="尚無資料", xref="paper", yref="paper",
                          x=0.5, y=0.5, showarrow=False, font=dict(size=18, color="gray"))
        fig.update_layout(height=300)
        return fig

    type_colors = {"固定支出": "#3498db", "變動支出": "#e67e22", "儲蓄支出": "#2ecc71"}
    colors = [type_colors.get(k, "#95a5a6") for k in filtered.keys()]

    fig = go.Figure(data=[go.Pie(
        labels=list(filtered.keys()),
        values=list(filtered.values()),
        hole=0.4,
        marker=dict(colors=colors),
        textinfo="label+percent+value",
        texttemplate="%{label}<br>$%{value:,.0f}<br>(%{percent})",
        hovertemplate="<b>%{label}</b><br>$%{value:,.0f}<br>%{percent}<extra></extra>"
    )])

    fig.update_layout(
        title="📋 支出類型佔比（固定/變動/儲蓄）",
        height=380,
        font=dict(size=12),
        margin=dict(l=10, r=10, t=60, b=10)
    )
    return fig


def create_sub_category_treemap(budget_data_categories: list, expense_by_sub: dict) -> go.Figure:
    """建立子分類支出的 Treemap 圖"""
    labels = []
    parents = []
    values = []
    colors_list = []

    proj_icons = {
        "食": "🍜", "住": "🏠", "行": "🚗", "育": "📚",
        "大寶": "👶", "樂": "🎮", "衣": "👕", "儲蓄": "💰"
    }

    projects_with_data = set()
    mains_with_data = set()

    for cat in budget_data_categories:
        sub = cat["sub_category"]
        actual = expense_by_sub.get(sub, 0)
        if actual > 0:
            projects_with_data.add(cat["project"])
            mains_with_data.add((cat["project"], cat["main_category"]))

    if not projects_with_data:
        fig = go.Figure()
        fig.add_annotation(text="尚無支出紀錄", xref="paper", yref="paper",
                          x=0.5, y=0.5, showarrow=False, font=dict(size=18, color="gray"))
        fig.update_layout(height=400)
        return fig

    labels.append("支出總覽")
    parents.append("")
    values.append(0)
    colors_list.append("#ecf0f1")

    for proj in projects_with_data:
        icon = proj_icons.get(proj, "📌")
        label = f"{icon} {proj}"
        labels.append(label)
        parents.append("支出總覽")
        values.append(0)
        colors_list.append("#dfe6e9")

    for proj, main in mains_with_data:
        icon = proj_icons.get(proj, "📌")
        parent_label = f"{icon} {proj}"
        label = f"{main}"
        if label not in labels:
            labels.append(label)
            parents.append(parent_label)
            values.append(0)
            colors_list.append("#b2bec3")

    for cat in budget_data_categories:
        sub = cat["sub_category"]
        actual = expense_by_sub.get(sub, 0)
        if actual > 0:
            main = cat["main_category"]
            labels.append(sub)
            parents.append(main)
            values.append(actual)
            budget = cat["budget"]
            if budget > 0 and actual > budget:
                colors_list.append("#e74c3c")
            elif budget > 0 and actual > budget * 0.7:
                colors_list.append("#f39c12")
            else:
                colors_list.append("#2ecc71")

    fig = go.Figure(go.Treemap(
        labels=labels,
        parents=parents,
        values=values,
        marker=dict(colors=colors_list),
        textinfo="label+value",
        texttemplate="%{label}<br>$%{value:,.0f}",
        hovertemplate="<b>%{label}</b><br>$%{value:,.0f}<extra></extra>"
    ))

    fig.update_layout(
        title="🗂️ 支出分類總覽（專案→主分類→子分類）",
        height=500,
        margin=dict(l=10, r=10, t=60, b=10)
    )
    return fig
