"""
Dashboard Charts Module - Director-Level Business Intelligence
Builds Plotly charts with explanatory annotations for IPP Production.
"""
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import pandas as pd
import numpy as np


# Premium color palette
COLORS = {
    'primary': '#4F46E5',      # Indigo
    'secondary': '#7C3AED',    # Purple
    'accent': '#DB2777',       # Pink
    'success': '#059669',      # Emerald
    'warning': '#D97706',      # Amber
    'danger': '#DC2626',       # Red
    'info': '#2563EB',         # Blue
    'dark': '#111827',         # Gray-900
    'bg': '#FFFFFF',           # White
    'card': '#F8FAFC',         # Slate-50
    'text': '#1E293B',         # Slate-800
    'muted': '#64748B',        # Slate-500
}

CHART_COLORS = ['#6366F1', '#EC4899', '#10B981', '#F59E0B', '#3B82F6',
                '#8B5CF6', '#EF4444', '#14B8A6', '#F97316', '#A855F7']

LAYOUT_DEFAULTS = dict(
    paper_bgcolor='rgba(0,0,0,0)',
    plot_bgcolor='rgba(0,0,0,0)',
    font=dict(family='Inter, sans-serif', color=COLORS['text'], size=12),
    margin=dict(l=20, r=20, t=50, b=20),
    legend=dict(bgcolor='rgba(0,0,0,0)', font=dict(size=11)),
)

# ── Column name resolution ──
# Jobtrack DataFrame columns use Rate, Rate.1, etc. but charts
# reference them as 'Film Rate', '1st Fresh Rate', etc.
# This map resolves the correct column name dynamically.
COL_ALIASES = {
    'Film Rate': ['Film Rate', 'Rate'],
    'Film Value': ['Film Value'],
    '1st Fresh Rate': ['1st Fresh Rate', 'Rate.1'],
    '1st Fresh Value': ['1st Fresh Value'],
    '2nd Fresh Rate': ['2nd Fresh Rate', 'Rate.2'],
    '2nd Fresh Value': ['2nd Fresh Value'],
    'Adh Rate': ['Adh Rate', 'Rate.3'],
    'Adh Value': ['Adh Value'],
    'Hard Rate': ['Hard Rate', 'Rate.4'],
    'Hard Value': ['Hard Value'],
    'Sol Rate': ['Sol Rate', 'Rate.5'],
    'Sol Value': ['Sol Value'],
    'Waste': ['Waste', 'Total Wastage'],
    'Net Wt. (Kgs-Output)': ['Net Wt. (Kgs-Output)'],
    '1st Input Name': ['1st Input Name'],
    '1st Fresh Material': ['1st Fresh Material'],
    '2nd Fresh Material': ['2nd Fresh Material'],
}


def _col(df, name):
    """Resolve a column name using aliases. Returns the actual column name or None."""
    if name in df.columns:
        return name
    for alias in COL_ALIASES.get(name, []):
        if alias in df.columns:
            return alias
    return None


def _apply_layout(fig, title=None, height=380):
    """Apply standard chart styling."""
    fig.update_layout(
        **LAYOUT_DEFAULTS,
        title=dict(text=title, font=dict(size=16, color=COLORS['text'])) if title else None,
        height=height,
    )
    fig.update_xaxes(gridcolor='rgba(0,0,0,0.06)', zeroline=False)
    fig.update_yaxes(gridcolor='rgba(0,0,0,0.06)', zeroline=False)
    return fig


# ─────────────────────────────────────────────────────────────────────
# 1. COST BREAKDOWN DONUT
# ─────────────────────────────────────────────────────────────────────
def create_cost_breakdown_pie(stats: dict) -> go.Figure:
    """Material cost breakdown by category.
    HOW TO READ: Each slice = one cost category's share of total material spend.
    The center shows total AED. Hover for exact values."""
    labels = ['Film', '1st Fresh', '2nd Fresh', 'Adhesive', 'Hardener', 'Solvent']
    values = [
        stats.get('total_film_cost', 0),
        stats.get('total_fresh1_cost', 0),
        stats.get('total_fresh2_cost', 0),
        stats.get('total_adh_cost', 0),
        stats.get('total_hard_cost', 0),
        stats.get('total_sol_cost', 0),
    ]
    data = [(l, v) for l, v in zip(labels, values) if v > 0]
    if not data:
        return go.Figure()
    labels, values = zip(*data)

    fig = go.Figure(data=[go.Pie(
        labels=labels, values=values, hole=0.55,
        marker=dict(colors=CHART_COLORS[:len(labels)]),
        textinfo='label+percent',
        textfont=dict(size=12, color='white'),
        hovertemplate='<b>%{label}</b><br>Cost: AED %{value:,.2f}<br>Share: %{percent}<extra></extra>',
    )])
    total = sum(values)
    fig.add_annotation(
        text=f"<b>AED<br>{total:,.0f}</b>",
        x=0.5, y=0.5, font=dict(size=18, color=COLORS['text']),
        showarrow=False,
    )
    return _apply_layout(fig, "Material Cost Breakdown", 400)


# ─────────────────────────────────────────────────────────────────────
# 2. PROCESS COST STACKED BAR
# ─────────────────────────────────────────────────────────────────────
def _resolve_cost_cols(df):
    """Resolve cost value columns — returns list of (real_col_name, label) that exist."""
    mapping = [
        ('Film Value', 'Film'), ('1st Fresh Value', '1st Fresh'),
        ('2nd Fresh Value', '2nd Fresh'), ('Adh Value', 'Adhesive'),
        ('Hard Value', 'Hardener'), ('Sol Value', 'Solvent'),
    ]
    result = []
    for col_name, label in mapping:
        real = _col(df, col_name)
        if real:
            result.append((real, label))
    return result


def create_process_cost_bar(df: pd.DataFrame) -> go.Figure:
    """Stacked cost by process (Printing vs LAM)."""
    process_col = 'Process'
    if process_col not in df.columns:
        return None

    resolved = _resolve_cost_cols(df)
    if not resolved:
        return None

    fig = go.Figure()
    has_data = False
    for i, (real_col, name) in enumerate(resolved):
        grouped = df.groupby(process_col)[real_col].sum().reset_index()
        grouped = grouped[grouped[real_col] > 0]
        if grouped.empty:
            continue
        has_data = True
        fig.add_trace(go.Bar(
            x=grouped[process_col], y=grouped[real_col], name=name,
            marker_color=CHART_COLORS[i],
            hovertemplate=f'<b>{name}</b><br>Process: %{{x}}<br>Cost: AED %{{y:,.2f}}<extra></extra>',
        ))
    if not has_data:
        return None
    fig.update_layout(barmode='stack')
    return _apply_layout(fig, "Cost by Process (Stacked)", 400)


# ─────────────────────────────────────────────────────────────────────
# 3. DAILY COST TREND
# ─────────────────────────────────────────────────────────────────────
def create_daily_cost_trend(df: pd.DataFrame) -> go.Figure:
    """Daily production cost trend with moving average."""
    date_col = 'Date'
    if date_col not in df.columns:
        return None
    resolved = _resolve_cost_cols(df)
    available = [r[0] for r in resolved]
    if not available:
        return None

    df_c = df.copy()
    df_c[date_col] = pd.to_datetime(df_c[date_col], errors='coerce')
    df_c = df_c.dropna(subset=[date_col])
    df_c['Total Cost'] = df_c[available].fillna(0).sum(axis=1)
    daily = df_c.groupby(date_col)['Total Cost'].sum().reset_index().sort_values(date_col)
    if daily.empty or daily['Total Cost'].sum() <= 0:
        return None

    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=daily[date_col], y=daily['Total Cost'],
        mode='lines+markers', name='Daily Cost',
        fill='tozeroy', line=dict(color=COLORS['primary'], width=2),
        marker=dict(size=6, color=COLORS['accent']),
        fillcolor='rgba(99, 102, 241, 0.12)',
        hovertemplate='<b>%{x|%b %d}</b><br>Cost: AED %{y:,.0f}<extra></extra>',
    ))
    if len(daily) >= 3:
        daily['MA3'] = daily['Total Cost'].rolling(3, min_periods=1).mean()
        fig.add_trace(go.Scatter(
            x=daily[date_col], y=daily['MA3'], mode='lines', name='3-Day Avg',
            line=dict(color=COLORS['warning'], width=2, dash='dash'),
            hovertemplate='<b>%{x|%b %d}</b><br>3-Day Avg: AED %{y:,.0f}<extra></extra>',
        ))

    return _apply_layout(fig, "Daily Material Cost Trend", 380)


# ─────────────────────────────────────────────────────────────────────
# 4. MATERIAL RATE COMPARISON
# ─────────────────────────────────────────────────────────────────────
def create_rate_comparison_chart(df: pd.DataFrame) -> go.Figure:
    """Compare material rates across orders."""
    rate_data = []
    for col_name, label in [('Film Rate', 'Film'), ('1st Fresh Rate', 'Fresh 1'), ('2nd Fresh Rate', 'Fresh 2')]:
        real_col = _col(df, col_name)
        if real_col:
            vals = pd.to_numeric(df[real_col], errors='coerce').dropna()
            vals = vals[vals > 0]
            if not vals.empty:
                for v in vals:
                    rate_data.append({'Material': label, 'Rate (AED/Kg)': v})

    if not rate_data:
        return None

    rate_df = pd.DataFrame(rate_data)
    fig = go.Figure()
    for i, mat in enumerate(rate_df['Material'].unique()):
        subset = rate_df[rate_df['Material'] == mat]
        fig.add_trace(go.Box(
            y=subset['Rate (AED/Kg)'], name=mat,
            marker_color=CHART_COLORS[i],
            boxmean='sd',
            hovertemplate='<b>%{x}</b><br>Rate: AED %{y:.3f}<extra></extra>',
        ))
    return _apply_layout(fig, "Material Rate Distribution", 380)


# ─────────────────────────────────────────────────────────────────────
# 5. TOP ORDERS BY COST
# ─────────────────────────────────────────────────────────────────────
def create_top_orders_chart(df: pd.DataFrame) -> go.Figure:
    """Top 15 most expensive production orders.
    HOW TO READ: Longest bar = highest material cost order.
    Review top orders to identify high-cost production runs."""
    cost_cols = ['Film Value', '1st Fresh Value', '2nd Fresh Value',
                 'Adh Value', 'Hard Value', 'Sol Value']
    available = [c for c in cost_cols if c in df.columns]
    if not available:
        return go.Figure()

    order_col = 'Order No'
    if order_col not in df.columns:
        return go.Figure()

    df_c = df.copy()
    df_c['Total Cost'] = df_c[available].fillna(0).sum(axis=1)
    order_cost = df_c.groupby(order_col)['Total Cost'].sum().reset_index()
    order_cost = order_cost.nlargest(15, 'Total Cost').sort_values('Total Cost', ascending=True)

    fig = go.Figure(go.Bar(
        x=order_cost['Total Cost'], y=order_cost[order_col],
        orientation='h',
        marker=dict(
            color=order_cost['Total Cost'],
            colorscale=[[0, COLORS['info']], [0.5, COLORS['warning']], [1, COLORS['danger']]],
        ),
        hovertemplate='<b>%{y}</b><br>Total Cost: AED %{x:,.2f}<extra></extra>',
    ))
    return _apply_layout(fig, "Top 15 Orders by Material Cost", 450)


# ─────────────────────────────────────────────────────────────────────
# 6. WASTAGE ANALYSIS
# ─────────────────────────────────────────────────────────────────────
def create_wastage_chart(df: pd.DataFrame) -> go.Figure:
    """Wastage by process."""
    waste_real = _col(df, 'Waste')
    process_col = 'Process'
    if not waste_real or process_col not in df.columns:
        return None

    df_c = df.copy()
    df_c[waste_real] = pd.to_numeric(df_c[waste_real], errors='coerce').fillna(0)
    grouped = df_c.groupby(process_col).agg({waste_real: 'sum'}).reset_index()
    grouped = grouped[grouped[waste_real] > 0]
    if grouped.empty:
        return None

    fig = go.Figure(go.Bar(
        x=grouped[process_col], y=grouped[waste_real],
        marker=dict(
            color=grouped[waste_real],
            colorscale=[[0, COLORS['success']], [0.5, COLORS['warning']], [1, COLORS['danger']]],
        ),
        hovertemplate='<b>%{x}</b><br>Waste: %{y:,.1f} Kg<extra></extra>',
        text=grouped[waste_real].apply(lambda x: f'{x:,.0f} Kg'),
        textposition='auto', textfont=dict(color='white', size=12),
    ))
    return _apply_layout(fig, "Wastage by Process (Kg)", 380)


# ─────────────────────────────────────────────────────────────────────
# 7. MATERIAL USAGE FREQUENCY
# ─────────────────────────────────────────────────────────────────────
def create_material_usage_chart(df: pd.DataFrame) -> go.Figure:
    """How frequently each material type is used.
    HOW TO READ: Taller bar = more production runs using that material.
    Helps identify the most-consumed materials for procurement planning."""
    materials = {}
    for col_name in ['1st Input Name', '1st Fresh Material', '2nd Fresh Material']:
        if col_name in df.columns:
            counts = df[col_name].dropna().value_counts()
            for mat, count in counts.items():
                if mat and str(mat).strip():
                    materials[str(mat).strip()] = materials.get(str(mat).strip(), 0) + count

    if not materials:
        return go.Figure()

    sorted_mats = sorted(materials.items(), key=lambda x: x[1], reverse=True)[:12]
    names, counts = zip(*sorted_mats)

    fig = go.Figure(go.Bar(
        x=list(names), y=list(counts),
        marker=dict(
            color=list(range(len(counts))),
            colorscale=[[0, COLORS['primary']], [1, COLORS['accent']]],
        ),
        text=list(counts), textposition='auto',
        textfont=dict(color='white', size=11),
        hovertemplate='<b>%{x}</b><br>Usage Count: %{y}<extra></extra>',
    ))
    fig.update_xaxes(tickangle=-45)
    return _apply_layout(fig, "Material Usage Frequency", 400)


# ─────────────────────────────────────────────────────────────────────
# 8. COVERAGE GAUGE
# ─────────────────────────────────────────────────────────────────────
def create_coverage_gauge(stats: dict) -> go.Figure:
    """Data coverage rate - how many cells were auto-filled.
    HOW TO READ: Higher percentage = more complete auto-fill.
    Green zone (>80%) = production-ready. Below = manual review needed."""
    # LAM has 5 fill targets: Fresh1, Fresh2, Adhesive, Hardener, Solvent
    total_possible = stats.get('printing_rows', 0) + stats.get('lam_rows', 0) * 5
    filled = (stats.get('film_filled', 0) + stats.get('fresh1_filled', 0) +
              stats.get('fresh2_filled', 0) + stats.get('adh_filled', 0) +
              stats.get('hard_filled', 0) + stats.get('sol_filled', 0))
    pct = (filled / max(total_possible, 1)) * 100 if total_possible > 0 else 0

    fig = go.Figure(go.Indicator(
        mode="gauge+number",
        value=pct,
        number=dict(suffix="%", font=dict(size=40, color=COLORS['text'])),
        gauge=dict(
            axis=dict(range=[0, 100], tickcolor=COLORS['muted']),
            bar=dict(color=COLORS['primary']),
            bgcolor=COLORS['card'], borderwidth=0,
            steps=[
                dict(range=[0, 50], color='rgba(239, 68, 68, 0.15)'),
                dict(range=[50, 80], color='rgba(245, 158, 11, 0.15)'),
                dict(range=[80, 100], color='rgba(16, 185, 129, 0.15)'),
            ],
            threshold=dict(line=dict(color=COLORS['success'], width=3), thickness=0.8, value=90),
        ),
        title=dict(text="Auto-Fill Coverage", font=dict(size=14, color=COLORS['muted'])),
    ))
    return _apply_layout(fig, None, 280)


# ─────────────────────────────────────────────────────────────────────
# 9. COST PER KG EFFICIENCY
# ─────────────────────────────────────────────────────────────────────
def create_cost_per_kg_chart(df: pd.DataFrame) -> go.Figure:
    """Cost per Kg of output by order - production efficiency metric.
    HOW TO READ: Higher bars = more expensive per Kg output (less efficient).
    Compare orders to spot inefficiencies in material usage."""
    cost_cols = ['Film Value', '1st Fresh Value', '2nd Fresh Value',
                 'Adh Value', 'Hard Value', 'Sol Value']
    available = [c for c in cost_cols if c in df.columns]
    output_col = 'Net Wt. (Kgs-Output)'
    order_col = 'Order No'

    if not available or output_col not in df.columns or order_col not in df.columns:
        return go.Figure()

    df_c = df.copy()
    df_c['Total Cost'] = df_c[available].fillna(0).sum(axis=1)
    df_c['Output'] = pd.to_numeric(df_c[output_col], errors='coerce').fillna(0)

    order_agg = df_c.groupby(order_col).agg({'Total Cost': 'sum', 'Output': 'sum'}).reset_index()
    order_agg = order_agg[(order_agg['Output'] > 0) & (order_agg['Total Cost'] > 0)]
    order_agg['Cost/Kg'] = order_agg['Total Cost'] / order_agg['Output']
    order_agg = order_agg.nlargest(15, 'Cost/Kg').sort_values('Cost/Kg', ascending=True)

    fig = go.Figure(go.Bar(
        x=order_agg['Cost/Kg'], y=order_agg[order_col], orientation='h',
        marker=dict(
            color=order_agg['Cost/Kg'],
            colorscale=[[0, COLORS['success']], [0.5, COLORS['warning']], [1, COLORS['danger']]],
        ),
        text=order_agg['Cost/Kg'].apply(lambda x: f'AED {x:.2f}/Kg'),
        textposition='auto', textfont=dict(color='white', size=10),
        hovertemplate='<b>%{y}</b><br>Cost/Kg: AED %{x:.2f}<br><extra></extra>',
    ))

    avg_cpk = order_agg['Cost/Kg'].mean()
    fig.add_vline(x=avg_cpk, line_dash="dash", line_color=COLORS['warning'],
                  annotation_text=f"Avg: AED {avg_cpk:.2f}/Kg")
    return _apply_layout(fig, "Top 15 Orders - Cost per Kg (Efficiency)", 450)


# ─────────────────────────────────────────────────────────────────────
# 10. CHEMICAL RATES TREND
# ─────────────────────────────────────────────────────────────────────
def create_chemical_rates_summary(stats: dict) -> go.Figure:
    """Chemical rates used this month — Adhesive, Hardener, Solvent.
    HOW TO READ: Shows the rate (AED/Kg) used for each chemical category.
    Compare month-over-month to track procurement cost changes."""
    chemicals = {
        'Adhesive': stats.get('adh_rate', 0),
        'Hardener': stats.get('hard_rate', 0),
        'Solvent': stats.get('sol_rate', 0),
    }
    chemicals = {k: v for k, v in chemicals.items() if v > 0}
    if not chemicals:
        return go.Figure()

    fig = go.Figure(go.Bar(
        x=list(chemicals.keys()), y=list(chemicals.values()),
        marker=dict(color=[COLORS['primary'], COLORS['accent'], COLORS['success']][:len(chemicals)]),
        text=[f'AED {v:.3f}' for v in chemicals.values()],
        textposition='auto', textfont=dict(color='white', size=13),
        hovertemplate='<b>%{x}</b><br>Rate: AED %{y:.4f}/Kg<extra></extra>',
    ))
    return _apply_layout(fig, "Chemical Rates This Month (AED/Kg)", 350)


# ─────────────────────────────────────────────────────────────────────
# COMPUTE STATS
# ─────────────────────────────────────────────────────────────────────
def compute_dashboard_stats(df: pd.DataFrame, fill_stats: dict) -> dict:
    """Compute all KPI values from the filled DataFrame."""
    stats = dict(fill_stats)
    cost_map = {
        'total_film_cost': 'Film Value',
        'total_fresh1_cost': '1st Fresh Value',
        'total_fresh2_cost': '2nd Fresh Value',
        'total_adh_cost': 'Adh Value',
        'total_hard_cost': 'Hard Value',
        'total_sol_cost': 'Sol Value',
    }
    for key, col_name in cost_map.items():
        real_col = _col(df, col_name)
        stats[key] = pd.to_numeric(df[real_col], errors='coerce').fillna(0).sum() if real_col else 0

    stats['total_material_cost'] = sum(stats[k] for k in cost_map)

    # Wastage
    w_col = _col(df, 'Waste')
    stats['total_waste_kg'] = pd.to_numeric(df[w_col], errors='coerce').fillna(0).sum() if w_col else 0

    # Output
    out_col = _col(df, 'Net Wt. (Kgs-Output)')
    stats['total_output_kg'] = pd.to_numeric(df[out_col], errors='coerce').fillna(0).sum() if out_col else 0

    if stats['total_output_kg'] > 0:
        stats['waste_pct'] = (stats['total_waste_kg'] / (stats['total_output_kg'] + stats['total_waste_kg'])) * 100
        stats['cost_per_kg'] = stats['total_material_cost'] / stats['total_output_kg']
    else:
        stats['waste_pct'] = 0
        stats['cost_per_kg'] = 0

    # Rates — use median (resolve column aliases)
    for col_name, key in [('Sol Rate', 'sol_rate'), ('Adh Rate', 'adh_rate'), ('Hard Rate', 'hard_rate')]:
        real_col = _col(df, col_name)
        if real_col:
            vals = pd.to_numeric(df[real_col], errors='coerce').dropna()
            vals = vals[vals > 0]
            stats[key] = vals.median() if not vals.empty else 0
        else:
            stats[key] = 0

    # Order count
    if 'Order No' in df.columns:
        stats['unique_orders'] = df['Order No'].nunique()
    else:
        stats['unique_orders'] = 0

    return stats


# ─────────────────────────────────────────────────────────────────────
# 11. AI-GENERATED BUSINESS INSIGHTS
# ─────────────────────────────────────────────────────────────────────
def generate_ai_insights(df: pd.DataFrame, stats: dict) -> list:
    """Auto-generate actionable business insights from the data."""
    insights = []
    total_cost = stats.get('total_material_cost', 0)
    if total_cost <= 0:
        return insights

    # Biggest cost driver
    cost_items = [
        ('Film', stats.get('total_film_cost', 0)),
        ('1st Fresh', stats.get('total_fresh1_cost', 0)),
        ('2nd Fresh', stats.get('total_fresh2_cost', 0)),
        ('Adhesive', stats.get('total_adh_cost', 0)),
        ('Hardener', stats.get('total_hard_cost', 0)),
        ('Solvent', stats.get('total_sol_cost', 0)),
    ]
    cost_items.sort(key=lambda x: x[1], reverse=True)
    top = cost_items[0]
    if top[1] > 0:
        pct = top[1] / total_cost * 100
        insights.append({
            'icon': '💰', 'type': 'cost',
            'title': f'{top[0]} is the #1 cost driver',
            'detail': f'AED {top[1]:,.0f} ({pct:.0f}% of total material spend). Focus procurement negotiations here for maximum savings impact.',
        })

    # Cost per Kg analysis
    cpk = stats.get('cost_per_kg', 0)
    if cpk > 0:
        insights.append({
            'icon': '⚡', 'type': 'efficiency',
            'title': f'Material cost efficiency: AED {cpk:.2f}/Kg',
            'detail': f'Every kilogram of output costs AED {cpk:.2f} in raw materials. Track this monthly to detect cost creep.',
        })

    # Wastage alert
    waste_pct = stats.get('waste_pct', 0)
    if waste_pct > 5:
        insights.append({
            'icon': '🚨', 'type': 'alert',
            'title': f'Wastage at {waste_pct:.1f}% — exceeds 5% target',
            'detail': f'{stats.get("total_waste_kg", 0):,.0f} Kg wasted. At current rates, this equals ~AED {stats.get("total_waste_kg", 0) * cpk:,.0f} in lost material value.',
        })
    elif waste_pct > 0:
        insights.append({
            'icon': '✅', 'type': 'success',
            'title': f'Wastage at {waste_pct:.1f}% — within target',
            'detail': f'Production waste is under the 5% benchmark. Good operational discipline.',
        })

    # Top expensive order
    cost_cols = ['Film Value', '1st Fresh Value', '2nd Fresh Value',
                 'Adh Value', 'Hard Value', 'Sol Value']
    available = [c for c in cost_cols if c in df.columns]
    if available and 'Order No' in df.columns:
        df_c = df.copy()
        df_c['_total'] = df_c[available].fillna(0).sum(axis=1)
        order_cost = df_c.groupby('Order No')['_total'].sum()
        if not order_cost.empty:
            top_order = order_cost.idxmax()
            top_val = order_cost.max()
            pct = top_val / total_cost * 100
            insights.append({
                'icon': '📋', 'type': 'order',
                'title': f'Highest cost order: {top_order}',
                'detail': f'AED {top_val:,.0f} ({pct:.0f}% of total). Review this order for material optimization opportunities.',
            })

    # Printing vs LAM split
    film_cost = stats.get('total_film_cost', 0)
    lam_cost = sum(stats.get(k, 0) for k in ['total_fresh1_cost', 'total_fresh2_cost',
                                                'total_adh_cost', 'total_hard_cost', 'total_sol_cost'])
    if film_cost > 0 and lam_cost > 0:
        ratio = film_cost / lam_cost
        insights.append({
            'icon': '🏭', 'type': 'process',
            'title': f'Printing vs LAM cost ratio: {ratio:.1f}x',
            'detail': f'Printing: AED {film_cost:,.0f} | LAM: AED {lam_cost:,.0f}. {"Printing dominates cost — focus optimization there." if ratio > 1 else "LAM is the bigger cost center — review fresh material and chemical spending."}',
        })

    # Coverage / automation
    total_possible = stats.get('printing_rows', 0) + stats.get('lam_rows', 0) * 5
    filled = sum(stats.get(k, 0) for k in ['film_filled', 'fresh1_filled', 'fresh2_filled',
                                             'adh_filled', 'hard_filled', 'sol_filled'])
    if total_possible > 0:
        auto_pct = filled / total_possible * 100
        insights.append({
            'icon': '🤖', 'type': 'automation',
            'title': f'Automation coverage: {auto_pct:.0f}%',
            'detail': f'{filled}/{total_possible} cells auto-filled. {"Excellent — production ready." if auto_pct >= 90 else "Some rows need manual review — check the Processing Log tab."}',
        })

    return insights


# ─────────────────────────────────────────────────────────────────────
# 12. ACTIONABLE RECOMMENDATIONS
# ─────────────────────────────────────────────────────────────────────
def generate_recommendations(df: pd.DataFrame, stats: dict) -> list:
    """Generate prioritized action items for the CEO/Costing team."""
    recs = []
    total_cost = stats.get('total_material_cost', 0)
    if total_cost <= 0:
        return recs

    cost_cols = ['Film Value', '1st Fresh Value', '2nd Fresh Value',
                 'Adh Value', 'Hard Value', 'Sol Value']
    available = [c for c in cost_cols if c in df.columns]
    out_col = 'Net Wt. (Kgs-Output)'

    # Rec 1: Material with highest spend
    cost_items = {
        'Film': stats.get('total_film_cost', 0),
        '1st Fresh': stats.get('total_fresh1_cost', 0),
        '2nd Fresh': stats.get('total_fresh2_cost', 0),
        'Adhesive': stats.get('total_adh_cost', 0),
    }
    top_mat = max(cost_items, key=cost_items.get)
    if cost_items[top_mat] > 0:
        recs.append({
            'priority': 'HIGH', 'category': 'Procurement',
            'action': f'Negotiate volume discounts for {top_mat} material',
            'impact': f'AED {cost_items[top_mat]:,.0f} ({cost_items[top_mat]/total_cost*100:.0f}% of spend)',
            'detail': f'{top_mat} is the largest material cost. Even a 5% rate reduction saves AED {cost_items[top_mat]*0.05:,.0f}/month.',
        })

    # Rec 2: Inefficient orders
    if available and out_col in df.columns and 'Order No' in df.columns:
        df_c = df.copy()
        df_c['_total'] = df_c[available].fillna(0).sum(axis=1)
        df_c['_output'] = pd.to_numeric(df_c[out_col], errors='coerce').fillna(0)
        order_agg = df_c.groupby('Order No').agg({'_total': 'sum', '_output': 'sum'}).reset_index()
        order_agg = order_agg[(order_agg['_output'] > 0) & (order_agg['_total'] > 0)]
        if not order_agg.empty:
            order_agg['cpk'] = order_agg['_total'] / order_agg['_output']
            avg_cpk = order_agg['cpk'].mean()
            expensive = order_agg[order_agg['cpk'] > avg_cpk * 1.5]
            if not expensive.empty:
                worst = expensive.nlargest(1, 'cpk').iloc[0]
                recs.append({
                    'priority': 'HIGH', 'category': 'Production',
                    'action': f'Investigate Order {worst["Order No"]} — cost/kg {worst["cpk"]:.2f} vs avg {avg_cpk:.2f}',
                    'impact': f'{worst["cpk"]/avg_cpk*100-100:.0f}% above average efficiency',
                    'detail': f'{len(expensive)} orders have cost/kg more than 50% above average. Root causes may include material waste, rate outliers, or small batch sizes.',
                })

    # Rec 3: Wastage
    waste_pct = stats.get('waste_pct', 0)
    if waste_pct > 5:
        recs.append({
            'priority': 'MEDIUM', 'category': 'Operations',
            'action': f'Reduce wastage from {waste_pct:.1f}% to below 5% target',
            'impact': f'Could save ~AED {stats.get("total_waste_kg",0) * stats.get("cost_per_kg",0) * 0.5:,.0f}',
            'detail': 'Review machine calibration, operator procedures, and material quality for root cause.',
        })

    # Rec 4: Rate variance
    for col, label in [('Film Rate', 'Film'), ('1st Fresh Rate', 'Fresh')]:
        if col in df.columns:
            rates = pd.to_numeric(df[col], errors='coerce').dropna()
            rates = rates[rates > 0]
            if len(rates) > 3:
                cv = rates.std() / rates.mean() * 100
                if cv > 15:
                    recs.append({
                        'priority': 'MEDIUM', 'category': 'Procurement',
                        'action': f'Investigate {label} rate variance (CV={cv:.0f}%)',
                        'impact': f'Range: AED {rates.min():.2f} - {rates.max():.2f}',
                        'detail': f'High price dispersion suggests inconsistent supplier pricing or mixed material grades. Standardize procurement.',
                    })

    return recs


# ─────────────────────────────────────────────────────────────────────
# 13. MATERIAL COST WATERFALL
# ─────────────────────────────────────────────────────────────────────
def create_material_cost_waterfall(stats: dict) -> go.Figure:
    """Waterfall showing how each material builds up total cost."""
    items = [
        ('Film', stats.get('total_film_cost', 0)),
        ('1st Fresh', stats.get('total_fresh1_cost', 0)),
        ('2nd Fresh', stats.get('total_fresh2_cost', 0)),
        ('Adhesive', stats.get('total_adh_cost', 0)),
        ('Hardener', stats.get('total_hard_cost', 0)),
        ('Solvent', stats.get('total_sol_cost', 0)),
    ]
    items = [(l, v) for l, v in items if v > 0]
    if not items:
        return go.Figure()

    labels = [l for l, _ in items] + ['Total']
    values = [v for _, v in items] + [sum(v for _, v in items)]
    measures = ['relative'] * len(items) + ['total']

    fig = go.Figure(go.Waterfall(
        x=labels, y=values, measure=measures,
        connector=dict(line=dict(color='rgba(99,102,241,0.3)', width=1)),
        increasing=dict(marker=dict(color='#6366F1')),
        totals=dict(marker=dict(color='#059669')),
        texttemplate='AED %{y:,.0f}', textposition='outside',
        textfont=dict(size=11, color=COLORS['text']),
        hovertemplate='<b>%{x}</b><br>Cost: AED %{y:,.0f}<extra></extra>',
    ))
    return _apply_layout(fig, 'Cost Build-Up (Waterfall)', 420)


# ─────────────────────────────────────────────────────────────────────
# 14. EFFICIENCY SCATTER
# ─────────────────────────────────────────────────────────────────────
def create_efficiency_scatter(df: pd.DataFrame) -> go.Figure:
    """Scatter: Output vs Cost/Kg per order. Size = total cost."""
    cost_cols = ['Film Value', '1st Fresh Value', '2nd Fresh Value',
                 'Adh Value', 'Hard Value', 'Sol Value']
    available = [c for c in cost_cols if c in df.columns]
    out_col = 'Net Wt. (Kgs-Output)'
    order_col = 'Order No'

    if not available or out_col not in df.columns or order_col not in df.columns:
        return go.Figure()

    df_c = df.copy()
    df_c['_total'] = df_c[available].fillna(0).sum(axis=1)
    df_c['_output'] = pd.to_numeric(df_c[out_col], errors='coerce').fillna(0)
    oa = df_c.groupby(order_col).agg({'_total': 'sum', '_output': 'sum'}).reset_index()
    oa = oa[(oa['_output'] > 0) & (oa['_total'] > 0)]
    if oa.empty:
        return go.Figure()
    oa['cpk'] = oa['_total'] / oa['_output']

    avg_cpk = oa['cpk'].mean()

    fig = go.Figure(go.Scatter(
        x=oa['_output'], y=oa['cpk'],
        mode='markers', text=oa[order_col],
        marker=dict(
            size=np.clip(oa['_total'] / oa['_total'].max() * 40, 8, 50),
            color=oa['cpk'],
            colorscale=[[0, '#10B981'], [0.5, '#F59E0B'], [1, '#EF4444']],
            line=dict(width=1, color='white'),
            opacity=0.85,
        ),
        hovertemplate='<b>%{text}</b><br>Output: %{x:,.0f} Kg<br>Cost/Kg: AED %{y:.2f}<br><extra></extra>',
    ))
    fig.add_hline(y=avg_cpk, line_dash='dash', line_color=COLORS['warning'],
                  annotation_text=f'Avg: AED {avg_cpk:.2f}/Kg')
    fig.update_xaxes(title_text='Output (Kg)')
    fig.update_yaxes(title_text='Cost per Kg (AED)')
    return _apply_layout(fig, 'Order Efficiency Map', 450)


# ─────────────────────────────────────────────────────────────────────
# 15. COST PARETO (80/20)
# ─────────────────────────────────────────────────────────────────────
def create_cost_pareto(df: pd.DataFrame) -> go.Figure:
    """Pareto chart: which orders drive 80% of cost."""
    cost_cols = ['Film Value', '1st Fresh Value', '2nd Fresh Value',
                 'Adh Value', 'Hard Value', 'Sol Value']
    available = [c for c in cost_cols if c in df.columns]
    if not available or 'Order No' not in df.columns:
        return go.Figure()

    df_c = df.copy()
    df_c['_total'] = df_c[available].fillna(0).sum(axis=1)
    oc = df_c.groupby('Order No')['_total'].sum().sort_values(ascending=False).reset_index()
    oc = oc[oc['_total'] > 0]
    if oc.empty:
        return go.Figure()

    oc['cumulative'] = oc['_total'].cumsum()
    oc['cum_pct'] = oc['cumulative'] / oc['_total'].sum() * 100

    fig = make_subplots(specs=[[{"secondary_y": True}]])
    fig.add_trace(go.Bar(
        x=oc['Order No'], y=oc['_total'], name='Cost',
        marker=dict(color=CHART_COLORS[0]),
        hovertemplate='<b>%{x}</b><br>Cost: AED %{y:,.0f}<extra></extra>',
    ), secondary_y=False)
    fig.add_trace(go.Scatter(
        x=oc['Order No'], y=oc['cum_pct'], name='Cumulative %',
        mode='lines+markers', line=dict(color=COLORS['accent'], width=2),
        marker=dict(size=4),
        hovertemplate='<b>%{x}</b><br>Cumulative: %{y:.0f}%<extra></extra>',
    ), secondary_y=True)
    fig.add_hline(y=80, line_dash='dot', line_color=COLORS['danger'],
                  annotation_text='80% threshold', secondary_y=True)
    fig.update_yaxes(title_text='Cost (AED)', secondary_y=False)
    fig.update_yaxes(title_text='Cumulative %', range=[0, 105], secondary_y=True)
    fig.update_xaxes(tickangle=-45)
    return _apply_layout(fig, 'Pareto Analysis — 80/20 Rule', 420)


# ─────────────────────────────────────────────────────────────────────
# 16. PROCESS EFFICIENCY RADAR
# ─────────────────────────────────────────────────────────────────────
def create_process_efficiency_radar(df: pd.DataFrame, stats: dict) -> go.Figure:
    """Radar comparing Printing vs LAM across multiple dimensions."""
    process_col = 'Process'
    if process_col not in df.columns:
        return go.Figure()

    cost_cols = ['Film Value', '1st Fresh Value', '2nd Fresh Value',
                 'Adh Value', 'Hard Value', 'Sol Value']
    available = [c for c in cost_cols if c in df.columns]
    out_col = 'Net Wt. (Kgs-Output)'

    categories = ['Volume (Rows)', 'Total Cost', 'Avg Cost/Row', 'Output (Kg)', 'Fill Rate %']
    printing_vals = [0] * 5
    lam_vals = [0] * 5

    for proc, vals in [('PRINTING', printing_vals), ('LAM', lam_vals)]:
        subset = df[df[process_col].astype(str).str.upper().str.strip() == proc]
        if subset.empty:
            continue
        vals[0] = len(subset)
        total_c = subset[available].fillna(0).sum().sum() if available else 0
        vals[1] = total_c
        vals[2] = total_c / max(len(subset), 1)
        vals[3] = pd.to_numeric(subset[out_col], errors='coerce').fillna(0).sum() if out_col in subset.columns else 0

    # Fill rates
    p_rows = stats.get('printing_rows', 0)
    l_rows = stats.get('lam_rows', 0)
    printing_vals[4] = stats.get('film_filled', 0) / max(p_rows, 1) * 100
    lam_filled = sum(stats.get(k, 0) for k in ['fresh1_filled', 'fresh2_filled',
                                                  'adh_filled', 'hard_filled', 'sol_filled'])
    lam_vals[4] = lam_filled / max(l_rows * 5, 1) * 100

    # Normalize to 0-100 scale
    max_vals = [max(printing_vals[i], lam_vals[i], 1) for i in range(5)]
    p_norm = [printing_vals[i] / max_vals[i] * 100 for i in range(5)]
    l_norm = [lam_vals[i] / max_vals[i] * 100 for i in range(5)]

    fig = go.Figure()
    fig.add_trace(go.Scatterpolar(
        r=p_norm + [p_norm[0]], theta=categories + [categories[0]],
        fill='toself', name='Printing',
        fillcolor='rgba(99, 102, 241, 0.15)',
        line=dict(color='#6366F1', width=2),
    ))
    fig.add_trace(go.Scatterpolar(
        r=l_norm + [l_norm[0]], theta=categories + [categories[0]],
        fill='toself', name='LAM',
        fillcolor='rgba(236, 72, 153, 0.15)',
        line=dict(color='#EC4899', width=2),
    ))
    fig.update_layout(polar=dict(
        radialaxis=dict(visible=True, range=[0, 110], showticklabels=False),
        bgcolor='rgba(0,0,0,0)',
    ))
    return _apply_layout(fig, 'Process Comparison Radar', 420)
