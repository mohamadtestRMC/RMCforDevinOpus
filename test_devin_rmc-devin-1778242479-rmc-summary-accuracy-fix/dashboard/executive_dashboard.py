"""
Executive Dashboard — CEO-level KPIs, charts, and AI insights for RMC reporting.
"""
from __future__ import annotations
import logging
import pandas as pd
import numpy as np
from typing import Dict, List, Any

logger = logging.getLogger(__name__)

COLORS = ["#6366F1","#8B5CF6","#3B82F6","#10B981","#F59E0B",
          "#EF4444","#06B6D4","#EC4899","#F97316","#84CC16"]

VAL_COLS = {
    'OPN_WIP_Val': 'Opening WIP',
    'Print_Film_Val': 'Print Film',
    'Lam_Fresh_Val': 'Lam Fresh',
    'Slit_Val': 'Other Film',
    'Ink_Val': 'Ink & Solvent',
    'Lam_Chem_Val': 'Adh+Hard+Sol',
    'BP_SV_Val': 'Zipper/Valve',
    'CLS_WIP_Val': 'CLS WIP (-)',
}


def compute_kpis(df: pd.DataFrame) -> Dict[str, Any]:
    """Compute executive KPIs from RMC Summary DataFrame."""
    if df.empty:
        return {}
    total_cost = df['Total_Cost'].sum() if 'Total_Cost' in df.columns else 0
    total_output = df['FG_Output'].sum() if 'FG_Output' in df.columns else 0
    avg_rmc = total_cost / total_output if total_output > 0 else 0

    # Wastage
    waste_cols = [c for c in df.columns if 'wastage' in c.lower() or 'waste' in c.lower()]
    total_waste_val = sum(df[c].sum() for c in waste_cols if df[c].dtype in ['float64','int64'])

    # Coverage
    orders_with_rmc = (df['RMC_Kg'] > 0).sum() if 'RMC_Kg' in df.columns else 0
    coverage = orders_with_rmc / len(df) * 100 if len(df) > 0 else 0

    return {
        'total_cost': total_cost,
        'total_output': total_output,
        'avg_rmc': avg_rmc,
        'order_count': len(df),
        'orders_with_rmc': int(orders_with_rmc),
        'coverage_pct': coverage,
        'total_waste_val': total_waste_val,
        'waste_pct': total_waste_val / total_cost * 100 if total_cost > 0 else 0,
    }


def compute_cost_breakdown(df: pd.DataFrame) -> Dict[str, float]:
    """Compute cost breakdown by category."""
    breakdown = {}
    for col, label in VAL_COLS.items():
        if col in df.columns:
            val = df[col].sum()
            if val > 0 and 'CLS' not in col:
                breakdown[label] = val
    return breakdown


def compute_top_orders(df: pd.DataFrame, n: int = 15) -> pd.DataFrame:
    """Get top N orders by total cost."""
    if 'Total_Cost' not in df.columns or 'Order' not in df.columns:
        return pd.DataFrame()
    return df.nlargest(n, 'Total_Cost')[['Order', 'Total_Cost', 'FG_Output', 'RMC_Kg']].copy()


def compute_rmc_distribution(df: pd.DataFrame) -> Dict[str, int]:
    """Bucket orders by RMC/Kg ranges."""
    if 'RMC_Kg' not in df.columns:
        return {}
    valid = df[df['RMC_Kg'] > 0]['RMC_Kg']
    if valid.empty:
        return {}
    buckets = {'<50': 0, '50-100': 0, '100-200': 0, '200-500': 0, '>500': 0}
    for v in valid:
        if v < 50: buckets['<50'] += 1
        elif v < 100: buckets['50-100'] += 1
        elif v < 200: buckets['100-200'] += 1
        elif v < 500: buckets['200-500'] += 1
        else: buckets['>500'] += 1
    return buckets


def generate_insights(df: pd.DataFrame, kpis: Dict) -> List[str]:
    """Generate AI-style business insights from data."""
    insights = []
    if not kpis:
        return insights

    tc = kpis.get('total_cost', 0)
    to = kpis.get('total_output', 0)
    ar = kpis.get('avg_rmc', 0)
    cov = kpis.get('coverage_pct', 0)

    insights.append(f"📊 Total material cost is AED {tc:,.0f} across {kpis['order_count']} orders "
                    f"with average RMC of AED {ar:,.2f}/Kg")

    if cov < 100:
        insights.append(f"⚠️ Coverage is {cov:.1f}% — {kpis['order_count'] - kpis['orders_with_rmc']} "
                        f"orders have no computed RMC/Kg")

    # Top cost drivers
    if 'Total_Cost' in df.columns and 'Order' in df.columns:
        top3 = df.nlargest(3, 'Total_Cost')
        top3_pct = top3['Total_Cost'].sum() / tc * 100 if tc > 0 else 0
        insights.append(f"🔝 Top 3 orders account for {top3_pct:.1f}% of total cost")

    # Wastage
    wp = kpis.get('waste_pct', 0)
    if wp > 5:
        insights.append(f"🔴 Wastage is {wp:.1f}% of total cost — consider process optimization")
    elif wp > 0:
        insights.append(f"✅ Wastage is {wp:.1f}% — within acceptable range")

    return insights


def generate_recommendations(df: pd.DataFrame, kpis: Dict) -> List[Dict]:
    """Generate actionable recommendations."""
    recs = []
    if not kpis:
        return recs

    if kpis.get('coverage_pct', 0) < 95:
        recs.append({
            'priority': 'HIGH',
            'title': 'Improve RMC Coverage',
            'detail': f"Only {kpis['coverage_pct']:.1f}% of orders have computed RMC. "
                      f"Review MRR lookup logic for the remaining "
                      f"{kpis['order_count'] - kpis['orders_with_rmc']} orders.",
            'icon': '🎯'
        })

    if kpis.get('waste_pct', 0) > 8:
        recs.append({
            'priority': 'HIGH',
            'title': 'Reduce Material Wastage',
            'detail': f"Wastage is {kpis['waste_pct']:.1f}% of cost. Focus on Print and Lam "
                      f"processes which typically have the highest waste.",
            'icon': '♻️'
        })

    if 'RMC_Kg' in df.columns:
        valid = df[df['RMC_Kg'] > 0]['RMC_Kg']
        if not valid.empty:
            std = valid.std()
            mean = valid.mean()
            cv = std / mean if mean > 0 else 0
            if cv > 0.5:
                recs.append({
                    'priority': 'MEDIUM',
                    'title': 'High RMC Variability',
                    'detail': f"RMC/Kg varies significantly (CV={cv:.2f}). "
                              f"Investigate outlier orders for pricing inconsistencies.",
                    'icon': '📈'
                })

    recs.append({
        'priority': 'LOW',
        'title': 'Validate Against Previous Month',
        'detail': 'Upload previous month data for cross-month validation '
                  'to ensure rate consistency and catch data entry errors.',
        'icon': '📆'
    })

    return recs
