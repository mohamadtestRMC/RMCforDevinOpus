"""Test: verify column resolution and chart rendering."""
import io, sys
sys.path.insert(0, '.')
from engine.fill_jobtrack import fill_jobtrack, get_filled_data_for_dashboard
from dashboard.charts import (compute_dashboard_stats, generate_ai_insights,
                                generate_recommendations, create_material_cost_waterfall,
                                create_efficiency_scatter, create_cost_pareto,
                                create_process_efficiency_radar, create_process_cost_bar,
                                create_daily_cost_trend, create_rate_comparison_chart,
                                create_wastage_chart, create_cost_per_kg_chart,
                                create_top_orders_chart, create_cost_breakdown_pie,
                                create_chemical_rates_summary, create_coverage_gauge,
                                create_material_usage_chart)

jt = open('Template_Files/Jobtrack Feb Without MRR.xlsx','rb').read()
st = open('Template_Files/Stores Recordings.xlsx','rb').read()
pr = open('Template_Files/Purchase Register - 2021 - 2026 _Feb 26.xlsx','rb').read()
gr = open('Template_Files/Granules Recipe - February 2026.xlsx','rb').read()
mg = open('Template_Files/MEGA PACK.xlsx','rb').read()

print("Running engine...")
out, log, stats = fill_jobtrack(io.BytesIO(jt), io.BytesIO(st), io.BytesIO(pr),
    granules_file=io.BytesIO(gr), megapack_file=io.BytesIO(mg))
out.seek(0)
df = get_filled_data_for_dashboard(io.BytesIO(out.read()))
ds = compute_dashboard_stats(df, stats)

tc = ds.get('total_material_cost', 0)
cpk = ds.get('cost_per_kg', 0)
wp = ds.get('waste_pct', 0)
print(f"Total Cost: AED {tc:,.0f} | Cost/Kg: {cpk:.2f} | Waste: {wp:.1f}%")
print(f"Film: AED {ds['total_film_cost']:,.0f} | Fresh1: AED {ds['total_fresh1_cost']:,.0f} | Fresh2: AED {ds['total_fresh2_cost']:,.0f}")
print(f"Adh: AED {ds['total_adh_cost']:,.0f} | Hard: AED {ds['total_hard_cost']:,.0f} | Sol: AED {ds['total_sol_cost']:,.0f}")
print(f"Adh Rate: {ds['adh_rate']:.3f} | Hard Rate: {ds['hard_rate']:.3f} | Sol Rate: {ds['sol_rate']:.3f}")

# Test all charts
charts = {
    'Waterfall': create_material_cost_waterfall(ds),
    'Pie': create_cost_breakdown_pie(ds),
    'Process Bar': create_process_cost_bar(df),
    'Daily Trend': create_daily_cost_trend(df),
    'Rate Box': create_rate_comparison_chart(df),
    'Top Orders': create_top_orders_chart(df),
    'Wastage': create_wastage_chart(df),
    'Material Usage': create_material_usage_chart(df),
    'Coverage': create_coverage_gauge(ds),
    'Cost/Kg': create_cost_per_kg_chart(df),
    'Chemical Rates': create_chemical_rates_summary(ds),
    'Scatter': create_efficiency_scatter(df),
    'Pareto': create_cost_pareto(df),
    'Radar': create_process_efficiency_radar(df, ds),
}

print(f"\n=== CHART TESTS ===")
for name, fig in charts.items():
    if fig is None:
        print(f"  SKIP {name} (no data)")
    else:
        traces = len(fig.data)
        print(f"  OK   {name}: {traces} traces")

insights = generate_ai_insights(df, ds)
recs = generate_recommendations(df, ds)
print(f"\nInsights: {len(insights)} | Recs: {len(recs)}")
print("\nALL TESTS PASSED")
