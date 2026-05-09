rows = {
    42: {85588: 1857.7, 85547: 1809.5, 85572: 991.2, 85157: 497.8, 85226: 305.0},
    45: {85573: 2621.3, 85330: 253.0, 85157: 180.0},
    46: {85572: 376.8, 85226: 200.0},
}
rates = {85588: 4.22, 85547: 4.22, 85572: 4.588, 85157: 4.404, 85226: 4.22, 85573: 4.588, 85330: 4.22}

for row, mrrs in rows.items():
    total = sum(mrrs.values())
    threshold = total * 0.10
    print(f"\nRow {row}: total={total:.1f}, 10% threshold={threshold:.1f}")
    dominant = {}
    for m, q in mrrs.items():
        pct = q/total*100
        passes = q >= threshold
        status = "PASS" if passes else "FAIL"
        print(f"  MRR {m}: {q:.1f}kg ({pct:.1f}%) -> {status}")
        if passes:
            dominant[m] = q
    
    # What the 10% filter produces
    if dominant:
        w = sum(rates[m] * q for m, q in dominant.items()) / sum(dominant.values())
        print(f"  10% filter rate: {w:.6f}")
    
    # What ALL MRRs produce
    w_all = sum(rates[m] * q for m, q in mrrs.items()) / total
    print(f"  ALL MRRs rate:   {w_all:.6f}")
