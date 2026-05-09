"""
Rate Lookup Engine - v2
Finds material rates from the Purchase Register by matching
Tracking Number (=MRR), Material, Size, and Mic.
Supports qty-weighted average rate calculation.
"""
import pandas as pd
import logging

logger = logging.getLogger(__name__)

# Known Adhesive -> Hardener pairings (legacy, kept for backward compat)
ADH_HARDENER_PAIRS = {
    'MB655': 'CT85',
    'CT85': 'MB655',
    '75-300': 'CR 84',
    'MF 75-300': 'CR 84',
    'CR 84': 'MF 75-300',
    'S110': 'S621',
    'S621': 'S110',
    'S310': 'S631',
    'S631': 'S310',
    'SL816': 'SL720',
    'SL720': 'SL816',
    'LG 59A': 'CAT B',
    'CAT B': 'LG 59A',
    'CR 800-300': 'MF 75-300',
}

# Hardener name aliases: Jobtrack DA column name -> Purchase Register name
HARDENER_NAME_MAP = {
    'CR84': 'CR 84',
    'CR 88-300': 'CR 800-300',
}

# Material aliases for Purchase Register matching
PR_MAT_ALIASES = {
    'WPE': ['WPE', 'PE WHITE', 'WLDPE'],
    'WLDPE': ['WLDPE', 'PE WHITE', 'WPE'],
    'PTD WPE': ['WPE', 'PE WHITE', 'WLDPE'],
    'MATTE TOPP': ['MATTE TOPP', 'MATTE OPP'],
    'MET PET': ['MET PET', 'MET PET UPF', 'MET PET HIGH OD'],
    'PET': ['PET', 'PET CHEM', 'PET UPF'],
}


def load_purchase_register(file) -> pd.DataFrame:
    """Load and parse the Purchase Register Excel file."""
    df = pd.read_excel(file, sheet_name=0, header=2)
    col_map = {c: str(c).strip().replace('\n', ' ').replace('\r', '') for c in df.columns}
    df.rename(columns=col_map, inplace=True)
    logger.info(f"Loaded Purchase Register: {len(df)} rows")
    return df


def _find_col(df, *keywords):
    """Find column by keyword matching."""
    for c in df.columns:
        cl = str(c).lower().strip()
        for kw in keywords:
            if kw in cl:
                return c
    return None


def _get_rate_for_mrr(pr_df, tracking_col, rate_col, material_col, size_col, mic_col,
                      mrr_val, material=None, size=None, mic=None):
    """Get rate for a single MRR from Purchase Register."""
    mask = pd.to_numeric(pr_df[tracking_col], errors='coerce') == mrr_val
    filtered = pr_df[mask]

    if filtered.empty:
        return 0.0

    # Narrow by material — exact match first, then fallback to fuzzy
    if material and material_col:
        mat_upper = str(material).strip().upper()
        aliases = PR_MAT_ALIASES.get(mat_upper, [mat_upper])

        # Pass 1: Exact match on material name or aliases
        def exact_match(x):
            xu = str(x).strip().upper()
            if xu == mat_upper:
                return True
            for alias in aliases:
                if xu == alias:
                    return True
            return False

        exact_filtered = filtered[filtered[material_col].apply(exact_match)]
        if not exact_filtered.empty:
            filtered = exact_filtered
        else:
            # Pass 2: Substring match (but avoid 'PET' matching 'MET PET')
            def fuzzy_match(x):
                xu = str(x).strip().upper()
                # Must start with the material name, or contain it as full word
                for alias in aliases:
                    if alias == xu:
                        return True
                    # Only match if alias IS the material (not a substring of different material)
                    if xu.startswith(alias + ' ') or xu.endswith(' ' + alias):
                        return True
                return False

            fuzzy_filtered = filtered[filtered[material_col].apply(fuzzy_match)]
            if not fuzzy_filtered.empty:
                filtered = fuzzy_filtered

    # Narrow by mic
    if mic is not None and not pd.isna(mic) and mic_col:
        try:
            mic_val = float(mic)
            mic_mask = pd.to_numeric(filtered[mic_col], errors='coerce') == mic_val
            mic_filtered = filtered[mic_mask]
            if not mic_filtered.empty:
                filtered = mic_filtered
        except (ValueError, TypeError):
            pass

    # Narrow by size (with tolerance +/-5)
    if size is not None and not pd.isna(size) and size_col:
        try:
            size_val = float(size)
            sizes_num = pd.to_numeric(filtered[size_col], errors='coerce')
            exact = filtered[sizes_num == size_val]
            if not exact.empty:
                filtered = exact
            else:
                close = filtered[(sizes_num >= size_val - 5) & (sizes_num <= size_val + 5)]
                if not close.empty:
                    filtered = close
        except (ValueError, TypeError):
            pass

    rate_values = pd.to_numeric(filtered[rate_col], errors='coerce').dropna()
    if not rate_values.empty:
        return rate_values.iloc[0]
    return 0.0


def lookup_film_rate(pr_df: pd.DataFrame, mrr_numbers: list,
                     material: str = None, size=None, mic=None) -> float:
    """
    Look up the rate for a film material from Purchase Register.
    Simple average across MRR rates (all MRRs should be pre-filtered to dominant ones).
    """
    tracking_col = _find_col(pr_df, 'tracking')
    material_col = _find_col(pr_df, 'material')
    size_col = _find_col(pr_df, 'size')
    mic_col = _find_col(pr_df, 'mic')
    rate_col = 'Rate'
    rate_cols = [c for c in pr_df.columns if str(c).strip().lower() == 'rate']
    if rate_cols:
        rate_col = rate_cols[0]

    if not tracking_col:
        return 0.0

    rates = []
    for mrr in mrr_numbers:
        try:
            mrr_val = int(float(mrr))
        except (ValueError, TypeError):
            continue
        rate = _get_rate_for_mrr(pr_df, tracking_col, rate_col, material_col,
                                 size_col, mic_col, mrr_val, material, size, mic)
        if rate > 0:
            rates.append(rate)

    if not rates:
        return 0.0
    if len(set(rates)) == 1:
        return rates[0]
    return sum(rates) / len(rates)


def lookup_film_rate_weighted(pr_df: pd.DataFrame, mrr_qty_dict: dict,
                               material: str = None, size=None, mic=None) -> float:
    """
    Look up rate with QTY-WEIGHTED average.
    mrr_qty_dict: {mrr_number: total_issue_qty} from Stores
    
    Rate for each MRR is looked up from PR (per-MRR rate),
    then weighted by the STORES issue qty to produce the final rate.
    """
    tracking_col = _find_col(pr_df, 'tracking')
    material_col = _find_col(pr_df, 'material')
    size_col = _find_col(pr_df, 'size')
    mic_col = _find_col(pr_df, 'mic')
    rate_col = 'Rate'
    rate_cols = [c for c in pr_df.columns if str(c).strip().lower() == 'rate']
    if rate_cols:
        rate_col = rate_cols[0]

    if not tracking_col:
        return 0.0

    # Get per-MRR rate from PR
    mrr_rates = {}  # {mrr: (rate, stores_qty)}
    
    for mrr, qty in mrr_qty_dict.items():
        try:
            mrr_val = int(float(mrr))
        except (ValueError, TypeError):
            continue
        
        # Try with size filter first
        rate = _get_rate_for_mrr(pr_df, tracking_col, rate_col, material_col,
                                 size_col, mic_col, mrr_val, material, size, mic)
        if rate > 0:
            mrr_rates[mrr_val] = (rate, qty)

    # If no MRRs matched with size, try without size  
    if not mrr_rates:
        for mrr, qty in mrr_qty_dict.items():
            try:
                mrr_val = int(float(mrr))
            except (ValueError, TypeError):
                continue
            rate = _get_rate_for_mrr(pr_df, tracking_col, rate_col, material_col,
                                     size_col, mic_col, mrr_val, material, None, mic)
            if rate > 0:
                mrr_rates[mrr_val] = (rate, qty)

    if not mrr_rates:
        return 0.0

    # If all rates are the same, return that rate
    unique_rates = set(r for r, _ in mrr_rates.values())
    if len(unique_rates) == 1:
        return unique_rates.pop()

    # Qty-weighted average using stores issue qty (ALL MRRs, no filtering)
    weighted_sum = sum(rate * qty for rate, qty in mrr_rates.values())
    total_qty = sum(qty for _, qty in mrr_rates.values())
    
    if total_qty > 0:
        return weighted_sum / total_qty
    return 0.0


def lookup_material_rate_for_month(pr_df: pd.DataFrame, material: str,
                                    mic=None, report_month: str = None) -> float:
    """Fallback: get the qty-weighted average rate for a material from ALL PR entries
    in the reporting month. Used when MRR-specific lookups fail (e.g. MRR not in PR)."""
    if not material:
        return 0.0

    material_col = _find_col(pr_df, 'material')
    mic_col = _find_col(pr_df, 'mic')
    category_col = _find_col(pr_df, 'categery', 'category')
    rate_col = [c for c in pr_df.columns if str(c).strip().lower() == 'rate']
    rate_col = rate_col[0] if rate_col else 'Rate'

    if not material_col:
        return 0.0

    mat_upper = str(material).strip().upper()
    aliases = PR_MAT_ALIASES.get(mat_upper, [mat_upper])

    # Filter by material
    def mat_match(x):
        xu = str(x).strip().upper()
        for alias in aliases:
            if xu == alias or xu.startswith(alias + ' '):
                return True
        return xu == mat_upper

    mask = pr_df[material_col].apply(mat_match)

    # Filter by category = Film (not adhesive/solvent)
    if category_col:
        cat_mask = pr_df[category_col].astype(str).str.lower().str.contains('film', na=False)
        combined = mask & cat_mask
        if combined.any():
            mask = combined

    filtered = pr_df[mask]
    if filtered.empty:
        return 0.0

    # Filter by month
    filtered = _filter_by_month(pr_df, filtered, report_month)

    # Filter by mic if provided — STRICT: no fallback to other mics
    # Cross-mic averages produce meaningless rates (e.g. WPE mic=100 using mic=40/85/90)
    if mic is not None and not pd.isna(mic) and mic_col:
        try:
            mic_val = float(mic)
            mic_f = filtered[pd.to_numeric(filtered[mic_col], errors='coerce') == mic_val]
            if not mic_f.empty:
                filtered = mic_f
            else:
                logger.warning(f"lookup_material_rate_for_month: No PR entries for "
                               f"{material} mic={mic_val} in {report_month}. "
                               f"Returning 0 (not mixing other mics).")
                return 0.0
        except (ValueError, TypeError):
            pass

    return _qty_weighted_rate(filtered, rate_col)


def filter_mrr_by_pr(pr_df: pd.DataFrame, mrr_qty_dict: dict,
                      material: str = None, size=None, mic=None) -> dict:
    """
    Filter mrr_qty_dict to only include MRRs that exist in Purchase Register
    with matching material/size/mic. Returns filtered {mrr: qty}.
    """
    tracking_col = _find_col(pr_df, 'tracking')
    material_col = _find_col(pr_df, 'material')
    size_col = _find_col(pr_df, 'size')
    mic_col = _find_col(pr_df, 'mic')
    
    if not tracking_col:
        return mrr_qty_dict
    
    valid = {}
    for mrr, qty in mrr_qty_dict.items():
        try:
            mrr_val = int(float(mrr))
        except (ValueError, TypeError):
            continue
        
        mask = pd.to_numeric(pr_df[tracking_col], errors='coerce') == mrr_val
        filtered = pr_df[mask]
        
        if filtered.empty:
            continue  # MRR not in PR at all — skip
        
        # Check if size matches (with tolerance)
        if size is not None and not pd.isna(size) and size_col:
            try:
                size_val = float(size)
                sizes_num = pd.to_numeric(filtered[size_col], errors='coerce')
                has_size = ((sizes_num >= size_val - 5) & (sizes_num <= size_val + 5)).any()
                if has_size:
                    valid[mrr] = qty
                    continue
            except (ValueError, TypeError):
                pass
        
        # No size filter or size not relevant — include if in PR
        valid[mrr] = qty
    
    return valid if valid else mrr_qty_dict  # Fallback to original if all filtered out


def _filter_by_month(pr_df: pd.DataFrame, filtered: pd.DataFrame, report_month: str = None):
    """Filter PR entries by reporting month (e.g. '2-2026' for Feb 2026).
    Falls back to MOST RECENT month <= report_month (not all entries)."""
    if not report_month:
        return filtered
    month_col = None
    for c in pr_df.columns:
        if str(c).strip().lower() == 'month':
            month_col = c
            break
    if month_col is None:
        return filtered
    month_filtered = filtered[filtered[month_col].astype(str).str.strip() == report_month]
    if not month_filtered.empty:
        return month_filtered

    # Fallback: find the most recent month <= report_month
    try:
        parts = report_month.split('-')
        report_key = int(parts[1]) * 100 + int(parts[0])  # e.g. 202611 for 11-2025
    except (ValueError, IndexError):
        return filtered  # Can't parse, fall back to all

    best_month = None
    best_key = 0
    for m in filtered[month_col].astype(str).str.strip().unique():
        mp = m.split('-')
        if len(mp) == 2:
            try:
                mk = int(mp[1]) * 100 + int(mp[0])
                if mk <= report_key and mk > best_key:
                    best_key = mk
                    best_month = m
            except (ValueError, IndexError):
                pass

    if best_month:
        logger.info(f"No data for month {report_month}, falling back to {best_month}")
        return filtered[filtered[month_col].astype(str).str.strip() == best_month]

    return filtered


def _qty_weighted_rate(filtered: pd.DataFrame, rate_col: str) -> float:
    """Calculate qty-weighted average rate using Amount / Actual Quantity.
    This is the general rule: Total Amount / Total Qty for the period."""
    # Try Amount / Actual Quantity first
    amt_col = None
    qty_col = None
    for c in filtered.columns:
        cl = str(c).strip().lower()
        if cl == 'amount':
            amt_col = c
        if cl == 'actual quantity':
            qty_col = c
    
    if amt_col and qty_col:
        amounts = pd.to_numeric(filtered[amt_col], errors='coerce').fillna(0)
        qtys = pd.to_numeric(filtered[qty_col], errors='coerce').fillna(0)
        total_amt = amounts.sum()
        total_qty = qtys.sum()
        if total_qty > 0:
            return total_amt / total_qty
    
    # Fallback: simple last rate
    rates = pd.to_numeric(filtered[rate_col], errors='coerce').dropna()
    return rates.iloc[-1] if not rates.empty else 0.0


def lookup_adhesive_rate(pr_df: pd.DataFrame, adh_name: str, report_month: str = None) -> float:
    """Look up adhesive rate from Purchase Register.
    Uses EXACT material matching and month-filtered qty-weighted average."""
    if not adh_name or pd.isna(adh_name):
        return 0.0

    category_col = _find_col(pr_df, 'categery', 'category')
    material_col = _find_col(pr_df, 'material')
    rate_col = [c for c in pr_df.columns if str(c).strip().lower() == 'rate']
    rate_col = rate_col[0] if rate_col else 'Rate'

    adh_upper = str(adh_name).strip().upper()
    name_map = {'75-300': 'MF 75-300', '85-300': 'MF 75-300'}
    lookup_name = name_map.get(adh_upper, adh_upper)

    mask = pd.Series([False] * len(pr_df))
    if category_col:
        mask = pr_df[category_col].astype(str).str.lower().str.contains('adhesive', na=False)

    if material_col:
        # EXACT match only — no fuzzy substring matching
        mat_mask = pr_df[material_col].astype(str).str.upper().str.strip() == lookup_name
        mask = mask & mat_mask

    filtered = pr_df[mask]
    if filtered.empty:
        return 0.0

    filtered = _filter_by_month(pr_df, filtered, report_month)
    return _qty_weighted_rate(filtered, rate_col)


def lookup_hardener_rate(pr_df: pd.DataFrame, adh_name: str, report_month: str = None) -> float:
    """Look up paired hardener rate based on adhesive name (LEGACY).
    Kept for backward compatibility. Prefer lookup_hardener_rate_by_name()."""
    if not adh_name or pd.isna(adh_name):
        return 0.0

    adh_upper = str(adh_name).strip().upper()
    name_map = {'75-300': 'MF 75-300', '85-300': 'MF 75-300'}
    adh_lookup = name_map.get(adh_upper, adh_upper)

    hardener_name = ADH_HARDENER_PAIRS.get(adh_lookup)
    if not hardener_name:
        for key, val in ADH_HARDENER_PAIRS.items():
            if key in adh_lookup or adh_lookup in key:
                hardener_name = val
                break

    if not hardener_name:
        return 0.0

    return lookup_hardener_rate_by_name(pr_df, hardener_name, report_month)


def lookup_hardener_rate_by_name(pr_df: pd.DataFrame, hardener_name: str, report_month: str = None) -> float:
    """Look up hardener rate directly by hardener material name (from DA column).
    Uses EXACT matching with name alias support and month-filtered qty-weighted average."""
    if not hardener_name or pd.isna(hardener_name):
        return 0.0

    category_col = _find_col(pr_df, 'categery', 'category')
    material_col = _find_col(pr_df, 'material')
    rate_col = [c for c in pr_df.columns if str(c).strip().lower() == 'rate']
    rate_col = rate_col[0] if rate_col else 'Rate'

    h_upper = str(hardener_name).strip().upper()
    # Apply name aliases (Jobtrack DA name -> PR material name)
    lookup_name = HARDENER_NAME_MAP.get(h_upper, h_upper)

    mask = pd.Series([False] * len(pr_df))
    if category_col:
        # Hardener materials are under 'adhesive' category in PR
        mask = pr_df[category_col].astype(str).str.lower().apply(
            lambda x: 'adhesive' in x or 'hardener' in x
        )

    if material_col:
        # EXACT match only — no fuzzy substring matching
        mat_mask = pr_df[material_col].astype(str).str.upper().str.strip() == lookup_name
        mask = mask & mat_mask

    filtered = pr_df[mask]
    if filtered.empty:
        return 0.0

    filtered = _filter_by_month(pr_df, filtered, report_month)
    return _qty_weighted_rate(filtered, rate_col)


def lookup_solvent_rate(pr_df: pd.DataFrame, solvent_col_header: str = "E/A", report_month: str = None) -> float:
    """Look up solvent (Ethyl Acetate) rate from Purchase Register.
    Uses EXACT matching and month-filtered qty-weighted average."""
    category_col = _find_col(pr_df, 'categery', 'category')
    material_col = _find_col(pr_df, 'material')
    rate_col = [c for c in pr_df.columns if str(c).strip().lower() == 'rate']
    rate_col = rate_col[0] if rate_col else 'Rate'

    solvent_name = 'ETHYL ACETATE'

    mask = pd.Series([False] * len(pr_df))
    if category_col:
        mask = pr_df[category_col].astype(str).str.lower().str.contains('solvent', na=False)

    if material_col:
        # EXACT match only
        mat_mask = pr_df[material_col].astype(str).str.upper().str.strip() == solvent_name
        mask = mask & mat_mask

    filtered = pr_df[mask]
    if filtered.empty:
        return 0.0

    filtered = _filter_by_month(pr_df, filtered, report_month)
    return _qty_weighted_rate(filtered, rate_col)

