"""
MRR Lookup Engine - v3
Finds MR# (Material Receipt Register number) from Stores Recordings.

KEY: Returns only DOMINANT MRRs (those contributing >= 10% of total issued qty).
MRRs are sorted by issue quantity descending.
"""
import pandas as pd
import logging

logger = logging.getLogger(__name__)


def load_stores_recordings(file) -> pd.DataFrame:
    """Load and parse Stores Recordings Excel file."""
    df = pd.read_excel(file, sheet_name=0, header=1)
    col_map = {c: str(c).strip().replace('\n', ' ').replace('\r', '') for c in df.columns}
    df.rename(columns=col_map, inplace=True)
    logger.info(f"Loaded Stores Recordings: {len(df)} rows")
    return df


# Material name aliases: Jobtrack name -> possible Stores Sub Category values
MATERIAL_ALIASES = {
    'WPE': ['PE WHITE', 'WPE', 'WLDPE', 'WLDPE NATURAL'],
    'WLDPE': ['WLDPE', 'PE WHITE', 'WPE', 'WLDPE NATURAL'],
    'PTD WPE': ['PE WHITE', 'WPE', 'WLDPE'],
    'PET': ['PET CHEM', 'PET UPF', 'PET CORONA', 'PET ANTI STATIC', 'PET MATTE',
            'PET UPF CORONA TRT', 'PET ALOX', 'PET SIOX', 'PET'],
    'MET PET': ['MET PET UPF', 'MET PET HIGH OD', 'MET PET TWIST', 'MET PET'],
    'MATTE TOPP': ['MATTE OPP', 'MATTE TOPP'],
    'MOPP': ['MOPP'],
    'MPOPP': ['MPOPP'],
    'TOPP': ['TOPP', 'TOPP LTS LABLE GRADE', 'TOPP BOTH SIDE ACRYLIC COATED'],
    'FOIL': ['FOIL', 'FOIL RETORT'],
    'TPE': ['TPE', 'TPE LOW SIT', 'TPE EASY TEARABLE', 'TPE EASY PEEL',
            'TPE PEELABLE', 'TPE HIGH CLARITY', 'TPE HIGH DART', 'TPE HD',
            'TPE SURLYN', 'TPE MET EVOH(2~3 micron)'],
    'BOPP': ['BOPP'],
    'CPP': ['WCPP', 'CPP'],
    'LDPE': ['LDPE', 'LDPE NATURAL'],
    'L1': ['L1'],
    'L2': ['L2'],
}


def _find_stores_columns(df):
    """Find key column names in Stores DataFrame."""
    cols = {}
    for c in df.columns:
        cl = str(c).lower().strip()
        if 'sub' in cl and 'cat' in cl:
            cols['sub_cat'] = c
        elif cl == 'mic':
            cols['mic'] = c
        elif cl == 'width':
            cols['width'] = c
        elif 'm.r.r' in cl and 'no' in cl:
            cols['mrr'] = c
        elif 'issue wo' in cl:
            cols['wo'] = c
        elif 'issue' in cl and 'process' in cl:
            cols['process'] = c
        elif 'issue' in cl and 'qty' in cl:
            cols['issue_qty'] = c
    return cols


def _material_matches(stores_sub_cat: str, jt_material: str) -> bool:
    """Check if a Stores sub-category matches a Jobtrack material name."""
    if not stores_sub_cat or not jt_material:
        return False
    sc = str(stores_sub_cat).strip().upper()
    jm = str(jt_material).strip().upper()
    if not sc or not jm:
        return False
    if sc == jm:
        return True
    aliases = MATERIAL_ALIASES.get(jm, [])
    for alias in aliases:
        if sc == alias or sc.startswith(alias + ' '):
            return True
    if sc.startswith(jm + ' ') or sc == jm:
        return True
    return False


def lookup_mrr(stores_df: pd.DataFrame, material: str, mic, width,
               order_no: str, process_filter: str = None) -> list:
    """
    Look up MR# from Stores Recordings.

    Returns only DOMINANT MRRs (those contributing >=10% of total issue qty),
    sorted by total issue quantity descending.
    """
    if not material or pd.isna(material):
        return []

    cols = _find_stores_columns(stores_df)
    if 'sub_cat' not in cols or 'mic' not in cols or 'mrr' not in cols:
        logger.warning("Missing required columns in Stores Recordings")
        return []

    sub_cat_col = cols['sub_cat']
    mic_col = cols['mic']
    width_col = cols.get('width')
    mrr_col = cols['mrr']
    wo_col = cols.get('wo')
    process_col = cols.get('process')
    qty_col = cols.get('issue_qty')

    # Step 1: Filter by material match
    mask = stores_df[sub_cat_col].apply(lambda x: _material_matches(str(x), material))

    # Step 2: Filter by MIC
    if mic is not None and not pd.isna(mic):
        try:
            mic_val = float(mic)
            mask = mask & (pd.to_numeric(stores_df[mic_col], errors='coerce') == mic_val)
        except (ValueError, TypeError):
            pass

    # Step 3: Filter by Order No
    if order_no and wo_col and not pd.isna(order_no):
        order_str = str(order_no).strip()
        mask = mask & (stores_df[wo_col].astype(str).str.strip() == order_str)

    # Step 4: Filter by Process
    if process_filter and process_col:
        mask = mask & stores_df[process_col].astype(str).str.upper().str.strip().apply(
            lambda x: process_filter.upper() in str(x) or str(x) == process_filter.upper()
        )

    filtered = stores_df[mask]

    if filtered.empty:
        return []

    # NOTE: Width is NOT filtered in Stores lookup.
    # The manual process includes ALL rolls of the same material/mic
    # for the same order+process, regardless of width.
    # Width matching is only applied in the PR rate lookup step.

    # Step 6: Only include entries with positive issue qty
    if qty_col:
        qty_numeric = pd.to_numeric(filtered[qty_col], errors='coerce').fillna(0)
        has_qty = filtered[qty_numeric > 0]
        if not has_qty.empty:
            filtered = has_qty

    # Step 7: Aggregate qty per MRR and select dominant ones
    if qty_col:
        filtered_c = filtered.copy()
        filtered_c['_mrr_int'] = pd.to_numeric(filtered_c[mrr_col], errors='coerce')
        filtered_c['_qty'] = pd.to_numeric(filtered_c[qty_col], errors='coerce').fillna(0)
        mrr_qty = filtered_c.groupby('_mrr_int')['_qty'].sum().sort_values(ascending=False)
        mrr_qty = mrr_qty[mrr_qty.index.notna()]

        if mrr_qty.empty:
            return []

        total_qty = mrr_qty.sum()
        if total_qty > 0:
            # Only include MRRs contributing >= 10% of total qty
            threshold = total_qty * 0.10
            dominant = mrr_qty[mrr_qty >= threshold]
            if dominant.empty:
                dominant = mrr_qty.head(1)  # At least return the top one
            return [int(m) for m in dominant.index]
        else:
            return [int(m) for m in mrr_qty.index[:1]]
    else:
        # Fallback: count entries per MRR
        from collections import Counter
        mrr_values = filtered[mrr_col].dropna()
        clean_mrrs = []
        for m in mrr_values:
            try:
                clean_mrrs.append(int(float(m)))
            except (ValueError, TypeError):
                continue
        if not clean_mrrs:
            return []
        counter = Counter(clean_mrrs)
        total = sum(counter.values())
        threshold = total * 0.10
        return [m for m, c in counter.most_common() if c >= threshold] or [counter.most_common(1)[0][0]]


def lookup_mrr_with_qty(stores_df: pd.DataFrame, material: str, mic, width,
                        order_no: str, process_filter: str = None) -> dict:
    """
    Like lookup_mrr but returns {mrr: total_qty} for weighted rate calculation.
    """
    if not material or pd.isna(material):
        return {}

    cols = _find_stores_columns(stores_df)
    if 'sub_cat' not in cols or 'mic' not in cols or 'mrr' not in cols:
        return {}

    sub_cat_col = cols['sub_cat']
    mic_col = cols['mic']
    width_col = cols.get('width')
    mrr_col = cols['mrr']
    wo_col = cols.get('wo')
    process_col = cols.get('process')
    qty_col = cols.get('issue_qty')

    mask = stores_df[sub_cat_col].apply(lambda x: _material_matches(str(x), material))

    if mic is not None and not pd.isna(mic):
        try:
            mic_val = float(mic)
            mask = mask & (pd.to_numeric(stores_df[mic_col], errors='coerce') == mic_val)
        except (ValueError, TypeError):
            pass

    if order_no and wo_col and not pd.isna(order_no):
        mask = mask & (stores_df[wo_col].astype(str).str.strip() == str(order_no).strip())

    if process_filter and process_col:
        mask = mask & stores_df[process_col].astype(str).str.upper().str.strip().apply(
            lambda x: process_filter.upper() in str(x) or str(x) == process_filter.upper()
        )

    filtered = stores_df[mask]
    if filtered.empty:
        return {}

    # NOTE: Width is NOT filtered in Stores lookup.
    # All rolls of same material/mic for the order+process are included.
    # Width matching happens only in the PR rate lookup step.

    # Only positive qty
    if qty_col:
        qty_numeric = pd.to_numeric(filtered[qty_col], errors='coerce').fillna(0)
        has_qty = filtered[qty_numeric > 0]
        if not has_qty.empty:
            filtered = has_qty

    # Aggregate
    if qty_col:
        filtered_c = filtered.copy()
        filtered_c['_mrr_int'] = pd.to_numeric(filtered_c[mrr_col], errors='coerce')
        filtered_c['_qty'] = pd.to_numeric(filtered_c[qty_col], errors='coerce').fillna(0)
        mrr_qty = filtered_c.groupby('_mrr_int')['_qty'].sum().sort_values(ascending=False)
        mrr_qty = mrr_qty[mrr_qty.index.notna()]
        return {int(m): q for m, q in mrr_qty.items()}
    else:
        from collections import Counter
        mrr_values = filtered[mrr_col].dropna()
        clean = [int(float(m)) for m in mrr_values if not pd.isna(m)]
        return dict(Counter(clean))


def match_formula_qtys_to_store(stores_df: pd.DataFrame, formula_qtys: list,
                                 material: str, mic, order_no: str,
                                 process_filter: str = None,
                                 tolerance: float = 1.0) -> dict:
    """
    Match individual formula quantity components to specific Store issue entries.

    Given a list of quantities extracted from a formula (e.g. [40.0, 486.2, 486.8]),
    find the matching Store issue entries and return {mrr: total_matched_qty}.

    This enables precise MRR identification per Jobtrack row, rather than
    returning ALL MRRs for the same work order.

    Args:
        stores_df: Stores Recordings DataFrame
        formula_qtys: List of float quantities from parsed formula
        material: Material name (e.g. 'PET')
        mic: Micron value
        order_no: Work order number
        process_filter: Issue process filter (e.g. 'PRINTING', 'LAMINATION')
        tolerance: Maximum absolute difference for qty matching (default 0.5)

    Returns:
        dict of {mrr_number: total_matched_qty}, or empty dict if matching fails
    """
    if not formula_qtys or not material or pd.isna(material):
        return {}

    cols = _find_stores_columns(stores_df)
    if 'sub_cat' not in cols or 'mic' not in cols or 'mrr' not in cols:
        return {}

    sub_cat_col = cols['sub_cat']
    mic_col = cols['mic']
    mrr_col = cols['mrr']
    wo_col = cols.get('wo')
    process_col = cols.get('process')
    qty_col = cols.get('issue_qty')

    if not qty_col:
        return {}

    # Build filter mask (same as lookup_mrr_with_qty)
    mask = stores_df[sub_cat_col].apply(lambda x: _material_matches(str(x), material))

    if mic is not None and not pd.isna(mic):
        try:
            mic_val = float(mic)
            mask = mask & (pd.to_numeric(stores_df[mic_col], errors='coerce') == mic_val)
        except (ValueError, TypeError):
            pass

    if order_no and wo_col and not pd.isna(order_no):
        mask = mask & (stores_df[wo_col].astype(str).str.strip() == str(order_no).strip())

    if process_filter and process_col:
        mask = mask & stores_df[process_col].astype(str).str.upper().str.strip().apply(
            lambda x: process_filter.upper() in str(x) or str(x) == process_filter.upper()
        )

    filtered = stores_df[mask].copy()
    if filtered.empty:
        return {}

    # Only positive qty entries
    filtered['_qty'] = pd.to_numeric(filtered[qty_col], errors='coerce').fillna(0)
    filtered = filtered[filtered['_qty'] > 0]
    if filtered.empty:
        return {}

    filtered['_mrr_int'] = pd.to_numeric(filtered[mrr_col], errors='coerce')

    # Greedy matching: for each formula qty, find the best matching Store entry
    # Use a copy of available entries (each store entry can only be matched once)
    available = filtered[['_qty', '_mrr_int']].reset_index(drop=True).copy()
    available['_used'] = False

    matched_mrr_qty = {}  # {mrr: total_qty}
    matched_count = 0

    for fq in formula_qtys:
        if available[~available['_used']].empty:
            break

        # Find closest unused entry
        unused = available[~available['_used']]
        diffs = (unused['_qty'] - fq).abs()
        best_idx = diffs.idxmin()
        best_diff = diffs.loc[best_idx]

        if best_diff <= tolerance:
            mrr = available.loc[best_idx, '_mrr_int']
            qty = available.loc[best_idx, '_qty']
            available.loc[best_idx, '_used'] = True
            matched_count += 1

            if pd.notna(mrr):
                mrr_int = int(mrr)
                matched_mrr_qty[mrr_int] = matched_mrr_qty.get(mrr_int, 0) + qty
        # If no match within tolerance, skip (it might be a balance component)

    if matched_count == 0:
        return {}

    # Only return results if we matched a reasonable fraction of formula components
    # (some components may be balance quantities that won't match Store entries)
    match_ratio = matched_count / len(formula_qtys)
    if match_ratio < 0.3:
        # Too few matches — formula components don't align with Store entries
        logger.warning(f"Formula matching: only {matched_count}/{len(formula_qtys)} "
                       f"components matched Store entries for {material}/{order_no}. "
                       f"Falling back to full lookup.")
        return {}

    logger.info(f"Formula matching: {matched_count}/{len(formula_qtys)} components "
                f"matched -> MRRs: {matched_mrr_qty}")
    return matched_mrr_qty
