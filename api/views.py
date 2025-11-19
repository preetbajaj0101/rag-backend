from rest_framework.decorators import api_view
from rest_framework.response import Response
from rest_framework import status
import pandas as pd
import os, re
from difflib import SequenceMatcher
from threading import Lock

# --------------------------------------------
# LOAD DATASET
# --------------------------------------------

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SAMPLE_PATH = os.path.join(BASE_DIR, "data", "sample_realestate.xlsx")

_SAMPLE_DF = None
_SAMPLE_LOCK = Lock()


def get_sample_df():
    global _SAMPLE_DF
    if _SAMPLE_DF is not None:
        return _SAMPLE_DF

    with _SAMPLE_LOCK:
        if _SAMPLE_DF is not None:
            return _SAMPLE_DF

        try:
            df = pd.read_excel(SAMPLE_PATH)
            df.columns = [str(c).strip() for c in df.columns]
            _SAMPLE_DF = df
            print("[api] Loaded sample dataset columns:", _SAMPLE_DF.columns.tolist())
        except Exception as e:
            print("[api] Error loading sample file:", e)
            _SAMPLE_DF = pd.DataFrame()

        return _SAMPLE_DF


# --------------------------------------------
# CLEANING & MATCHING HELPERS
# --------------------------------------------

def clean_text(s):
    """Lowercase + punctuation removal + space normalize."""
    if pd.isna(s):
        return ""
    s = str(s).lower()
    s = s.replace("₹", " ").replace("rs.", " ").replace("inr", " ")
    s = re.sub(r"[()\[\]/,.-]+", " ", s)
    s = re.sub(r"[^a-z0-9\s]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def smart_match_string(candidate, target):
    c = clean_text(candidate)
    t = clean_text(target)

    if not c or not t:
        return False, 0.0, "empty"

    # Strong substring
    if t in c or c in t:
        return True, 1.0, "substring"

    # Token match
    c_tokens = set(c.split())
    t_tokens = set(t.split())
    if c_tokens.intersection(t_tokens):
        return True, 0.85, "token"

    # Fuzzy
    ratio = SequenceMatcher(None, c, t).ratio()
    if ratio >= 0.55:
        return True, ratio, "fuzzy"

    return False, ratio, "no_match"


# --------------------------------------------
# COLUMN DETECTION
# --------------------------------------------

def find_area_column(df):
    prefs = ["final location", "location", "area", "locality"]
    lower = [c.lower() for c in df.columns]
    for p in prefs:
        if p in lower:
            return df.columns[lower.index(p)]
    # fallback
    for c in df.columns:
        if df[c].dtype == object:
            return c
    return None


def find_price_column(df):
    prefs = [
        "flat - weighted average rate",
        "shop - weighted average rate",
        "office - weighted average rate",
        "price",
    ]
    lower = [c.lower() for c in df.columns]
    for p in prefs:
        if p in lower:
            return df.columns[lower.index(p)]

    # fallback numeric
    for c in df.columns:
        if pd.api.types.is_numeric_dtype(df[c]):
            return c
    return None


def find_demand_column(df):
    prefs = [
        "total_sales - igr",
        "total sold - igr",
        "total units",
    ]
    lower = [c.lower() for c in df.columns]
    for p in prefs:
        if p in lower:
            return df.columns[lower.index(p)]

    # fallback numeric
    for c in df.columns:
        if pd.api.types.is_numeric_dtype(df[c]):
            return c
    return None


def to_number_safe(x):
    if pd.isna(x):
        return None
    if isinstance(x, (int, float)):
        return float(x)
    s = str(x).replace(",", "")
    m = re.search(r"[-+]?\d+(\.\d+)?", s)
    return float(m.group(0)) if m else None


# --------------------------------------------
# AREA FILTER
# --------------------------------------------

def area_filter(df, area_name, top_n=8):
    col = find_area_column(df)
    if col is None:
        return pd.DataFrame(), None, []

    work = df.copy()
    work["_clean_loc"] = work[col].astype(str).apply(clean_text)

    results = []
    for idx, row in work.iterrows():
        matched, score, reason = smart_match_string(row["_clean_loc"], area_name)
        results.append((idx, matched, score, reason, row["_clean_loc"]))

    score_df = pd.DataFrame(results, columns=["idx", "matched", "score", "reason", "clean"])

    # strict match
    strict = score_df[score_df["matched"] == True].sort_values("score", ascending=False)
    if not strict.empty:
        selected = strict["idx"].tolist()
        return work.loc[selected], col, strict.head(top_n).to_dict(orient="records")

    # fuzzy fallback
    fuzzy = score_df[score_df["score"] >= 0.45].sort_values("score", ascending=False)
    if not fuzzy.empty:
        selected = fuzzy["idx"].tolist()
        return work.loc[selected], col, fuzzy.head(top_n).to_dict(orient="records")

    return pd.DataFrame(), col, []


# --------------------------------------------
# ENDPOINT: QUERY AREA
# --------------------------------------------

@api_view(["POST"])
def query_area(request):
    q = request.data.get("query", "")
    if not isinstance(q, str) or not q.strip():
        return Response({"error": "Empty query"}, status=400)

    SAMPLE_DF = get_sample_df()
    if SAMPLE_DF.empty:
        return Response({"summary": "No data"}, status=200)

    raw_query = q.strip()
    ql = raw_query.lower()

    # ---- detect price-growth mode and parse N (default to 3 if specified phrase present without number) ----
    mode_price_growth = "price growth" in ql or "price trend" in ql or "price increase" in ql
    n_years = None
    m = re.search(r"(?:last|past)\s+(\d+)\s+years?", ql)
    if m:
        n_years = int(m.group(1))
    elif mode_price_growth and ("last 3 years" in ql or "past 3 years" in ql or "last three years" in ql):
        n_years = 3
    # if price growth requested but no number, we will default to last 3 years
    if mode_price_growth and n_years is None:
        n_years = 3

    # ---- extract candidate location phrase (keep multi-word) ----
    # remove common command words but leave location
    tokens_to_remove = [
        "show", "please", "give", "me", "analyze", "analysis", "price growth",
        "price trend", "price increase", "over the last", "over the", "for", "in"
    ]
    tmp = ql
    for t in tokens_to_remove:
        tmp = tmp.replace(t, " ")
    tmp = re.sub(r"\s+", " ", tmp).strip()

    candidate = tmp.strip()

    # ---- build unique cleaned location list from dataset ----
    area_col = find_area_column(SAMPLE_DF)
    unique_locations = []
    if area_col and area_col in SAMPLE_DF.columns:
        unique_locations = SAMPLE_DF[area_col].astype(str).dropna().map(lambda x: x.strip()).drop_duplicates().tolist()

    chosen_target = None
    # try direct candidate match first
    if candidate:
        f_try, _, m_try = area_filter(SAMPLE_DF, candidate)
        if not f_try.empty:
            chosen_target = candidate

    # if not found, try dataset-driven best match against the whole query (to catch "show price growth for Akurdi")
    if not chosen_target and unique_locations:
        best_score = 0.0
        best_loc = None
        for loc in unique_locations:
            matched, score, reason = smart_match_string(clean_text(loc), ql)
            if matched and score > best_score:
                best_score = score
                best_loc = loc
        if best_loc:
            chosen_target = best_loc

    # fallback to raw candidate
    if not chosen_target and candidate:
        chosen_target = candidate

    if not chosen_target:
        return Response({"summary": f"Could not detect location in query '{raw_query}'"}, status=200)

    # ---- get filtered rows for chosen_target ----
    filtered, used_col, match_info = area_filter(SAMPLE_DF, chosen_target)
    # if area_filter failed but chosen_target exactly matches a unique location, try exact substring filter
    if filtered.empty and chosen_target in unique_locations and area_col:
        filtered = SAMPLE_DF[SAMPLE_DF[area_col].astype(str).str.lower().str.contains(chosen_target.lower(), na=False)]

    if filtered.empty:
        return Response({"summary": f"No results for '{chosen_target}'", "match_info": match_info}, status=200)

    price_col = find_price_column(SAMPLE_DF)
    if price_col and price_col in filtered.columns:
        filtered["_price_num"] = filtered[price_col].apply(to_number_safe)
    else:
        return Response({"summary": f"No price column detected for dataset; cannot compute price growth", "match_info": match_info}, status=200)

    # ---- aggregate by year: compute mean price per distinct year ----
    if "year" not in filtered.columns:
        return Response({"summary": f"No 'year' column in dataset for {chosen_target}", "match_info": match_info}, status=200)

    # drop rows that don't have a valid year
    df_year = filtered.dropna(subset=["year"]).copy()
    # ensure year numeric/int
    try:
        df_year["year"] = df_year["year"].astype(int)
    except Exception:
        # try safe conversion
        df_year["year"] = pd.to_numeric(df_year["year"], errors="coerce").dropna().astype(int)

    agg = df_year.groupby("year")["_price_num"].mean().reset_index().sort_values("year")
    # list of (year, price)
    year_price = [(int(r["year"]), float(r["_price_num"]) if not pd.isna(r["_price_num"]) else None) for _, r in agg.iterrows()]

    if not year_price:
        return Response({"summary": f"No price-by-year data available for '{chosen_target}'", "match_info": match_info}, status=200)

    # ---- pick most recent N distinct years ----
    if n_years is not None:
        last_years = year_price[-n_years:] if n_years <= len(year_price) else year_price
    else:
        last_years = year_price

    # Build chart (year, price) for those years
    chart = [{"year": y, "price": p} for (y, p) in last_years]

    # ---- compute per-year percent change and overall percent ----
    per_year_pct = []
    overall_pct = None
    prices = [p for _, p in last_years]
    years = [y for y, _ in last_years]

    prev = None
    for idx, (y, p) in enumerate(last_years):
        if prev is None or prev is None or p is None:
            per_year_pct.append({"year": y, "pct_change_vs_prev": None})
        else:
            try:
                per_year_pct.append({"year": y, "pct_change_vs_prev": (p - prev) / prev * 100.0})
            except Exception:
                per_year_pct.append({"year": y, "pct_change_vs_prev": None})
        prev = p if p is not None else prev

    # overall: use first non-null -> last non-null
    non_null_prices = [p for p in prices if p is not None]
    if len(non_null_prices) >= 2 and non_null_prices[0] not in (0, None):
        try:
            overall_pct = (non_null_prices[-1] - non_null_prices[0]) / non_null_prices[0] * 100.0
        except Exception:
            overall_pct = None

    # avg price across selected years
    avg_price = None
    valid_prices = [p for p in prices if p is not None]
    if valid_prices:
        avg_price = sum(valid_prices) / len(valid_prices)

    summary = f"Price growth for {chosen_target}: {years[0]}→{years[-1]} · Avg price = {avg_price:.0f}" if years else f"Price growth for {chosen_target}"

    return Response({
        "summary": summary,
        "chart": chart,
        "per_year_pct": per_year_pct,
        "overall_pct": overall_pct,
        "years": years,
        "match_info": match_info
    }, status=200)

@api_view(['POST'])
def upload_excel(request):
    global _SAMPLE_DF

    file = request.FILES.get('file')
    if not file:
        return Response({'error': 'No file uploaded'}, status=400)

    try:
        df = pd.read_excel(file)
        df.columns = [str(c).strip() for c in df.columns]
        _SAMPLE_DF = df

        return Response({
            'message': 'File uploaded successfully',
            'columns': df.columns.tolist(),
            'rows': len(df)
        })

    except Exception as e:
        return Response({'error': str(e)}, status=400)
    
@api_view(["POST"])
def compare_areas(request):
    """
    Robust comparison endpoint.
    Accepts either:
      - JSON { "areas": ["Akrudi", "Wakad"] }
      - JSON { "area1": "Akrudi", "area2": "Wakad" }
      - JSON { "query": "compare akrudi wakad" }  (supports commas, 'vs', 'and')
    Returns per-area stats, per-area chart (year -> price/demand), and combined_chart
    """
    data = request.data or {}

    # 1) prefer explicit 'areas' list
    areas = data.get("areas")
    if isinstance(areas, list) and len(areas) >= 2:
        areas_list = [str(a).strip() for a in areas if str(a).strip()]
    else:
        # 2) fallback to area1 / area2 keys
        a1 = data.get("area1")
        a2 = data.get("area2")
        if a1 and a2:
            areas_list = [str(a1).strip(), str(a2).strip()]
        else:
            # 3) fallback to parsing a natural language query
            q = data.get("query", "")
            if not q or not isinstance(q, str):
                return Response({"error": "Provide 'areas' list, or 'area1' and 'area2', or 'query'."}, status=400)
            q_lower = q.lower()
            # remove the leading 'compare' if present
            if q_lower.startswith("compare"):
                q_body = q[len("compare"):].strip()
            else:
                q_body = q

            # split by common separators while preserving multi-word tokens
            # prefer commas, then ' vs ', then ' and ', otherwise split on multiple spaces
            if "," in q_body:
                parts = [p.strip() for p in q_body.split(",") if p.strip()]
            elif " vs " in q_body:
                parts = [p.strip() for p in q_body.split(" vs ") if p.strip()]
            elif " and " in q_body:
                parts = [p.strip() for p in q_body.split(" and ") if p.strip()]
            else:
                # last resort: split on whitespace and try to reconstruct multi-word names
                # we'll attempt to match against dataset known locations to reconstruct—fallback simple split
                # Try to use dataset values to find best segmentation:
                SAMPLE_DF = get_sample_df()
                loc_col = find_area_column(SAMPLE_DF) if not SAMPLE_DF.empty else None
                text_parts = q_body.split()
                parts = []
                if loc_col is not None and not SAMPLE_DF.empty:
                    # attempt greedy longest-match using cleaned names
                    cleaned_locs = (
                        SAMPLE_DF[loc_col].astype(str)
                        .dropna()
                        .map(clean_text)
                        .drop_duplicates()
                        .tolist()
                    )
                    # reconstruct by scanning tokens greedily
                    i = 0
                    while i < len(text_parts):
                        # try longest span
                        matched = None
                        for j in range(len(text_parts), i, -1):
                            candidate = " ".join(text_parts[i:j])
                            if clean_text(candidate) in cleaned_locs:
                                matched = " ".join(text_parts[i:j])
                                parts.append(matched)
                                i = j
                                break
                        if matched is None:
                            # if no match, take single token
                            parts.append(text_parts[i])
                            i += 1
                else:
                    parts = [p.strip() for p in q_body.split() if p.strip()]

            # final sanitize
            areas_list = [p for p in parts if p]

    # enforce at least 2 areas
    if not isinstance(areas_list, list) or len(areas_list) < 2:
        return Response({"error": "Please provide at least two areas to compare."}, status=400)

    SAMPLE_DF = get_sample_df()
    if SAMPLE_DF.empty:
        return Response({"error": "Dataset not loaded"}, status=500)

    price_col = find_price_column(SAMPLE_DF)
    demand_col = find_demand_column(SAMPLE_DF)
    area_col = find_area_column(SAMPLE_DF)

    results = []
    all_years = set()

    for area in areas_list:
        filtered, used_col, match_info = area_filter(SAMPLE_DF, area)
        if filtered.empty:
            results.append({
                "area": area,
                "found": False,
                "summary": f"No results for '{area}'",
                "chart": [],
                "table": [],
                "match_info": match_info
            })
            continue

        work = filtered.copy()
        # numeric coercion
        if price_col and price_col in work.columns:
            work["_price_num"] = work[price_col].apply(to_number_safe)
        if demand_col and demand_col in work.columns:
            work["_demand_num"] = work[demand_col].apply(to_number_safe)

        # build per-year chart
        chart = []
        if "year" in work.columns:
            agg = {}
            if "_price_num" in work.columns:
                agg["_price_num"] = "mean"
            if "_demand_num" in work.columns:
                agg["_demand_num"] = "mean"
            if agg:
                grp = work.groupby("year").agg(agg).reset_index().sort_values("year")
                for _, r in grp.iterrows():
                    yr = int(r["year"]) if not pd.isna(r["year"]) else None
                    entry = {
                        "year": yr,
                        "price": None if "_price_num" not in r or pd.isna(r["_price_num"]) else float(r["_price_num"])
                    }
                    if "_demand_num" in r.index:
                        entry["demand"] = None if pd.isna(r["_demand_num"]) else float(r["_demand_num"])
                    chart.append(entry)
                    if yr is not None:
                        all_years.add(yr)

        # small safe table (first few rows)
        safe_cols = [c for c in [ "year", used_col or area_col, price_col, demand_col, "city" ] if c in work.columns]
        small_table = work[safe_cols].head(8).to_dict(orient="records")

        # stats
        stats = {}
        if "_price_num" in work.columns and not work["_price_num"].dropna().empty:
            stats["avg_price"] = float(work["_price_num"].mean())
        if "_demand_num" in work.columns and not work["_demand_num"].dropna().empty:
            stats["avg_demand"] = float(work["_demand_num"].mean())

        results.append({
            "area": area,
            "found": True,
            "stats": stats,
            "chart": chart,
            "table": small_table,
            "match_info": match_info
        })

    # build combined chart using union of years
    all_years = sorted(list(all_years))
    combined_chart = []
    for yr in all_years:
        row = {"year": yr}
        for res in results:
            key_prefix = re.sub(r"\s+", "_", res["area"].lower())
            # find entry in res chart for year
            match = next((c for c in res["chart"] if c["year"] == yr), None)
            row[f"{key_prefix}_price"] = match["price"] if match else None
            row[f"{key_prefix}_demand"] = match.get("demand") if (match and "demand" in match) else None
        combined_chart.append(row)

    return Response({
        "areas": results,
        "combined_chart": combined_chart,
    }, status=200)
