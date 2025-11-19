# backend/api/views.py
from rest_framework.decorators import api_view
from rest_framework.response import Response
from rest_framework import status
import pandas as pd
import os
from threading import Lock
import re

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SAMPLE_PATH = os.path.join(BASE_DIR, 'data', 'sample_realestate.xlsx')

_SAMPLE_DF = None
_SAMPLE_LOCK = Lock()

def _to_number_safe(x):
    """Try to coerce x to float. Handles strings like '1,234,567', '₹1,23,456', or '10-12' (takes first number)."""
    if pd.isna(x):
        return None
    if isinstance(x, (int, float)):
        return float(x)
    s = str(x)
    # find first number (integer or decimal) in the string
    m = re.search(r'[-+]?\d{1,3}(?:[,\d]*\d)?(?:\.\d+)?|\d+\.\d+|\d+', s.replace('₹','').replace('INR',''))
    if not m:
        return None
    num_str = m.group(0)
    # remove commas
    num_str = num_str.replace(',', '')
    try:
        return float(num_str)
    except Exception:
        return None

def get_sample_df():
    global _SAMPLE_DF
    if _SAMPLE_DF is not None:
        return _SAMPLE_DF
    with _SAMPLE_LOCK:
        if _SAMPLE_DF is not None:
            return _SAMPLE_DF
        if not os.path.exists(SAMPLE_PATH):
            _SAMPLE_DF = pd.DataFrame()
            return _SAMPLE_DF
        try:
            df = pd.read_excel(SAMPLE_PATH)
            # Preserve original column names but normalise a lower-case list for matching
            df.columns = [str(c).strip() for c in df.columns]
            _SAMPLE_DF = df
            print(f"[api] Loaded sample with columns: {_SAMPLE_DF.columns.tolist()}")
            return _SAMPLE_DF
        except Exception as e:
            print(f"[api] Error loading sample Excel at {SAMPLE_PATH}: {e}")
            _SAMPLE_DF = pd.DataFrame()
            return _SAMPLE_DF

def find_area_column(df: pd.DataFrame):
    # prefer explicit 'final location' from your file, else common names, else first text column
    if df.empty:
        return None
    cols_lower = [c.lower() for c in df.columns]
    prefs = ['final location', 'area', 'locality', 'location', 'neighbourhood', 'neighborhood', 'city', 'town']
    for p in prefs:
        if p in cols_lower:
            return df.columns[cols_lower.index(p)]
    # fallback: first object/string dtype column
    for i, dt in enumerate(df.dtypes):
        if str(dt).startswith('object'):
            return df.columns[i]
    return None

def find_price_column(df: pd.DataFrame):
    # try typical numeric / rate columns from your sheet
    if df.empty:
        return None
    cols_lower = [c.lower() for c in df.columns]
    prefs = [
        'price', 
        'flat - weighted average rate',
        'office - weighted average rate',
        'others - weighted average rate',
        'shop - weighted average rate',
        'flat - most prevailing rate - range',
        'office - most prevailing rate - range',
    ]
    for p in prefs:
        if p in cols_lower:
            return df.columns[cols_lower.index(p)]
    # fallback: any numeric column
    for i, dt in enumerate(df.dtypes):
        if pd.api.types.is_numeric_dtype(dt):
            return df.columns[i]
    # fallback: none
    return None

def find_demand_column(df: pd.DataFrame):
    # try sales/units columns as "demand"
    if df.empty:
        return None
    cols_lower = [c.lower() for c in df.columns]
    prefs = ['total_sales - igr', 'total_units', 'total_sales', 'total_sold - igr']
    for p in prefs:
        if p in cols_lower:
            return df.columns[cols_lower.index(p)]
    # fallback: any numeric column not chosen as price
    for i, dt in enumerate(df.dtypes):
        if pd.api.types.is_numeric_dtype(dt):
            return df.columns[i]
    return None

def area_filter(df: pd.DataFrame, area_name: str) -> pd.DataFrame:
    col = find_area_column(df)
    if col is None:
        return pd.DataFrame()
    try:
        return df[df[col].astype(str).str.lower().str.contains(area_name.lower(), na=False)]
    except Exception:
        # last-resort: simple substring in stringified row
        mask = df.apply(lambda row: area_name.lower() in ' '.join(map(str,row.values)).lower(), axis=1)
        return df[mask]

def make_summary(df: pd.DataFrame, area: str, price_col=None, demand_col=None) -> str:
    parts = []
    if price_col and price_col in df.columns and not df[price_col].dropna().empty:
        # coerce and average
        vals = df[price_col].apply(_to_number_safe).dropna()
        if not vals.empty:
            parts.append(f"Avg price {price_col}: {vals.mean():.0f}")
    if demand_col and demand_col in df.columns and not df[demand_col].dropna().empty:
        vals = df[demand_col].apply(_to_number_safe).dropna()
        if not vals.empty:
            parts.append(f"Avg demand ({demand_col}): {vals.mean():.1f}")
    if 'year' in df.columns and not df['year'].dropna().empty:
        years = sorted(df['year'].dropna().unique().tolist())
        parts.append(f"Years: {years}")
    if not parts:
        return f"{area}: data found but no numeric columns to summarize."
    return f"{area}: " + " · ".join(parts)

@api_view(['POST'])
def upload_excel(request):
    file = request.FILES.get('file')
    if not file:
        return Response({'error': 'No file uploaded'}, status=status.HTTP_400_BAD_REQUEST)
    try:
        df = pd.read_excel(file)
    except Exception as e:
        return Response({'error': f'Failed to read Excel file: {str(e)}'}, status=status.HTTP_400_BAD_REQUEST)
    rows = len(df)
    cols = df.columns.tolist()
    return Response({'message': 'file processed', 'rows': rows, 'columns': cols})

@api_view(['POST'])
def query_area(request):
    q = request.data.get('query', '')
    if not isinstance(q, str) or not q.strip():
        return Response({'error': 'Empty or invalid query'}, status=status.HTTP_400_BAD_REQUEST)

    # conservative parse for the area token
    tokens = q.strip().split()
    stop_words = {'analyze','analyser','analyze:','analyze,','analyze?','analyzes','please','show','for','in','the','of','what','give'}
    candidate = None
    for t in reversed(tokens):
        t_clean = t.strip().strip('.,?').lower()
        if t_clean and t_clean not in stop_words:
            candidate = t_clean
            break
    area = candidate or q.strip()

    SAMPLE_DF = get_sample_df()
    if SAMPLE_DF.empty:
        return Response({'summary': 'No sample data available on server.'})

    # columns for debugging / frontend adapt
    available_columns = SAMPLE_DF.columns.tolist()

    # best columns
    price_col = find_price_column(SAMPLE_DF)
    demand_col = find_demand_column(SAMPLE_DF)
    area_col = find_area_column(SAMPLE_DF)

    df_filtered = area_filter(SAMPLE_DF, area)
    if df_filtered.empty:
        return Response({
            'summary': f'No data found for \"{area}\" using available columns ({area_col}).',
            'chart': [],
            'table': [],
            'available_columns': available_columns
        })

    # Build chart data grouped by year
    chart = []
    if 'year' in df_filtered.columns:
        group_keys = ['year']
        # prepare a DataFrame with numeric-converted price/demand if possible
        working = df_filtered.copy()
        if price_col:
            working['_price_num'] = working[price_col].apply(_to_number_safe)
        if demand_col:
            working['_demand_num'] = working[demand_col].apply(_to_number_safe)
        agg = {}
        if '_price_num' in working.columns:
            agg['_price_num'] = 'mean'
        if '_demand_num' in working.columns:
            agg['_demand_num'] = 'mean'
        if agg:
            grouped = working.groupby('year').agg(agg).reset_index()
            # map names back to readable keys
            records = []
            for _, row in grouped.iterrows():
                rec = {'year': int(row['year']) if not pd.isna(row['year']) else row['year']}
                if '_price_num' in row.index:
                    rec['price'] = None if pd.isna(row['_price_num']) else float(row['_price_num'])
                if '_demand_num' in row.index:
                    rec['demand'] = None if pd.isna(row['_demand_num']) else float(row['_demand_num'])
                records.append(rec)
            chart = records

    # Build table (serialize safe numeric coercion for price/demand)
    table = []
    for r in df_filtered.to_dict(orient='records'):
        out = dict(r)  # copy
        if price_col and price_col in out:
            out['_price_num'] = _to_number_safe(out.get(price_col))
        if demand_col and demand_col in out:
            out['_demand_num'] = _to_number_safe(out.get(demand_col))
        table.append(out)

    summary = make_summary(df_filtered, area, price_col=price_col, demand_col=demand_col)
    return Response({
        'summary': summary,
        'chart': chart,
        'table': table,
        'available_columns': available_columns,
        'inferred': {
            'area_column': area_col,
            'price_column': price_col,
            'demand_column': demand_col
        }
    })
