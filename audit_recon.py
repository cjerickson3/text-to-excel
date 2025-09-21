# --- Reconciliation/Audit Helpers -----------------------------------------
import pandas as pd
from collections import defaultdict

def _col(df, *cands, req=True):
    for c in cands:
        if c in df.columns: return c
    if req:
        raise KeyError(f"Missing required column; tried {cands}")
    return None

def normalize(df):
    df = df.copy()
    # Try to standardize expected columns (tweak names here if yours differ)
    c_date   = _col(df, 'Date','TxnDate','Post Date','Posted','TransactionDate')
    c_amt    = _col(df, 'Amount','Amt','Transaction Amount')
    c_desc   = _col(df, 'Description','Payee','Memo','Details')
    c_file   = _col(df, 'SourceFile','Statement','Source','SrcFile', req=False)
    c_month  = _col(df, 'StatementMonth','Month','StmtMonth', req=False)
    c_year   = _col(df, 'StatementYear','Year','StmtYear', req=False)
    c_cat    = _col(df, 'Category','Cat', req=False)
    c_sub    = _col(df, 'Subcategory','SubCat', req=False)

    out = pd.DataFrame({
        'date': pd.to_datetime(df[c_date]),
        'amount': pd.to_numeric(df[c_amt]),
        'desc': df[c_desc].astype(str).str.strip().str.replace(r'\s+', ' ', regex=True),
    })
    if c_file:  out['source_file'] = df[c_file]
    if c_month: out['stmt_month']  = pd.to_numeric(df[c_month], errors='coerce').astype('Int64')
    if c_year:  out['stmt_year']   = pd.to_numeric(df[c_year],  errors='coerce').astype('Int64')
    if c_cat:   out['category']    = df[c_cat]
    if c_sub:   out['subcategory'] = df[c_sub]
    out['abs_amount'] = out['amount'].abs()
    out['sign'] = out['amount'].apply(lambda x: 1 if x>0 else (-1 if x<0 else 0))
    out['y'] = out['date'].dt.year
    out['m'] = out['date'].dt.month
    if 'category' not in out.columns:
        out['category'] = ''
    if 'subcategory' not in out.columns:
        out['subcategory'] = ''
    return out

def imbalance_summary(df_norm):
    # 1) Total by calendar month vs statement-month (if available)
    pieces = []
    cal = df_norm.groupby(['y','m'], dropna=False)['amount'].sum().reset_index().rename(columns={'amount':'sum_calendar'})
    pieces.append(cal)

    if 'stmt_year' in df_norm and 'stmt_month' in df_norm:
        stmt = df_norm.groupby(['stmt_year','stmt_month'], dropna=False)['amount'].sum().reset_index()
        stmt = stmt.rename(columns={'stmt_year':'y','stmt_month':'m','amount':'sum_stmt'})
        merged = pd.merge(cal, stmt, on=['y','m'], how='outer').fillna(0)
    else:
        merged = cal
        merged['sum_stmt'] = 0.0

    merged['delta_stmt_vs_calendar'] = merged['sum_stmt'] - merged['sum_calendar']

    # 2) Sign anomalies: deposits with negative words & withdrawals with positive words
    dep_words = ('deposit','payroll','transfer in','refund','return','credit')
    wd_words  = ('payment','debit','ach','check','withdrawal','fee','card','purchase','pos')
    sign_flags = df_norm.assign(
        has_dep_word = df_norm['desc'].str.lower().str.contains('|'.join(dep_words)),
        has_wd_word  = df_norm['desc'].str.lower().str.contains('|'.join(wd_words)),
    )
    # NEW: pick only available columns (subcategory is optional)
    cols = ['date', 'amount', 'desc'] + [c for c in ('category','subcategory') if c in df_norm.columns]
    sign_anom = sign_flags.query(
        '(has_dep_word and amount<0) or (has_wd_word and amount>0)'
    )[cols]
    
    # 3) Near-duplicates (same date, |amount|, and fuzzy desc)
    #    This does not delete anything; it just shows clusters that look duplicated.
    key = (df_norm['date'].dt.date.astype(str) + '|' +
           df_norm['abs_amount'].round(2).astype(str) + '|' +
           df_norm['desc'].str.lower().str.replace(r'[^a-z0-9 ]','', regex=True).str.replace(r'\s+',' ', regex=True).str.slice(0,32))
    dup = df_norm.copy()
    dup['dup_key'] = key
    dup_groups = dup.groupby('dup_key')
    dup_candidates = dup_groups.filter(lambda g: len(g)>1).sort_values(['date','abs_amount'])

    # 4) Pershing mid-month triplet check (your 3,000 + 500 + 500 expectation)
    pershing = df_norm[df_norm['desc'].str.contains('pershing', case=False, na=False)].copy()
    pershing['ym'] = pershing['date'].dt.to_period('M')
    pershing_counts = pershing.groupby('ym')['abs_amount'].agg(list).reset_index()
    def check_triplet(amts):
        try:
            rounded = [round(float(a), 2) for a in (amts or [])]
        except Exception:
            rounded = []
        return (rounded.count(500.00), rounded.count(3000.00))
    if pershing_counts.empty:
    # nothing to check, return an empty frame with expected columns
        pershing_issues = pershing_counts.assign(cnt_500=pd.Series(dtype=int),
                                             cnt_3000=pd.Series(dtype=int)).head(0)
    else:
        counts = pershing_counts['abs_amount'].apply(check_triplet).tolist()
        pershing_counts[['cnt_500','cnt_3000']] = pd.DataFrame(counts, index=pershing_counts.index)
        pershing_issues = pershing_counts.query('cnt_500 < 2 or cnt_3000 < 1')
    # 5) Optional: by source file totals (helps when a single statement file is off)
    by_file = None
    if 'source_file' in df_norm.columns:
        by_file = df_norm.groupby('source_file')['amount'].sum().reset_index() \
                         .sort_values('amount')

    return {
        'month_recon': merged.sort_values(['y','m']),
        'sign_anomalies': sign_anom,
        'near_duplicates': dup_candidates,
        'pershing_issues': pershing_issues,
        'by_source_file': by_file
    }

# --- USAGE: df_all is your combined transactions dataframe -----------------
# df_norm = normalize(df_all)
# report = imbalance_summary(df_norm)
# # Write to Excel audit sheet next to your dashboard if you want:
# with pd.ExcelWriter('Chase_Budget_Audit.xlsx', engine='xlsxwriter') as xw:
#     report['month_recon'].to_excel(xw, 'Month_Recon', index=False)
#     report['sign_anomalies'].to_excel(xw, 'Sign_Anomalies', index=False)
#     report['near_duplicates'].to_excel(xw, 'Near_Duplicates', index=False)
#     report['pershing_issues'].to_excel(xw, 'Pershing_Check', index=False)
#     if report['by_source_file'] is not None:
#         report['by_source_file'].to_excel(xw, 'By_Source_File', index=False)
