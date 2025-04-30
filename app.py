import streamlit as st
import pandas as pd
import io
from datetime import datetime
from openpyxl.styles import PatternFill, Font
from openpyxl.utils import get_column_letter
from pandas.api.types import is_numeric_dtype

# ── App-Layout & Tabs ────────────────────────────────────────────────────────
st.set_page_config(page_title="IFC Räume Excel Vergleich", layout="wide")
st.title("🔍 Excel Vergleichstool für IFC-Räume")

tab1, tab2 = st.tabs(["🔄 Vergleich", "🗑️ Filter-Ersetzen"])

# ── Hilfsfunktion: Einlesen & Filter ──────────────────────────────────────────
def load_and_filter(orig, upd, key_suffix):
    dfb = pd.read_excel(orig)
    dfn = pd.read_excel(upd)
    non_num = dfb.select_dtypes(exclude="number").columns.tolist()
    if non_num:
        fc = st.selectbox("Filter-Spalte", non_num, key=f"filter_col_{key_suffix}")
        vals = dfb[fc].dropna().unique().tolist()
        sel  = st.multiselect("Filter-Werte", vals, default=vals, key=f"filter_val_{key_suffix}")
        mask_b = dfb[fc].isin(sel)
        mask_u = dfn[fc].isin(sel)
    else:
        mask_b = pd.Series(True, index=dfb.index)
        mask_u = pd.Series(True, index=dfn.index)
    return dfb, dfn, mask_b, mask_u, non_num, fc

# ── Tab 1: Vergleich ──────────────────────────────────────────────────────────
with tab1:
    orig_file1 = st.file_uploader("Original Excel hochladen", key="orig1", type=["xlsx"])
    upd_file1  = st.file_uploader("Update Excel hochladen",   key="upd1",  type=["xlsx"])
    if not (orig_file1 and upd_file1):
        st.info("Bitte beide Dateien hochladen.")
        st.stop()

    df_base, df_upd_base, mask_base, mask_upd, non_num, fc = load_and_filter(orig_file1, upd_file1, "tab1")

    # Parameter
    common = sorted(set(df_base.columns) & set(df_upd_base.columns))
    match_cols = st.multiselect("GUID-Matching", common, default=["GUID"], key="match1")
    overwrite_cols = st.multiselect("Zu überschreiben", common, default=["Fläche"], key="overwrite1")
    add_cols = st.multiselect("Ergänzen", [c for c in df_upd_base.columns if c not in df_base.columns], key="add1")
    fb_cols = st.multiselect("Fallback-Matching", [c for c in common if c not in match_cols], key="fb1")
    tol = st.number_input("Toleranz (m²)", min_value=0.0, max_value=100.0, value=0.5, step=0.1, key="tol1")

    # Vorbereitung
    df_out = df_base.copy()
    # Build full_map using selected match_cols
    full_map = df_upd_base[mask_upd].set_index(match_cols, drop=False)

    # Highlight- und Zähler
    green, orange, yellow, grey, lav = set(), set(), set(), set(), set()
    cnt_ok = cnt_upd = cnt_tol = cnt_new = 0
    updated_rows = set()

    if st.button("Vergleich starten", key="run1"):
        # GUID-Matching (vektorisiert)
        df_f = df_base[mask_base].reset_index()
        upd_cols = match_cols + overwrite_cols + add_cols
        upd_f = df_upd_base[mask_upd][upd_cols].drop_duplicates(subset=match_cols, keep="last")
        m_g = df_f.merge(upd_f, on=match_cols, how="left", suffixes=("", "_upd"))
        for c in overwrite_cols:
            col_upd = f"{c}_upd"
            if col_upd in m_g:
                m_val = m_g[col_upd].notna()
                m_diff = m_val & (m_g[c] != m_g[col_upd])
                rows = m_g.loc[m_diff, "index"].values
                df_out.loc[rows, f"{c} zuvor"] = df_out.loc[rows, c].values
                df_out.loc[rows, c] = m_g.loc[m_diff, col_upd].values
                for i in rows:
                    orange.add((i+2, df_out.columns.get_loc(c)+1)); updated_rows.add(i)
                cnt_upd += int(m_diff.sum())
                m_eq = m_val & (m_g[c] == m_g[col_upd])
                rows0 = m_g.loc[m_eq, "index"].values
                for i in rows0:
                    green.add((i+2, df_out.columns.get_loc(c)+1)); updated_rows.add(i)
                cnt_ok += int(m_eq.sum())
        # Fallback + Toleranz
        if fb_cols:
            df_f2 = df_base[mask_base].reset_index()
            fb_all = fb_cols + overwrite_cols + add_cols
            upd_fb = df_upd_base[mask_upd][fb_all].drop_duplicates(subset=fb_cols, keep="last")
            m_fb = df_f2.merge(upd_fb, on=fb_cols, how="left", suffixes=("", "_upd"))
            if overwrite_cols:
                first = overwrite_cols[0]
                d = (m_fb[f"{first}_upd"] - m_fb[first]).abs()
            else:
                d = pd.Series(False, index=m_fb.index)
            for c in overwrite_cols:
                col_upd = f"{c}_upd"
                if col_upd in m_fb:
                    m_t = m_fb[col_upd].notna() & (m_fb[c] != m_fb[col_upd]) & (d <= tol)
                    rows = m_fb.loc[m_t, "index"].values
                    df_out.loc[rows, f"{c} zuvor"] = df_out.loc[rows, c].values
                    df_out.loc[rows, c] = m_fb.loc[m_t, col_upd].values
                    for i in rows:
                        yellow.add((i+2, df_out.columns.get_loc(c)+1)); updated_rows.add(i)
                    cnt_tol += int(m_t.sum())
            # lavendel: fallback-match ohne Änderung (optional)
        # Neueinträge anhängen
        existing = {tuple(df_out.loc[i, c] for c in match_cols) for i in df_out.index}
        new_rows = []
        for key in full_map.index.unique():
            if key not in existing:
                u = full_map.loc[key]
                if isinstance(u, pd.DataFrame): u = u.iloc[0]
                d = u.to_dict()
                for c in overwrite_cols: d[f"{c} zuvor"] = None
                new_rows.append(d)
        if new_rows:
            df_new = pd.DataFrame(new_rows)
            start = len(df_out)
            df_out = pd.concat([df_out, df_new], ignore_index=True)
            for idx in range(start, len(df_out)):
                for j in range(len(df_out.columns)):
                    grey.add((idx+2, j+1)); updated_rows.add(idx)
            cnt_new = len(new_rows)
        # Updated-Spalte
        today = datetime.today().strftime("%Y-%m-%d")
        df_out["Updated"] = ""
        for i in updated_rows: df_out.at[i, "Updated"] = today
        # Spalten anordnen
        base_cols = list(df_base.columns)
        order = []
        for c in base_cols:
            order.append(c)
            pv = f"{c} zuvor"
            if pv in df_out: order.append(pv)
        order.append("Updated")
        order += [c for c in df_out.columns if c not in order]
        df_out = df_out[order]
        # Export + Styling
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            df_out.to_excel(writer, sheet_name="Vergleich", index=False)
            wb, ws = writer.book, writer.sheets["Vergleich"]
            ws.freeze_panes = "A2"
            mc, mr = len(df_out.columns), len(df_out)+1
            ws.auto_filter.ref = f"A1:{get_column_letter(mc)}{mr}"
            fmt = {
                "green": PatternFill(start_color="CCFFCC", fill_type="solid"),
                "orange": PatternFill(start_color="FFD966", fill_type="solid"),
                "yellow": PatternFill(start_color="FFFF00", fill_type="solid"),
                "grey": PatternFill(start_color="DDDDDD", fill_type="solid"),
            }
            for (r,c) in green:  ws.cell(r,c).fill = fmt["green"]
            for (r,c) in orange: ws.cell(r,c).fill = fmt["orange"]
            for (r,c) in yellow: ws.cell(r,c).fill = fmt["yellow"]
            for (r,c) in grey:   ws.cell(r,c).fill = fmt["grey"]
        buf.seek(0)
        st.download_button(
            "📥 Datei herunterladen",
            buf.getvalue(),
            file_name=f"Updated_{datetime.now():%y.%m.%d}_{orig_file1.name.replace('.xlsx','')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# ── Tab 2: Filter-Ersetzen ─────────────────────────────────────────────────────
with tab2:
    orig_file2 = st.file_uploader("Original Excel hochladen", key="orig2", type=["xlsx"])
    upd_file2  = st.file_uploader("Update Excel hochladen",   key="upd2",  type=["xlsx"])
    if not (orig_file2 and upd_file2):
        st.info("Bitte beide Dateien hochladen.")
        st.stop()

    df_base2, df_upd2, mask_b2, mask_u2, non_num2, fc2 = load_and_filter(orig_file2, upd_file2, "tab2")
    st.markdown("Alle **gefilterten** Original-Zeilen werden gelöscht und durch die **gefilterten** Update-Zeilen ersetzt.")
    if st.button("🔁 Filter Ersetzen", key="replace_btn"):
        df_out2 = df_base2.loc[~mask_b2].reset_index(drop=True)
        df_replace = df_upd2[mask_u2].reset_index(drop=True)
        start2 = len(df_out2)
        df_out2 = pd.concat([df_out2, df_replace], ignore_index=True)
        grey2 = {(i+2,j+1) for i in range(start2,len(df_out2)) for j in range(len(df_out2.columns))}
        df_out2["Updated"] = ""
        today2 = datetime.today().strftime("%Y-%m-%d")
        for i in range(start2,len(df_out2)): df_out2.at[i,"Updated"] = today2
        base2 = list(df_base2.columns)
        order2=[]
