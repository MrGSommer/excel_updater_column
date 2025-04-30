import streamlit as st
import pandas as pd
import io
from datetime import datetime
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
from pandas.api.types import is_numeric_dtype

# ── App-Layout ──────────────────────────────────────────────────────────────
st.set_page_config(page_title="IFC Räume Excel Vergleich", layout="wide")
st.title("🔍 Excel Vergleichstool für IFC-Räume")

# ── Datei-Uploads ───────────────────────────────────────────────────────────
orig_file = st.file_uploader("Original Excel hochladen", type=["xlsx"])
upd_file  = st.file_uploader("Update Excel hochladen",   type=["xlsx"])
if not (orig_file and upd_file):
    st.stop()

# ── Einlesen und Filter ──────────────────────────────────────────────────────
df_base     = pd.read_excel(orig_file)
df_upd_base = pd.read_excel(upd_file)
non_num = df_base.select_dtypes(exclude="number").columns.tolist()
if non_num:
    fc   = st.selectbox("Filter-Spalte", non_num)
    vals = df_base[fc].dropna().unique().tolist()
    sel  = st.multiselect("Filter-Werte", vals, default=vals)
    mask_base = df_base[fc].isin(sel)
    mask_upd  = df_upd_base[fc].isin(sel)
else:
    mask_base = pd.Series(True, index=df_base.index)
    mask_upd  = pd.Series(True, index=df_upd_base.index)

# Gefilterte Teiltabellen
df = df_base[mask_base].reset_index(drop=True)
df_upd = df_upd_base[mask_upd].reset_index(drop=True)

# ── Modus-Auswahl ────────────────────────────────────────────────────────────
mode = st.radio("Modus wählen", ["Vergleich", "Filter-Ersetzen"])

# Gemeinsame Parameter
df_out = None
highlight = {"green":set(), "orange":set(), "yellow":set(), "grey":set(), "lav":set()}
updated_rows = set()

if mode == "Vergleich":
    st.subheader("🛠 Vergleichs-Parameter")
    common = sorted(set(df_base.columns) & set(df_upd_base.columns))
    match_cols     = st.multiselect("GUID-Matching", common, default=["GUID"])
    overwrite_cols = st.multiselect("Zu überschreiben", common, default=["Fläche"])
    add_cols       = st.multiselect("Ergänzen", [c for c in df_upd_base.columns if c not in df_base.columns])
    fb_cols        = st.multiselect("Fallback-Matching", [c for c in common if c not in match_cols])
    tol            = st.number_input("Toleranz (m²)", 0.0, 100.0, 0.5, 0.1)
    if st.button("Vergleich ausführen"):
        df_out = df_base.copy()
        full_map = df_upd_base[mask_upd].set_index(match_cols, drop=False)
        # GUID-Matching
        df_f = df_base[mask_base].reset_index()
        upd_cols = match_cols + overwrite_cols + add_cols
        upd_f = df_upd_base[mask_upd][upd_cols].drop_duplicates(subset=match_cols, keep="last")
        m_g = df_f.merge(upd_f, on=match_cols, how="left", suffixes=("","_upd"))
        for c in overwrite_cols:
            cu = f"{c}_upd"
            if cu in m_g:
                val = m_g[cu].notna()
                diff = val & (m_g[c] != m_g[cu])
                idx = m_g.loc[diff, "index"].values
                df_out.loc[idx, f"{c} zuvor"] = df_out.loc[idx, c].values
                df_out.loc[idx, c] = m_g.loc[diff, cu].values
                for i in idx: highlight["orange"].add((i+2, df_out.columns.get_loc(c)+1)); updated_rows.add(i)
                eq = val & (m_g[c] == m_g[cu])
                idx0 = m_g.loc[eq, "index"].values
                for i in idx0: highlight["green"].add((i+2, df_out.columns.get_loc(c)+1)); updated_rows.add(i)
        # Fallback-Matching
        if fb_cols:
            df_fb = df_base[mask_base].reset_index()
            fb_all = fb_cols + overwrite_cols + add_cols
            upd_fb = df_upd_base[mask_upd][fb_all].drop_duplicates(subset=fb_cols, keep="last")
            m_fb = df_fb.merge(upd_fb, on=fb_cols, how="left", suffixes=("","_upd"))
            if overwrite_cols:
                dcol = overwrite_cols[0]
                d = (m_fb[f"{dcol}_upd"] - m_fb[dcol]).abs()
            else:
                d = pd.Series(False, index=m_fb.index)
            for c in overwrite_cols:
                cu = f"{c}_upd"
                if cu in m_fb:
                    cond = m_fb[cu].notna() & (m_fb[c] != m_fb[cu]) & (d <= tol)
                    idx2 = m_fb.loc[cond, "index"].values
                    df_out.loc[idx2, f"{c} zuvor"] = df_out.loc[idx2, c].values
                    df_out.loc[idx2, c] = m_fb.loc[cond, cu].values
                    for i in idx2: highlight["yellow"].add((i+2, df_out.columns.get_loc(c)+1)); updated_rows.add(i)
        # Neueinträge
        exist = {tuple(df_out.loc[i,c] for c in match_cols) for i in df_out.index}
        new = []
        for key in full_map.index.unique():
            if key not in exist:
                row = full_map.loc[key]
                if isinstance(row, pd.DataFrame): row=row.iloc[0]
                drow = row.to_dict()
                for c in overwrite_cols: drow[f"{c} zuvor"] = None
                new.append(drow)
        if new:
            df_new = pd.DataFrame(new)
            s = len(df_out)
            df_out = pd.concat([df_out, df_new], ignore_index=True)
            for i in range(s, len(df_out)):
                for j in range(len(df_out.columns)): highlight["grey"].add((i+2,j+1)); updated_rows.add(i)
        # Updated-Spalte
        today = datetime.now().strftime("%Y-%m-%d")
        df_out["Updated"] = ""
        for i in updated_rows: df_out.at[i, "Updated"] = today
        # Spaltenordnung
        cols = []
        for c in df_base.columns:
            cols.append(c)
            pv = f"{c} zuvor"
            if pv in df_out: cols.append(pv)
        cols.append("Updated")
        cols += [c for c in df_out.columns if c not in cols]
        df_out = df_out[cols]
        # Export + Styling
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            df_out.to_excel(writer, sheet_name="Vergleich", index=False)
            wb, ws = writer.book, writer.sheets["Vergleich"]
            ws.freeze_panes = "A2"
            mc, mr = len(df_out.columns), len(df_out)+1
            ws.auto_filter.ref = f"A1:{get_column_letter(mc)}{mr}"
            fills = {"green":"CCFFCC","orange":"FFD966","yellow":"FFFF00","grey":"DDDDDD"}
            for k, color in fills.items():
                fill = PatternFill(start_color=color, fill_type="solid")
                for (r,c) in highlight[k]: ws.cell(r,c).fill = fill
        buf.seek(0)
        st.download_button("📥 Excel herunterladen", buf.getvalue(), file_name=f"Updated_{datetime.now():%y.%m.%d}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

elif mode == "Filter-Ersetzen":
    st.subheader("🗑️ Filter-Ersetzen")
    if st.button("Ersetzen ausführen"):
        df_out = df_base.loc[~mask_base].reset_index(drop=True)
        df_rep = df_upd_base[mask_upd].reset_index(drop=True)
        s2 = len(df_out)
        df_out = pd.concat([df_out, df_rep], ignore_index=True)
        grey2 = {(i+2,j+1) for i in range(s2,len(df_out)) for j in range(len(df_out.columns))}
        df_out["Updated"] = ""
        t2 = datetime.now().strftime("%Y-%m-%d")
        for i in range(s2,len(df_out)): df_out.at[i,"Updated"]=t2
        # Spalten neu ordnen wie oben
        cols2 = []
        for c in df_base.columns:
            cols2.append(c)
            pv = f"{c} zuvor"
            if pv in df_out: cols2.append(pv)
        cols2.append("Updated")
        cols2 += [c for c in df_out.columns if c not in cols2]
        df_out = df_out[cols2]
        buf2 = io.BytesIO()
        with pd.ExcelWriter(buf2, engine="openpyxl") as w2:
            df_out.to_excel(w2, sheet_name="Ersetzt", index=False)
            wb2, ws2 = w2.book, w2.sheets["Ersetzt"]
            ws2.freeze_panes = "A2"
            mc2, mr2 = len(df_out.columns), len(df_out)+1
            ws2.auto_filter.ref = f"A1:{get_column_letter(mc2)}{mr2}"
            fill = PatternFill(start_color="DDDDDD", fill_type="solid")
            for (r,c) in grey2: ws2.cell(r,c).fill = fill
        buf2.seek(0)
        st.download_button("📥 Excel herunterladen", buf2.getvalue(), file_name=f"Ersetzt_{datetime.now():%y.%m.%d}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
