import streamlit as st
import pandas as pd
import io
from datetime import datetime
from openpyxl.styles import PatternFill, Font
from openpyxl.utils import get_column_letter
from pandas.api.types import is_numeric_dtype

# ── App-Layout ──────────────────────────────────────────────────────────────
st.set_page_config(page_title="IFC Räume Excel Vergleich", layout="wide")
st.title("🔍 Excel Vergleichstool für IFC-Räume")

# ── Upload ───────────────────────────────────────────────────────────────────
orig_file = st.file_uploader("Original Excel hochladen", type=["xlsx"])
upd_file  = st.file_uploader("Update Excel hochladen",   type=["xlsx"])
if not (orig_file and upd_file):
    st.stop()

# ── Einlesen ─────────────────────────────────────────────────────────────────
df_base     = pd.read_excel(orig_file)
df_upd_base = pd.read_excel(upd_file)

# ── Filter bestimmen ─────────────────────────────────────────────────────────
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

df     = df_base[mask_base].reset_index(drop=True)
df_upd = df_upd_base[mask_upd].reset_index(drop=True)

# ── Sidebar & Vorschau ──────────────────────────────────────────────────────
st.sidebar.header("📊 Übersicht")
if non_num:
    cnts = pd.DataFrame({
        "Original": df_base[mask_base][fc].value_counts(),
        "Update":   df_upd_base[mask_upd][fc].value_counts()
    }).fillna(0).astype(int)
    st.sidebar.subheader("Filter-Vorschau")
    st.sidebar.dataframe(cnts, use_container_width=True)
st.sidebar.metric("Zeilen Original", len(df_base), delta=len(df_upd_base)-len(df_base))
st.sidebar.metric("GUIDs Original", df_base["GUID"].nunique(),
                   delta=df_upd_base["GUID"].nunique()-df_base["GUID"].nunique())
if "Fläche" in df_base.columns:
    st.sidebar.metric("Fläche Original", f"{df_base['Fläche'].sum():.2f}",
                      delta=f"{df_upd_base['Fläche'].sum()-df_base['Fläche'].sum():+.2f}")

st.subheader("🔍 Vorschau (5 Zeilen gefiltert)")
c1, c2 = st.columns(2)
c1.write("Original"); c1.dataframe(df.head(5), use_container_width=True)
c2.write("Update");   c2.dataframe(df_upd.head(5), use_container_width=True)

# ── Parameter ────────────────────────────────────────────────────────────────
common         = sorted(set(df_base.columns)&set(df_upd_base.columns))
match_cols     = st.multiselect("GUID-Matching", common, default=["GUID"])
overwrite_cols = st.multiselect("Zu überschreiben", common, default=["Fläche"])
add_cols       = st.multiselect("Ergänzen", [c for c in df_upd_base.columns if c not in df_base.columns])
fb_cols        = st.multiselect("Fallback-Matching", [c for c in common if c not in match_cols])
tol            = st.number_input("Toleranz (m²)", 0.0, 100.0, 0.5, 0.1)

# ── Existenz-Prüfung ─────────────────────────────────────────────────────────
exist_cols = st.multiselect("Existenz-Attribute", [c for c in common if c!="GUID"])
adopt = []
if exist_cols:
    orig_keys = {tuple(r[c] for c in exist_cols) for _,r in df_base.iterrows()}
    cand      = df_upd_base.drop_duplicates(subset=exist_cols)
    miss      = cand[~cand.apply(lambda r: tuple(r[c] for c in exist_cols) in orig_keys, axis=1)]
    if not miss.empty:
        st.subheader("⚠️ Fehlende Kombinationen")
        for i,r in miss.iterrows():
            lbl = " | ".join(f"{c}: {r[c]}" for c in exist_cols)
            if st.checkbox(f"Übernehmen: {lbl}", key=f"adopt_{i}"):
                adopt.append(tuple(r[c] for c in exist_cols))

# ── Vergleich starten ────────────────────────────────────────────────────────
if st.button("Vergleich starten"):
    df_out      = df_base.copy()
    full_map    = df_upd.set_index(match_cols, drop=False)
    updated_rows= set()
    highlight   = {"green":set(),"orange":set(),"yellow":set(),"grey":set(),"lav":set()}

    # GUID-Matching
    df_f    = df_base[mask_base].reset_index()
    upd_cols= match_cols+overwrite_cols+add_cols
    upd_f   = df_upd[upd_cols].drop_duplicates(subset=match_cols, keep="last")
    m_g     = df_f.merge(upd_f, on=match_cols, how="left", suffixes=("","_upd"))
    for c in overwrite_cols:
        uc = f"{c}_upd"
        if uc in m_g:
            m_val  = m_g[uc].notna()
            m_diff = m_val & (m_g[c]!=m_g[uc])
            rows   = m_g.loc[m_diff,"index"].values
            df_out.loc[rows,f"{c} zuvor"] = df_out.loc[rows,c].values
            df_out.loc[rows,c]            = m_g.loc[m_diff,uc].values
            for i in rows: highlight["orange"].update({(i+2,df_out.columns.get_loc(c)+1)}); updated_rows.add(i)
            m_eq   = m_val & (m_g[c]==m_g[uc])
            rows0  = m_g.loc[m_eq,"index"].values
            for i in rows0: highlight["green"].update({(i+2,df_out.columns.get_loc(c)+1)}); updated_rows.add(i)

    # Fallback + Toleranz
    if fb_cols:
        df_f2  = df_base[mask_base].reset_index()
        fb_all = fb_cols+overwrite_cols+add_cols
        upd_fb = df_upd[fb_all].drop_duplicates(subset=fb_cols, keep="last")
        m_fb   = df_f2.merge(upd_fb,on=fb_cols,how="left",suffixes=("","_upd"))
        d = (m_fb["Fläche_upd"]-m_fb["Fläche"]).abs() if "Fläche" in overwrite_cols else pd.Series(False,index=m_fb.index)
        for c in overwrite_cols:
            uc = f"{c}_upd"
            m_t = m_fb[uc].notna()&(m_fb[c]!=m_fb[uc])&(d<=tol)
            rows= m_fb.loc[m_t,"index"].values
            df_out.loc[rows,f"{c} zuvor"] = df_out.loc[rows,c].values
            df_out.loc[rows,c]            = m_fb.loc[m_t,uc].values
            for i in rows: highlight["yellow"].update({(i+2,df_out.columns.get_loc(c)+1)}); updated_rows.add(i)
        m_ex    = m_fb["Fläche_upd"].notna()&(d>tol)
        rows_ex = m_fb.loc[m_ex,"index"].values
        for c in fb_cols:
            for i in rows_ex: highlight["lav"].update({(i+2,df_out.columns.get_loc(c)+1)}); updated_rows.add(i)

    # Neueinträge
    exist = {tuple(df_out.loc[i,c] for c in match_cols) for i in df_out.index}
    new   = []
    for key in full_map.index.unique():
        if key not in exist:
            u = full_map.loc[key]
            if isinstance(u,pd.DataFrame): u=u.iloc[0]
            if exist_cols and tuple(u[c] for c in exist_cols) not in adopt: continue
            dct = u.to_dict()
            for c in overwrite_cols: dct[f"{c} zuvor"] = None
            new.append(dct)
    if new:
        df_new = pd.DataFrame(new)
        st_idx = len(df_out)
        df_out = pd.concat([df_out,df_new],ignore_index=True)
        for i in range(st_idx,st_idx+len(df_new)):
            for col in range(len(df_out.columns)): highlight["grey"].add((i+2,col+1))
            updated_rows.add(i)

    # Spalte “Updated”
    today = datetime.today().strftime("%Y-%m-%d")
    df_out["Updated"] = ""
    for i in updated_rows: df_out.at[i,"Updated"] = today

    # Spalten neu ordnen
    base  = list(df_base.columns); order=[]
    for c in base:
        order += [c, f"{c} zuvor"] if f"{c} zuvor" in df_out else [c]
    order += ["Updated"]
    order += [c for c in df_out.columns if c not in order]
    df_out = df_out[order]

    # Export + Styling mit Strikethrough für andere Zahlen
    buf = io.BytesIO()
    with pd.ExcelWriter(buf,engine="openpyxl") as w:
        df_out.to_excel(w,sheet_name="Vergleich",index=False)
        wb,ws = w.book,w.sheets["Vergleich"]
        ws.freeze_panes="A2"
        mc,mr = len(df_out.columns),len(df_out)+1
        ws.auto_filter.ref = f"A1:{get_column_letter(mc)}{mr}"
        fmt = {
            "green":  PatternFill(start_color="CCFFCC",fill_type="solid"),
            "orange": PatternFill(start_color="FFD966",fill_type="solid"),
            "yellow": PatternFill(start_color="FFFF00",fill_type="solid"),
            "grey":   PatternFill(start_color="DDDDDD",fill_type="solid"),
            "lav":    PatternFill(start_color="E6E6FA",fill_type="solid"),
        }
        # apply fills
        for colset in fmt:
            for (r,c) in highlight[colset]:
                ws.cell(r,c).fill = fmt[colset]
        # strike through other numeric cells in updated rows, except GUID and updated cells
        num_cols = [i+1 for i,c in enumerate(df_out.columns) 
                    if is_numeric_dtype(df_out[c]) and c!="GUID":
        updated_cells = set().union(*highlight.values())
        for r in [i+2 for i in updated_rows]:
            for c in num_cols:
                if (r,c) not in updated_cells:
                    cell = ws.cell(row=r,column=c)
                    cell.font = Font(strike=True)
    buf.seek(0)

    # Zusammenfassung
    st.markdown("### Zusammenfassung")
    st.success(f"🟧 GUID-Updates: {len(highlight['orange'])}")
    st.success(f"🟩 GUID unverändert: {len(highlight['green'])}")
    st.success(f"🟨 Fallback ≤{tol} m²: {len(highlight['yellow'])}")
    st.success(f"⬜ Neueinträge: {len(highlight['grey'])}")
    st.success(f"🔷 Fallback ohne Änderung: {len(highlight['lav'])}")

    # Download
    ds   = datetime.now().strftime("%y.%m.%d")
    name = orig_file.name.replace(".xlsx","")
    st.download_button("📥 Herunterladen",buf.getvalue(),
                       file_name=f"Updated_{ds}_{name}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
