import streamlit as st
import pandas as pd
import io
from datetime import datetime
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

# ── App-Layout ──────────────────────────────────────────────────────────────
st.set_page_config(page_title="IFC Räume Excel Vergleich", layout="wide")
st.title("🔍 Excel Vergleichstool für IFC-Räume")

# ── Upload ───────────────────────────────────────────────────────────────────
orig_file = st.file_uploader("Original Excel hochladen", type=["xlsx"])
upd_file  = st.file_uploader("Update Excel hochladen",   type=["xlsx"])
if not (orig_file and upd_file):
    st.stop()

# ── Einlesen ─────────────────────────────────────────────────────────────────
df_base     = pd.read_excel(orig_file)       # komplettes Original
df_upd_base = pd.read_excel(upd_file)        # Update

# ── Filter bestimmen ─────────────────────────────────────────────────────────
non_num = df_base.select_dtypes(exclude="number").columns.tolist()
if non_num:
    fc  = st.selectbox("Filter-Spalte", non_num)
    vals = df_base[fc].dropna().unique()
    sel  = st.multiselect("Filter-Werte", vals, default=vals)
    # Maske, welche Zeilen für Updates gelten
    mask = df_base[fc].isin(sel)
else:
    mask = pd.Series(True, index=df_base.index)

# ── Kennzahlen & Vorschau ────────────────────────────────────────────────────
df      = df_base[mask].reset_index(drop=True)
df_upd  = df_upd_base[df_upd_base[mask.name if non_num else df_base.index.name]].reset_index(drop=True)

for tmp in (df, df_upd):
    if "Fläche" in tmp.columns:
        tmp["Fläche"] = pd.to_numeric(tmp["Fläche"], errors="coerce")

st.sidebar.header("📊 Übersicht")
st.sidebar.metric("Zeilen O/U", len(df_base), len(df_upd_base)-len(df_base))
st.sidebar.metric("GUIDs O/U", df_base["GUID"].nunique(), df_upd_base["GUID"].nunique()-df_base["GUID"].nunique())
if "Fläche" in df_base.columns:
    st.sidebar.metric(
        "Fläche O/U",
        f"{df_base['Fläche'].sum():.2f}",
        f"{df_upd_base['Fläche'].sum()-df_base['Fläche'].sum():+.2f}"
    )

st.subheader("🔍 Vorschau (5 Zeilen)")
c1, c2 = st.columns(2)
c1.dataframe(df_base.head(5), use_container_width=True)
c2.dataframe(df_upd_base.head(5), use_container_width=True)

# ── Matching-Parameter ───────────────────────────────────────────────────────
common         = sorted(set(df_base.columns)&set(df_upd_base.columns))
match_cols     = st.multiselect("GUID-Matching",     common, default=["GUID"])
overwrite_cols = st.multiselect("Zu überschreiben", common, default=["Fläche"])
add_cols       = st.multiselect("Ergänzen",          [c for c in df_upd_base.columns if c not in df_base.columns])
fb_cols        = st.multiselect("Fallback-Matching", [c for c in common if c not in match_cols])
tol            = st.number_input("Toleranz (m²)",    0.0, 10.0, 0.5, 0.1)

# ── Existenzprüfung ───────────────────────────────────────────────────────────
exist_cols = st.multiselect("Existenz-Attribute", [c for c in common if c!="GUID"])
adopt = []
if exist_cols:
    orig_keys = {tuple(r[c] for c in exist_cols) for _,r in df_base.iterrows()}
    cand = df_upd_base.drop_duplicates(exist_cols)
    miss = cand[~cand.apply(lambda r: tuple(r[c] for c in exist_cols) in orig_keys, axis=1)]
    if not miss.empty:
        st.subheader("⚠️ Fehlende Kombinationen")
        for i,r in miss.iterrows():
            lbl = " | ".join(f"{c}:{r[c]}" for c in exist_cols)
            if st.checkbox(f"Übernehmen: {lbl}", key=f"adopt_{i}"):
                adopt.append(tuple(r[c] for c in exist_cols))

# ── Vergleich ───────────────────────────────────────────────────────────────
if st.button("Vergleich starten"):
    # 1: Starte mit vollem Original
    df_out   = df_base.copy()
    full_map = df_upd_base.set_index(match_cols, drop=False)
    fb_map   = df_upd_base.set_index(fb_cols,   drop=False) if fb_cols else None

    # Highlight-Sets & Counters
    green, orange, yellow, grey, lav = set(), set(), set(), set(), set()
    cnt_ok, cnt_upd, cnt_tol, cnt_new = 0,0,0,0

    # Hilfs-Menge gefilterter Indizes
    idxs = df_out.index[mask]

    # Progress
    total = len(idxs) + (1 if fb_cols else 0) + len(full_map)
    pbar  = st.progress(0); step=0

    # --- A) GUID-Matching auf gefilterten Zeilen
    for i in idxs:
        row = df_out.loc[i]
        key = tuple(row[c] for c in match_cols)
        if key in full_map.index:
            upd = full_map.loc[key]
            if isinstance(upd,pd.DataFrame): upd=upd.iloc[0]
            updated=False
            for c in overwrite_cols:
                nv, ov = upd[c], row[c]
                if pd.notna(nv) and nv!=ov:
                    df_out.at[i, f"{c} zuvor"] = ov
                    df_out.at[i, c]            = nv
                    orange.add((i+2, df_out.columns.get_loc(c)+1))
                    updated=True
            for c in add_cols:
                df_out.at[i,c] = upd[c]
            if updated: cnt_upd += 1
            else:
                for c in overwrite_cols:
                    green.add((i+2, df_out.columns.get_loc(c)+1))
                cnt_ok += 1
        step+=1; pbar.progress(min(1.0,step/total))

    # --- B) Fallback + Toleranz auf gefilterten Zeilen
    if fb_cols:
        for i in idxs.difference(df_out.index[list(green)+list(orange)]):
            row = df_out.loc[i]
            fk = tuple(row[c] for c in fb_cols)
            if fk not in fb_map.index: continue
            fbr = fb_map.loc[fk]
            if isinstance(fbr,pd.DataFrame) and len(fbr)==1: fbr=fbr.iloc[0]
            if not isinstance(fbr,pd.Series): continue

            diff = abs(row["Fläche"] - fbr["Fläche"])
            if diff <= tol:
                for c in overwrite_cols:
                    ov = row[c]; nv = fbr[c]
                    if pd.notna(nv) and nv!=ov:
                        df_out.at[i, f"{c} zuvor"] = ov
                        df_out.at[i, c]            = nv
                        yellow.add((i+2, df_out.columns.get_loc(c)+1))
                for c in add_cols:
                    df_out.at[i,c] = fbr[c]
                cnt_tol += 1
            else:
                # Fallback-Match ohne Änderung
                for c in fb_cols:
                    lav.add((i+2, df_out.columns.get_loc(c)+1))
        step+=1; pbar.progress(min(1.0,step/total))

    # --- C) Neueinträge anhängen (nur wenn adopt erlaubt)
    for key in full_map.index.unique():
        if key in full_map.index and key not in {tuple(r[c] for c in match_cols) for _,r in df_out.iterrows()}:
            u = full_map.loc[key]
            if isinstance(u,pd.DataFrame): u=u.iloc[0]
            if exist_cols:
                ek = tuple(u[c] for c in exist_cols)
                if ek not in adopt: continue
            d = u.to_dict()
            for c in overwrite_cols:
                d[f"{c} zuvor"] = None
            df_out = pd.concat([df_out, pd.DataFrame([d])], ignore_index=True)
            ni = len(df_out)-1
            for ci in range(len(df_out.columns)):
                grey.add((ni+2,ci+1))
            cnt_new += 1
    step+=len(full_map); pbar.progress(1.0)

    # ── Spalten neu ordnen ────────────────────────────────────────────────────
    base = list(df_base.columns)
    order=[]
    for c in base:
        order.append(c)
        pv=f"{c} zuvor"
        if pv in df_out.columns: order.append(pv)
    for c in df_out.columns:
        if c not in order: order.append(c)
    df_out = df_out[order]

    # ── Export + Styling ──────────────────────────────────────────────────────
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df_out.to_excel(w, sheet_name="Vergleich", index=False)
        wb, ws = w.book, w.sheets["Vergleich"]
        ws.freeze_panes="A2"
        mc,mr = len(df_out.columns), len(df_out)+1
        ws.auto_filter.ref = f"A1:{get_column_letter(mc)}{mr}"
        fmt = {
            "green":  PatternFill(start_color="CCFFCC", fill_type="solid"),
            "orange": PatternFill(start_color="FFD966", fill_type="solid"),
            "yellow": PatternFill(start_color="FFFF00", fill_type="solid"),
            "grey":   PatternFill(start_color="DDDDDD", fill_type="solid"),
            "lav":    PatternFill(start_color="E6E6FA", fill_type="solid"),
        }
        for (r,c) in green:  ws.cell(r,c).fill = fmt["green"]
        for (r,c) in orange: ws.cell(r,c).fill = fmt["orange"]
        for (r,c) in yellow: ws.cell(r,c).fill = fmt["yellow"]
        for (r,c) in grey:   ws.cell(r,c).fill = fmt["grey"]
        for (r,c) in lav:    ws.cell(r,c).fill = fmt["lav"]
    buf.seek(0)

    # ── Zusammenfassung ───────────────────────────────────────────────────────
    st.markdown("### Zusammenfassung")
    st.success(f"🟧 GUID-Updates: {cnt_upd}")
    st.success(f"🟩 GUID unverändert: {cnt_ok}")
    st.success(f"🟨 Fallback ≤{tol} m²: {cnt_tol}")
    st.success(f"⬜ Neueinträge: {cnt_new}")
    st.success(f"🔷 Fallback ohne Änderung: {len(lav)}")

    # ── Download ─────────────────────────────────────────────────────────────
    ds   = datetime.now().strftime("%y.%m.%d")
    name = orig_file.name.replace(".xlsx","")
    st.download_button("📥 Herunterladen",
                       buf.getvalue(),
                       file_name=f"Updated_{ds}_{name}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
