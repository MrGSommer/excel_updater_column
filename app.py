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

# ── Einlesen & Filter ────────────────────────────────────────────────────────
df_base     = pd.read_excel(orig_file)
df_upd_base = pd.read_excel(upd_file)

# optionaler Filter
non_num = df_base.select_dtypes(exclude="number").columns.tolist()
if non_num:
    fc  = st.selectbox("Filter-Spalte", non_num)
    sel = st.multiselect("Filter-Werte", df_base[fc].dropna().unique(), default=df_base[fc].dropna().unique())
    df     = df_base   [df_base[fc].isin(sel)].reset_index(drop=True)
    df_upd = df_upd_base[df_upd_base[fc].isin(sel)].reset_index(drop=True)
else:
    df     = df_base.copy()
    df_upd = df_upd_base.copy()

# ── Kennzahlen ──────────────────────────────────────────────────────────────
for tmp in (df, df_upd):
    if "Fläche" in tmp.columns:
        tmp["Fläche"] = pd.to_numeric(tmp["Fläche"], errors="coerce")

st.sidebar.header("📊 Übersicht")
st.sidebar.metric("Zeilen O/U", len(df), len(df_upd)-len(df))
st.sidebar.metric("GUIDs O/U", df["GUID"].nunique(), df_upd["GUID"].nunique()-df["GUID"].nunique())
if "Fläche" in df.columns:
    st.sidebar.metric("Fläche O/U", f"{df['Fläche'].sum():.2f}", f"{df_upd['Fläche'].sum()-df['Fläche'].sum():+.2f}")

# ── Vorschau ────────────────────────────────────────────────────────────────
st.subheader("🔍 Vorschau (5 Zeilen)")
c1, c2 = st.columns(2)
c1.dataframe(df.head(5), use_container_width=True)
c2.dataframe(df_upd.head(5), use_container_width=True)

# ── Spalten & Summen ─────────────────────────────────────────────────────────
common = sorted(set(df.columns)&set(df_upd.columns))
if "GUID" not in common:
    st.error("Spalte 'GUID' fehlt."); st.stop()

st.subheader("📊 Summen nach Gruppe")
grp = st.selectbox("Gruppieren nach", common, index=common.index("Haus") if "Haus" in common else 0)
if "Fläche" in common:
    so = df.groupby(grp)["Fläche"].sum().reset_index(name="Summe O")
    su = df_upd.groupby(grp)["Fläche"].sum().reset_index(name="Summe U")
    st.dataframe(pd.merge(so,su,on=grp,how="outer").fillna(0), use_container_width=True)

# ── Parameter ────────────────────────────────────────────────────────────────
match_cols     = st.multiselect("GUID-Matching", common, default=["GUID"])
overwrite_cols = st.multiselect("Zu überschreiben", common, default=["Fläche"])
add_cols       = st.multiselect("Ergänzen", [c for c in df_upd.columns if c not in df.columns])
fb_cols        = st.multiselect("Fallback-Matching", [c for c in common if c not in match_cols])
tol            = st.number_input("Toleranz (m²)", 0.0, 10.0, 0.5, 0.1)

# ── Existenzprüfung ─────────────────────────────────────────────────────────
exist_cols = st.multiselect("Existenz-Attribute", [c for c in common if c!="GUID"])
adopt = []
if exist_cols:
    orig_keys = {tuple(r[c] for c in exist_cols) for _,r in df.iterrows()}
    cand = df_upd.drop_duplicates(exist_cols)
    miss = cand[~cand.apply(lambda r: tuple(r[c] for c in exist_cols) in orig_keys, axis=1)]
    if not miss.empty:
        st.subheader("⚠️ Fehlende Kombinationen")
        for i,r in miss.iterrows():
            label = " | ".join(f"{c}:{r[c]}" for c in exist_cols)
            if st.checkbox(f"Übernehmen: {label}", key=f"adopt_{i}"):
                adopt.append(tuple(r[c] for c in exist_cols))

# ── Vergleich ───────────────────────────────────────────────────────────────
if st.button("Vergleich starten"):
    df_out     = df.copy()
    full_map   = df_upd.set_index(match_cols, drop=False)
    fb_map     = df_upd.set_index(fb_cols,   drop=False) if fb_cols else None
    matched    = set()
    idx_guid   = set()
    idx_fb     = set()

    # Highlight-Sets
    green = set()   # GUID unverändert
    orange = set()  # GUID update
    yellow = set()  # Fallback mit Toleranz
    grey = set()    # neue Einträge
    lav = set()     # Fallback ohne Änderung

    # Counters
    c_guid_upd = 0
    c_guid_ok  = 0
    c_fb_tol   = 0
    c_new      = 0

    # Progress
    total = len(df_out)+len(full_map)
    pbar  = st.progress(0); step=0

    # 1) GUID-Matching
    for i in range(len(df_out)):
        row = df_out.loc[i]
        key = tuple(row[c] for c in match_cols)
        if key in full_map.index:
            upd = full_map.loc[key]
            if isinstance(upd,pd.DataFrame): upd=upd.iloc[0]
            # immer Series
            updated = False
            for c in overwrite_cols:
                nv, ov = upd[c], row[c]
                if pd.notna(nv) and nv!=ov:
                    df_out.at[i, f"{c} zuvor"] = ov
                    df_out.at[i, c]            = nv
                    orange.add((i+2, df_out.columns.get_loc(c)+1))
                    updated = True
            for c in add_cols:
                df_out.at[i,c] = upd[c]
            if updated:
                c_guid_upd += 1
            else:
                # exakter Match ohne Änderung
                for c in overwrite_cols:
                    green.add((i+2, df_out.columns.get_loc(c)+1))
                c_guid_ok += 1
            matched.add(key)
            idx_guid.add(i)
        step+=1; pbar.progress(min(1.0,step/total))

    # 2) Fallback mit Toleranz
    if fb_cols:
        orig_un = {i for i in range(len(df_out)) if i not in idx_guid}
        upd_un  = [k for k in full_map.index if k not in matched]
        for i in orig_un:
            row = df_out.loc[i]
            fk = tuple(row[c] for c in fb_cols)
            if fk not in fb_map.index: continue
            fbr = fb_map.loc[fk]
            if isinstance(fbr,pd.DataFrame) and len(fbr)==1: fbr=fbr.iloc[0]
            if not isinstance(fbr,pd.Series): continue
            # gleiche Fallback-Attribute
            updated=False
            # Toleranz-Check
            diff = abs(row["Fläche"] - fbr["Fläche"])
            if diff <= tol:
                for c in overwrite_cols:
                    ov = row[c]
                    nv = fbr[c]
                    if pd.notna(nv) and nv!=ov:
                        df_out.at[i, f"{c} zuvor"] = ov
                        df_out.at[i, c]            = nv
                        yellow.add((i+2, df_out.columns.get_loc(c)+1))
                        updated=True
                for c in add_cols:
                    df_out.at[i,c] = fbr[c]
                if updated:
                    c_fb_tol += 1
                else:
                    # Fallback ohne Änderung
                    for c in fb_cols:
                        lav.add((i+2, df_out.columns.get_loc(c)+1))
                matched.add(tuple(fbr[c] for c in match_cols))
        step+=len(upd_un); pbar.progress(1.0)

    # 3) Neueinträge
    for k in full_map.index.unique():
        if k not in matched:
            u = full_map.loc[k]
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
                grey.add((ni+2, ci+1))
            c_new += 1

    # Spalten neu anordnen
    base = list(df_base.columns)
    order=[]
    for c in base:
        order.append(c)
        pv=f"{c} zuvor"
        if pv in df_out.columns: order.append(pv)
    for c in df_out.columns:
        if c not in order: order.append(c)
    df_out = df_out[order]

    # 4) Export mit Styling
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df_out.to_excel(writer, sheet_name="Vergleich", index=False)
        wb, ws = writer.book, writer.sheets["Vergleich"]
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
        # Zeichne in sinnvoller Reihenfolge
        for (r,c) in green:  ws.cell(row=r, column=c).fill = fmt["green"]
        for (r,c) in orange: ws.cell(row=r, column=c).fill = fmt["orange"]
        for (r,c) in yellow: ws.cell(row=r, column=c).fill = fmt["yellow"]
        for (r,c) in grey:   ws.cell(row=r, column=c).fill = fmt["grey"]
        for (r,c) in lav:    ws.cell(row=r, column=c).fill = fmt["lav"]
    buf.seek(0)

    # ── Zusammenfassung ───────────────────────────────────────────────────────
    st.markdown("### Zusammenfassung")
    st.success(f"🟧 GUID-Updates: {c_guid_upd}")
    st.success(f"🟩 GUID unverändert: {c_guid_ok}")
    st.success(f"🟨 Fallback ≤{tol} m²: {c_fb_tol}")
    st.success(f"⬜ Neueinträge: {c_new}")
    st.success(f"🔷 Fallback ohne Änderung: {len(lav)}")

    # ── Download ─────────────────────────────────────────────────────────────
    ds = datetime.now().strftime("%y.%m.%d")
    name = orig_file.name.replace(".xlsx","")
    st.download_button("📥 Datei herunterladen",
                       buf.getvalue(),
                       file_name=f"Updated_{ds}_{name}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
