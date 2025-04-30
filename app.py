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
    st.info("Bitte beide Dateien hochladen, um zu vergleichen.")
    st.stop()

# ── Einlesen & Filter ────────────────────────────────────────────────────────
df_base = pd.read_excel(orig_file)
df_upd_base = pd.read_excel(upd_file)

# Filtern über eine nicht-numerische Spalte (z.B. Ausmasscode)
non_num = df_base.select_dtypes(exclude="number").columns.tolist()
if non_num:
    fc = st.selectbox("Filter-Spalte", non_num)
    vals = df_base[fc].dropna().unique().tolist()
    sel = st.multiselect("Filter-Werte", vals, default=vals)
    df = df_base[df_base[fc].isin(sel)].reset_index(drop=True)
    df_upd = df_upd_base[df_upd_base[fc].isin(sel)].reset_index(drop=True)
else:
    df = df_base.copy()
    df_upd = df_upd_base.copy()

# ── Sidebar-Kennzahlen ───────────────────────────────────────────────────────
for tmp in (df, df_upd):
    if "Fläche" in tmp.columns:
        tmp["Fläche"] = pd.to_numeric(tmp["Fläche"], errors="coerce")

st.sidebar.header("📊 Kurzübersicht")
st.sidebar.metric("Zeilen O/U", len(df), len(df_upd) - len(df))
st.sidebar.metric("GUIDs O/U", df["GUID"].nunique(), df_upd["GUID"].nunique() - df["GUID"].nunique())
if "Fläche" in df.columns:
    st.sidebar.metric("Summe Fläche O/U",
                       f"{df['Fläche'].sum():.2f}",
                       f"{df_upd['Fläche'].sum() - df['Fläche'].sum():+.2f}")

# ── Vorschau ────────────────────────────────────────────────────────────────
st.subheader("🔍 Vorschau (erste 5 Zeilen)")
c1, c2 = st.columns(2)
with c1:
    st.write("Original")
    st.dataframe(df.head(5), use_container_width=True)
with c2:
    st.write("Update")
    st.dataframe(df_upd.head(5), use_container_width=True)

# ── Spalten prüfen & Summen ─────────────────────────────────────────────────
common = sorted(set(df.columns) & set(df_upd.columns))
if "GUID" not in common:
    st.error("Spalte 'GUID' fehlt in einer der Dateien.")
    st.stop()

st.subheader("📊 Flächen-Summen nach Gruppe")
group = st.selectbox("Gruppieren nach", common, index=common.index("Haus") if "Haus" in common else 0)
if "Fläche" in common:
    so = df.groupby(group)["Fläche"].sum().reset_index(name="Summe O")
    su = df_upd.groupby(group)["Fläche"].sum().reset_index(name="Summe U")
    st.dataframe(pd.merge(so, su, on=group, how="outer").fillna(0), use_container_width=True)

# ── Matching-Parameter ─────────────────────────────────────────────────────
match_cols     = st.multiselect("GUID-Matching", common, default=["GUID"])
overwrite_cols = st.multiselect("Zu überschreiben", common, default=["Fläche"])
add_cols       = st.multiselect("Ergänzen aus Update", [c for c in df_upd.columns if c not in df.columns])
fb_cols        = st.multiselect("Fallback-Matching-Spalten", [c for c in common if c not in match_cols])
tol            = st.number_input("Toleranz (m²)", min_value=0.0, max_value=100.0, value=0.5, step=0.1)

# ── Existenzprüfung ─────────────────────────────────────────────────────────
exist_cols = st.multiselect("Existenz-Attribute (für neue Einträge)", [c for c in common if c != "GUID"])
adopt = []
if exist_cols:
    orig_set = {tuple(r[c] for c in exist_cols) for _, r in df.iterrows()}
    candidates = df_upd.drop_duplicates(exist_cols)
    missing = candidates[~candidates.apply(lambda r: tuple(r[c] for c in exist_cols) in orig_set, axis=1)]
    if not missing.empty:
        st.subheader("⚠️ Fehlende Kombinationen im Original")
        for i, r in missing.iterrows():
            label = " | ".join(f"{c}: {r[c]}" for c in exist_cols)
            if st.checkbox(f"Übernehmen: {label}", key=f"adopt_{i}"):
                adopt.append(tuple(r[c] for c in exist_cols))

# ── Vergleich laufen lassen ──────────────────────────────────────────────────
if st.button("Vergleich starten"):
    df_out = df.copy()
    full_map   = df_upd.set_index(match_cols, drop=False)
    fb_map     = df_upd.set_index(fb_cols,   drop=False) if fb_cols else None

    # Tracking und Highlight-Sets
    matched     = set()
    idx_guid    = set()
    idx_fb      = set()
    green, blue, orange, yellow, red, lav = set(), set(), set(), set(), set(), set()
    cnt_guid = cnt_fb = cnt_tol = cnt_new = cnt_unc = 0

    total = len(df_out) + len(full_map)
    pbar  = st.progress(0); step = 0

    # 1) GUID-Matching
    for i in range(len(df_out)):
        row = df_out.loc[i]
        key = tuple(row[c] for c in match_cols)
        if key in full_map.index:
            upd = full_map.loc[key]
            if isinstance(upd, pd.DataFrame): upd = upd.iloc[0]
            # jetzt eine Series
            used = False
            for c in overwrite_cols:
                nv, ov = upd[c], row[c]
                if pd.notna(nv) and nv != ov:
                    df_out.at[i, f"{c} zuvor"] = ov
                    df_out.at[i, c]            = nv
                    green.add((i+2, df_out.columns.get_loc(c)+1))
                    used = True
            for c in add_cols:
                df_out.at[i, c] = upd[c]
            cnt_guid += 1 if used else (cnt_unc := cnt_unc+1)
            matched.add(key)
            idx_guid.add(i)
        step += 1; pbar.progress(min(1.0, step/total))

    # 2) Fallback-Matching
    if fb_cols:
        for i in range(len(df_out)):
            if i in idx_guid: continue
            row = df_out.loc[i]
            fkey = tuple(row[c] for c in fb_cols)
            if fb_map is None or fkey not in fb_map.index:
                continue
            fbr = fb_map.loc[fkey]
            if isinstance(fbr, pd.DataFrame) and len(fbr)==1:
                fbr = fbr.iloc[0]
            if isinstance(fbr, pd.Series):
                fullk = tuple(fbr[c] for c in match_cols)
                if fullk in full_map.index and fullk not in matched:
                    used_fb = False
                    for c in overwrite_cols:
                        nv, ov = fbr[c], row[c]
                        if pd.notna(nv) and nv != ov:
                            df_out.at[i, f"{c} zuvor"] = ov
                            df_out.at[i, c]            = nv
                            blue.add((i+2, df_out.columns.get_loc(c)+1))
                            used_fb = True
                    for c in add_cols:
                        df_out.at[i, c] = fbr[c]
                    if not used_fb:
                        # Fallback-Match ohne Änderung → Lavendel
                        for c in fb_cols:
                            lav.add((i+2, df_out.columns.get_loc(c)+1))
                    cnt_fb += 1 if used_fb else (cnt_unc := cnt_unc+1)
                    matched.add(fullk)
                    idx_fb.add(i)
            step += 1; pbar.progress(min(1.0, step/total))

        # 3) Toleranz-Matching (nur innerhalb gleicher fb-Kombi)
        orig_un = {i for i in range(len(df_out)) if i not in idx_guid|idx_fb}
        upd_un  = [k for k in full_map.index if k not in matched]
        pairs   = []
        for i in orig_un:
            o = df_out.loc[i]
            for k in upd_un:
                u = full_map.loc[k]
                if isinstance(u, pd.DataFrame): u = u.iloc[0]
                if all(o[c]==u[c] for c in fb_cols):
                    diff = abs(o["Fläche"] - u["Fläche"])
                    if diff <= tol:
                        pairs.append((i,k,diff))
                    else:
                        red.add((i+2, df_out.columns.get_loc("Fläche")+1))
        pairs.sort(key=lambda x: x[2])
        for i,k,_ in pairs:
            if i not in idx_guid|idx_fb and k not in matched:
                u = full_map.loc[k]
                if isinstance(u, pd.DataFrame): u = u.iloc[0]
                for c in overwrite_cols:
                    ov = df_out.at[i,c]
                    df_out.at[i, f"{c} zuvor"] = ov
                    df_out.at[i, c]            = u[c]
                    orange.add((i+2, df_out.columns.get_loc(c)+1))
                for c in add_cols:
                    df_out.at[i,c] = u[c]
                cnt_tol += 1
                matched.add(k)
        # Entferne Rot, wenn Orange drüber liegt
        red -= orange
        step += len(upd_un); pbar.progress(1.0)

    # 4) Neueinträge
    for k in full_map.index.unique():
        if k not in matched:
            u  = full_map.loc[k]
            if isinstance(u, pd.DataFrame): u = u.iloc[0]
            if exist_cols:
                ek = tuple(u[c] for c in exist_cols)
                if ek not in adopt:
                    continue
            d = u.to_dict()
            for c in overwrite_cols:
                d[f"{c} zuvor"] = None
            df_out = pd.concat([df_out, pd.DataFrame([d])], ignore_index=True)
            ni = len(df_out)-1
            for ci in range(len(df_out.columns)):
                yellow.add((ni+2, ci+1))
            cnt_new += 1

    # Spalten neu anordnen: "... zuvor" immer rechts
    base = list(df_base.columns)
    order=[]
    for c in base:
        order.append(c)
        pv=f"{c} zuvor"
        if pv in df_out.columns:
            order.append(pv)
    for c in df_out.columns:
        if c not in order: order.append(c)
    df_out = df_out[order]

    # ── Export mit Styling ─────────────────────────────────────────────────
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df_out.to_excel(w, sheet_name="Vergleich", index=False)
        wb, ws = w.book, w.sheets["Vergleich"]
        ws.freeze_panes = "A2"
        max_c, max_r = len(df_out.columns), len(df_out)+1
        ws.auto_filter.ref = f"A1:{get_column_letter(max_c)}{max_r}"
        fmt = {
            "g": PatternFill(start_color="CCFFCC", fill_type="solid"),
            "b": PatternFill(start_color="ADD8E6", fill_type="solid"),
            "o": PatternFill(start_color="FFD966", fill_type="solid"),
            "y": PatternFill(start_color="FFFF00", fill_type="solid"),
            "r": PatternFill(start_color="FFC7CE", fill_type="solid"),
            "l": PatternFill(start_color="E6E6FA", fill_type="solid"),
        }
        # Zeichnen in der Reihenfolge: Grün, Blau, Rot, Orange, Gelb, Lavendel
        for (r,c) in green:    ws.cell(r,c).fill = fmt["g"]
        for (r,c) in blue:     ws.cell(r,c).fill = fmt["b"]
        for (r,c) in red:      ws.cell(r,c).fill = fmt["r"]
        for (r,c) in orange:   ws.cell(r,c).fill = fmt["o"]
        for (r,c) in yellow:   ws.cell(r,c).fill = fmt["y"]
        for (r,c) in lav:      ws.cell(r,c).fill = fmt["l"]
    buf.seek(0)

    # ── Zusammenfassung ───────────────────────────────────────────────────────
    st.markdown("### Zusammenfassung")
    st.success(f"🔁 GUID-Updates: {cnt_guid}")
    st.info   (f"🔷 Fallback-Updates: {cnt_fb}")
    st.info   (f"🔶 Toleranz-Updates ≤{tol} m²: {cnt_tol}")
    st.warning(f"❌ Toleranz überschritten >{tol} m²: {len(red)}")
    st.info   (f"🔮 Fallback ohne Änderung: {len(lav)}")
    st.info   (f"➕ Neueinträge übernommen: {cnt_new}")
    st.info   (f"✅ Unverändert: {cnt_unc}")

    # ── Download ─────────────────────────────────────────────────────────────
    name = orig_file.name.replace(".xlsx","")
    ds   = datetime.now().strftime("%y.%m.%d")
    st.download_button(
        "📥 Datei herunterladen",
        buf.getvalue(),
        file_name=f"Updated_{ds}_{name}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

