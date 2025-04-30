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

# ── Sidebar: Filter-Vorschau & Kennzahlen ────────────────────────────────────
st.sidebar.header("📊 Übersicht")

if non_num:
    st.sidebar.subheader("Filter-Vorschau")
    filter_counts = (
        pd.DataFrame({
            "Original": df_base[mask_base][fc].value_counts(),
            "Update":   df_upd_base[mask_upd][fc].value_counts()
        })
        .fillna(0)
        .astype(int)
    )
    st.sidebar.dataframe(filter_counts, use_container_width=True)

st.sidebar.subheader("Gesamt-Metriken")
st.sidebar.metric("Zeilen Original",    len(df_base),     delta= len(df_upd_base)-len(df_base))
st.sidebar.metric("GUIDs Original",     df_base["GUID"].nunique(),
                                delta= df_upd_base["GUID"].nunique() - df_base["GUID"].nunique())
if "Fläche" in df_base.columns:
    st.sidebar.metric("Fläche Original", 
                      f"{df_base['Fläche'].sum():.2f}",
                      delta= f"{df_upd_base['Fläche'].sum() - df_base['Fläche'].sum():+.2f}")

# ── Vorschau ────────────────────────────────────────────────────────────────
st.subheader("🔍 Vorschau (erste 5 Zeilen)")
c1, c2 = st.columns(2)
c1.write("Original (gefiltert)")
c1.dataframe(df_base[mask_base].head(5), use_container_width=True)
c2.write("Update (gefiltert)")
c2.dataframe(df_upd_base[mask_upd].head(5), use_container_width=True)

# ── Parameter ────────────────────────────────────────────────────────────────
common         = sorted(set(df_base.columns) & set(df_upd_base.columns))
match_cols     = st.multiselect("GUID-Matching",     common, default=["GUID"])
overwrite_cols = st.multiselect("Zu überschreiben", common, default=["Fläche"])
add_cols       = st.multiselect("Ergänzen",          [c for c in df_upd_base.columns if c not in df_base.columns])
fb_cols        = st.multiselect("Fallback-Matching", [c for c in common if c not in match_cols])
tol            = st.number_input("Toleranz (m²)",    min_value=0.0, max_value=100.0, value=0.5, step=0.1)

# ── Sidebar: Overwrite-Kolonnen-Metriken ────────────────────────────────────
if overwrite_cols:
    st.sidebar.subheader("Spalten O/U")
    for c in overwrite_cols:
        if pd.api.types.is_numeric_dtype(df_base[c]):
            orig_sum = df_base.loc[mask_base, c].sum()
            upd_sum  = df_upd_base.loc[mask_upd, c].sum()
            st.sidebar.metric(f"{c} O/U", f"{orig_sum:.2f}", delta=f"{upd_sum - orig_sum:+.2f}")

# ── Existenzprüfung ───────────────────────────────────────────────────────────
exist_cols = st.multiselect("Existenz-Attribute", [c for c in common if c!="GUID"])
adopt = []
if exist_cols:
    orig_keys = {tuple(r[c] for c in exist_cols) for _, r in df_base.iterrows()}
    cand      = df_upd_base.drop_duplicates(exist_cols)
    miss      = cand[~cand.apply(lambda r: tuple(r[c] for c in exist_cols) in orig_keys, axis=1)]
    if not miss.empty:
        st.subheader("⚠️ Fehlende Kombinationen")
        for i, r in miss.iterrows():
            label = " | ".join(f"{c}: {r[c]}" for c in exist_cols)
            if st.checkbox(f"Übernehmen: {label}", key=f"adopt_{i}"):
                adopt.append(tuple(r[c] for c in exist_cols))

# ── Vergleich starten ────────────────────────────────────────────────────────
if st.button("Vergleich starten"):
    # Output startet als volles Original
    df_out   = df_base.copy()
    full_map = df_upd_base.set_index(match_cols, drop=False)

    # Highlight-Sets & Zähler
    green, orange, yellow, grey, lav = set(), set(), set(), set(), set()
    cnt_ok, cnt_upd, cnt_tol, cnt_new = 0, 0, 0, 0

    # gefilterte Indizes
    idxs = df_base.index[mask_base]

    # Fortschrittsbalken
    total = len(idxs) + (len(idxs) if fb_cols else 0) + len(full_map)
    pbar  = st.progress(0); step = 0

    # ─ A) GUID-Matching (vektorisiert) ────────────────────────────────────────
    df_filt   = df_base[mask_base].reset_index()
    upd_cols  = match_cols + overwrite_cols + add_cols
    upd_filt  = df_upd_base[mask_upd][upd_cols].drop_duplicates(subset=match_cols)
    m_guid    = df_filt.merge(upd_filt, on=match_cols, how="left", suffixes=("", "_upd"))

    for c in overwrite_cols:
        upd_col = f"{c}_upd"
        if upd_col in m_guid.columns:
            # Update-Werte
            m_val = m_guid[upd_col].notna()
            m_diff= m_val & (m_guid[c] != m_guid[upd_col])
            rows = m_guid.loc[m_diff, "index"]
            df_out.loc[rows, f"{c} zuvor"] = df_out.loc[rows, c]
            df_out.loc[rows, c]            = m_guid.loc[m_diff, upd_col]
            orange.update(zip(rows + 2, [df_out.columns.get_loc(c) + 1] * len(rows)))
            cnt_upd += int(m_diff.sum())

            # Exakte ohne Änderung
            m_eq = m_val & (m_guid[c] == m_guid[upd_col])
            rows0= m_guid.loc[m_eq, "index"]
            green.update(zip(rows0 + 2, [df_out.columns.get_loc(c) + 1] * len(rows0)))
            cnt_ok += int(m_eq.sum())

    step += len(m_guid); pbar.progress(min(1.0, step/total))

    # ─ B) Fallback + Toleranz (vektorisiert) ──────────────────────────────────
    if fb_cols:
        df_filt2  = df_base[mask_base].reset_index()
        fb_cols_all = fb_cols + overwrite_cols + add_cols
        upd_fb    = df_upd_base[mask_upd][fb_cols_all].drop_duplicates(subset=fb_cols)
        m_fb      = df_filt2.merge(upd_fb, on=fb_cols, how="left", suffixes=("", "_upd"))

        # Diffs nur für 'Fläche'
        if "Fläche" in overwrite_cols:
            diff = (m_fb["Fläche_upd"] - m_fb["Fläche"]).abs()
        else:
            diff = pd.Series(False, index=m_fb.index)

        # innerhalb Toleranz → gelb
        for c in overwrite_cols:
            upd_col = f"{c}_upd"
            if upd_col in m_fb.columns:
                m_t = m_fb[upd_col].notna() & (m_fb[c] != m_fb[upd_col]) & (diff <= tol)
                rows_t = m_fb.loc[m_t, "index"]
                df_out.loc[rows_t, f"{c} zuvor"] = df_out.loc[rows_t, c]
                df_out.loc[rows_t, c]            = m_fb.loc[m_t, upd_col]
                yellow.update(zip(rows_t + 2, [df_out.columns.get_loc(c) + 1] * len(rows_t)))
                cnt_tol += int(m_t.sum())

        # außerhalb Toleranz → lavendel
        m_ex = m_fb["Fläche_upd"].notna() & (diff > tol)
        rows_ex = m_fb.loc[m_ex, "index"]
        for col in fb_cols:
            lav.update(zip(rows_ex + 2, [df_out.columns.get_loc(col) + 1] * len(rows_ex)))

        step += len(m_fb); pbar.progress(min(1.0, step/total))

    # ─ C) Neueinträge ────────────────────────────────────────────────────────
    existing = {
        tuple(df_out.loc[i, c] for c in match_cols)
        for i in df_out.index
    }
    new_rows = []
    for key in full_map.index.unique():
        if key not in existing:
            u = full_map.loc[key]
            if isinstance(u, pd.DataFrame): u = u.iloc[0]
            if exist_cols:
                ek = tuple(u[c] for c in exist_cols)
                if ek not in adopt:
                    continue
            d = u.to_dict()
            for c in overwrite_cols:
                d[f"{c} zuvor"] = None
            new_rows.append(d)

    if new_rows:
        df_new = pd.DataFrame(new_rows)
        start = len(df_out)
        df_out = pd.concat([df_out, df_new], ignore_index=True)
        for idx in range(start, start + len(df_new)):
            for ci in range(len(df_out.columns)):
                grey.add((idx + 2, ci + 1))
        cnt_new = len(df_new)

    # ── Spalten neu ordnen ───────────────────────────────────────────────────
    base = list(df_base.columns)
    order = []
    for c in base:
        order.append(c)
        pv = f"{c} zuvor"
        if pv in df_out.columns:
            order.append(pv)
    for c in df_out.columns:
        if c not in order:
            order.append(c)
    df_out = df_out[order]

    # ── Export + Styling ─────────────────────────────────────────────────────
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df_out.to_excel(writer, sheet_name="Vergleich", index=False)
        wb, ws = writer.book, writer.sheets["Vergleich"]
        ws.freeze_panes = "A2"
        mc, mr = len(df_out.columns), len(df_out) + 1
        ws.auto_filter.ref = f"A1:{get_column_letter(mc)}{mr}"
        fmt = {
            "green":  PatternFill(start_color="CCFFCC", fill_type="solid"),
            "orange": PatternFill(start_color="FFD966", fill_type="solid"),
            "yellow": PatternFill(start_color="FFFF00", fill_type="solid"),
            "grey":   PatternFill(start_color="DDDDDD", fill_type="solid"),
            "lav":    PatternFill(start_color="E6E6FA", fill_type="solid"),
        }
        for (r, c) in green:  ws.cell(row=r, column=c).fill = fmt["green"]
        for (r, c) in orange: ws.cell(row=r, column=c).fill = fmt["orange"]
        for (r, c) in yellow: ws.cell(row=r, column=c).fill = fmt["yellow"]
        for (r, c) in grey:   ws.cell(row=r, column=c).fill = fmt["grey"]
        for (r, c) in lav:    ws.cell(row=r, column=c).fill = fmt["lav"]
    buf.seek(0)

    # ── Zusammenfassung ───────────────────────────────────────────────────────
    st.markdown("### Zusammenfassung")
    st.success (f"🟧 GUID-Updates: {cnt_upd}")
    st.success (f"🟩 GUID unverändert: {cnt_ok}")
    st.success (f"🟨 Fallback ≤{tol} m²: {cnt_tol}")
    st.success (f"⬜ Neueinträge: {cnt_new}")
    st.success (f"🔷 Fallback ohne Änderung: {len(lav)}")

    # ── Download ─────────────────────────────────────────────────────────────
    ds   = datetime.now().strftime("%y.%m.%d")
    name = orig_file.name.replace(".xlsx", "")
    st.download_button("📥 Datei herunterladen",
                       buf.getvalue(),
                       file_name=f"Updated_{ds}_{name}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
