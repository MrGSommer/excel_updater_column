import streamlit as st
import pandas as pd
import io
from datetime import datetime
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

# ── App-Layout & Tabs ────────────────────────────────────────────────────────
st.set_page_config(page_title="IFC Räume Excel Vergleich", layout="wide")
st.title("🔍 Excel Vergleichstool für IFC-Räume")

tab1, tab2 = st.tabs(["🔄 Vergleich", "🗑️ Filter-Ersetzen"])

# ── Gemeinsame Uploads ───────────────────────────────────────────────────────
with tab1:
    orig_file = st.file_uploader("Original Excel hochladen", key="orig1", type=["xlsx"])
    upd_file  = st.file_uploader("Update Excel hochladen",   key="upd1",  type=["xlsx"])
with tab2:
    # dieselben Uploader, nur andere Keys
    orig_file2 = st.file_uploader("Original Excel hochladen", key="orig2", type=["xlsx"])
    upd_file2  = st.file_uploader("Update Excel hochladen",   key="upd2",  type=["xlsx"])

# ── Allgemeines Einlesen & Filter ─────────────────────────────────────────────
def load_and_filter(orig, upd):
    dfb = pd.read_excel(orig)
    dfn = pd.read_excel(upd)
    # Filter-UI
    non_num = dfb.select_dtypes(exclude="number").columns.tolist()
    if non_num:
        fc   = st.selectbox("Filter-Spalte", non_num)
        vals = dfb[fc].dropna().unique().tolist()
        sel  = st.multiselect("Filter-Werte", vals, default=vals)
        mask_b = dfb[fc].isin(sel)
        mask_u = dfn[fc].isin(sel)
    else:
        mask_b = pd.Series(True, index=dfb.index)
        mask_u = pd.Series(True, index=dfn.index)
    return dfb, dfn, mask_b, mask_u, non_num, fc

# ── Tab 1: Vergleich ──────────────────────────────────────────────────────────
with tab1:
    if not (orig_file and upd_file):
        st.info("Bitte beide Dateien hochladen.")
        st.stop()

    df_base, df_upd_base, mask_base, mask_upd, non_num, fc = load_and_filter(orig_file, upd_file)

    # hier kommt Dein kompletter Vergleichs-Code hinein, wie er ist,
    # ab Zeile "df_base = pd.read_excel(orig_file)" bis zum Download-Button.
    # Er liest df_base, df_upd_base, mask_base, mask_upd, ...
    # und füllt die highlight-Sets green, orange, yellow, grey, lav sowie updated_rows.
    #
    # Zum Schluss exportierst Du df_out genau wie bisher:
    # st.download_button(...)

# ── Tab 2: Filter-Ersetzen ─────────────────────────────────────────────────────
with tab2:
    if not (orig_file2 and upd_file2):
        st.info("Bitte beide Dateien hochladen.")
        st.stop()

    # gleiche Filter-Logik
    df_base, df_upd_base, mask_base, mask_upd, non_num, fc = load_and_filter(orig_file2, upd_file2)

    st.markdown(
        "Alle **gefilterten** Original-Zeilen werden gelöscht und durch "
        "die **gefilterten** Update-Zeilen ersetzt. "
        "Markierungen für neue Zeilen bleiben erhalten."
    )

    if st.button("🔁 Filter Ersetzen"):
        # 1) Nicht-gefilterte Original-Zeilen übernehmen
        df_out = df_base.loc[~mask_base].copy().reset_index(drop=True)

        # 2) Gefilterte Update-Zeilen anhängen
        df_replace = df_upd_base[mask_upd].copy().reset_index(drop=True)
        start = len(df_out)
        df_out = pd.concat([df_out, df_replace], ignore_index=True)

        # 3) Highlight: alle ersetzten Zeilen als 'grey'
        grey = set()
        for i in range(start, len(df_out)):
            for col in range(len(df_out.columns)):
                grey.add((i+2, col+1))

        # 4) Spalte "Updated" ergänzen
        today = datetime.today().strftime("%Y-%m-%d")
        df_out["Updated"] = ""
        for i in range(start, len(df_out)):
            df_out.at[i, "Updated"] = today

        # 5) Spalten neu ordnen (wie gehabt)
        base = list(df_base.columns)
        order = []
        for c in base:
            order.append(c)
            pv = f"{c} zuvor"
            if pv in df_out.columns:
                order.append(pv)
        order.append("Updated")
        for c in df_out.columns:
            if c not in order:
                order.append(c)
        df_out = df_out[order]

        # 6) Export + Styling (nur grey)
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            df_out.to_excel(writer, sheet_name="Ersetzt", index=False)
            wb, ws = writer.book, writer.sheets["Ersetzt"]
            ws.freeze_panes = "A2"
            mc, mr = len(df_out.columns), len(df_out) + 1
            ws.auto_filter.ref = f"A1:{get_column_letter(mc)}{mr}"
            fill_grey = PatternFill(start_color="DDDDDD", fill_type="solid")
            for (r, c) in grey:
                ws.cell(row=r, column=c).fill = fill_grey
        buf.seek(0)

        st.success(f"Filter-Ersetzen erfolgreich: {len(df_replace)} Zeilen angefügt.")
        st.download_button(
            "📥 Ersetztes File herunterladen",
            buf.getvalue(),
            file_name=f"Ersetzt_{datetime.now():%y.%m.%d}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
