import streamlit as st
import pandas as pd
import io
from datetime import datetime
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

# 📁 App-Layout
st.set_page_config(page_title="IFC Räume Excel Vergleich", layout="wide")
st.title("🔍 Excel Vergleichstool für IFC-Räume")

# 📤 Datei-Uploads
uploaded_original = st.file_uploader("Original Excel hochladen", type=["xlsx"])
uploaded_update   = st.file_uploader("Update Excel hochladen",   type=["xlsx"])

if uploaded_original and uploaded_update:
    try:
        # 📑 Einlesen der Daten
        df_original = pd.read_excel(uploaded_original)
        df_update   = pd.read_excel(uploaded_update)

        # ▶️ Sidebar: Schnellüberblick
        st.sidebar.subheader("Datenüberblick")
        n_orig      = df_original.shape[0]
        n_upd       = df_update.shape[0]
        orig_guids  = df_original['GUID'].nunique() if 'GUID' in df_original.columns else 0
        upd_guids   = df_update['GUID'].nunique()   if 'GUID' in df_update.columns   else 0
        # Fläche als Zahl sicherstellen
        if 'Fläche' in df_original.columns:
            df_original['Fläche'] = pd.to_numeric(df_original['Fläche'], errors='coerce')
        if 'Fläche' in df_update.columns:
            df_update['Fläche']   = pd.to_numeric(df_update['Fläche'],   errors='coerce')
        orig_total  = df_original['Fläche'].sum() if 'Fläche' in df_original.columns else 0
        upd_total   = df_update['Fläche'].sum()   if 'Fläche' in df_update.columns   else 0

        st.sidebar.write(f"Original: {n_orig} Zeilen, {orig_guids} GUIDs")
        st.sidebar.write(f"Update:   {n_upd} Zeilen, {upd_guids} GUIDs")
        st.sidebar.write(f"Gesamtfläche Original: {orig_total:.2f}")
        st.sidebar.write(f"Gesamtfläche Update:   {upd_total:.2f}")

        # ▶️ Vorschau
        st.subheader("🔍 Vorschau der Datensätze")
        col1, col2 = st.columns(2)
        with col1:
            st.write("**Original (erste 5 Zeilen)**")
            st.dataframe(df_original.head(5), use_container_width=True)
        with col2:
            st.write("**Update (erste 5 Zeilen)**")
            st.dataframe(df_update.head(5),   use_container_width=True)

        # Gemeinsame Spalten & GUID-Check
        common_cols = sorted(set(df_original.columns) & set(df_update.columns))
        if "GUID" not in common_cols:
            st.error("Beide Dateien müssen die Spalte 'GUID' enthalten.")
            st.stop()

        # ▶️ Summen nach Gruppe
        st.subheader("📊 Flächen-Summen nach Gruppe")
        group_col = st.selectbox(
            "Spalte zum Gruppieren auswählen",
            options=common_cols,
            index=common_cols.index('Haus') if 'Haus' in common_cols else 0
        )
        if 'Fläche' not in common_cols:
            st.warning("Spalte 'Fläche' nicht gefunden; Summen nicht verfügbar.")
        else:
            orig_sum = (
                df_original
                .groupby(group_col)['Fläche']
                .sum()
                .reset_index(name='Summe Original')
            )
            upd_sum = (
                df_update
                .groupby(group_col)['Fläche']
                .sum()
                .reset_index(name='Summe Update')
            )
            df_sum = pd.merge(orig_sum, upd_sum, on=group_col, how='outer').fillna(0)
            st.dataframe(df_sum, use_container_width=True)

        # ▶️ Matching-Parameter
        match_columns     = st.multiselect("🔗 Spalten für Matching auswählen",       options=common_cols,               default=["GUID"])
        overwrite_columns = st.multiselect("📝 Spalten zum Überschreiben bei Match", options=common_cols,               default=["Fläche"])
        additional_columns= st.multiselect("➕ Spalten aus Update ergänzen",          options=[c for c in df_update.columns if c not in df_original.columns])

        # ▶️ Vergleichs-Logik
        if st.button("Vergleich starten"):
            df_orig_copy = df_original.copy()
            df_upd_copy  = df_update.copy().set_index(match_columns, drop=False)

            green_cells  = set()
            yellow_cells = set()
            unchanged    = 0
            updated      = 0
            new_keys     = []

            # 1) Bestehende GUIDs vergleichen
            for idx, row in df_orig_copy.iterrows():
                key = tuple(row[col] for col in match_columns)
                try:
                    upd_row = df_upd_copy.loc[key]
                    if isinstance(upd_row, pd.DataFrame):
                        upd_row = upd_row.iloc[0]
                    changed = False

                    # Überschreiben
                    for col in overwrite_columns:
                        new_val = upd_row[col]
                        orig_val= row[col]
                        if pd.notna(new_val) and orig_val != new_val:
                            df_orig_copy.at[idx, f"{col} zuvor"] = orig_val
                            df_orig_copy.at[idx, col]             = new_val
                            # Excel-Koordinaten speichern (Header=1, Data ab Zeile 2)
                            excel_r = idx + 2
                            excel_c = df_orig_copy.columns.get_loc(col) + 1
                            green_cells.add((excel_r, excel_c))
                            changed = True

                    # Zusätzliche Spalten
                    for col in additional_columns:
                        df_orig_copy.at[idx, col] = upd_row[col]

                    if changed:
                        updated += 1
                    else:
                        unchanged += 1

                except KeyError:
                    # kein Match → wird später als neue GUID behandelt
                    continue

            # 2) Neue GUIDs anhängen & 'Fläche' gelb markieren
            orig_keys = {
                tuple(row[col] for col in match_columns)
                for _, row in df_original.iterrows()
            }
            upd_keys  = set(df_upd_copy.index)
            missing   = upd_keys - orig_keys

            for key in missing:
                try:
                    new_row = df_upd_copy.loc[key]
                    if isinstance(new_row, pd.DataFrame):
                        new_row = new_row.iloc[0]
                    d = new_row.to_dict()
                    for col in overwrite_columns:
                        d[f"{col} zuvor"] = None
                    df_orig_copy = pd.concat([df_orig_copy, pd.DataFrame([d])], ignore_index=True)
                    new_idx = df_orig_copy.shape[0] - 1
                    # Excel-Koordinate für 'Fläche'
                    if 'Fläche' in df_orig_copy.columns:
                        excel_r = new_idx + 2
                        excel_c = df_orig_copy.columns.get_loc('Fläche') + 1
                        yellow_cells.add((excel_r, excel_c))
                    new_keys.append(key)
                except Exception:
                    continue

            # 3) Excel-Export mit Styling
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df_orig_copy.to_excel(writer, sheet_name="Vergleich", index=False)
                wb  = writer.book
                ws  = writer.sheets["Vergleich"]
                # Freeze + Autofilter
                ws.freeze_panes = "A2"
                max_c = df_orig_copy.shape[1]
                max_r = df_orig_copy.shape[0] + 1
                ws.auto_filter.ref = f"A1:{get_column_letter(max_c)}{max_r}"
                # Füllungen
                green = PatternFill(start_color="CCFFCC", end_color="CCFFCC", fill_type="solid")
                yellow= PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
                for (r, c) in green_cells:
                    ws.cell(row=r, column=c).fill = green
                for (r, c) in yellow_cells:
                    ws.cell(row=r, column=c).fill = yellow

            buffer.seek(0)

            # ▶️ Zusammenfassung
            st.markdown("### Zusammenfassung")
            st.success(f"🔁 Aktualisierte Einträge: {updated}")
            st.info   (f"✅ Exakte Matches ohne Änderung: {unchanged}")
            st.warning(f"➕ Neue GUIDs ergänzt: {len(new_keys)}")

            # ▶️ Download
            date_str     = datetime.now().strftime("%y.%m.%d")
            orig_name    = uploaded_original.name.replace(".xlsx", "")
            export_name  = f"Updated_{date_str}_{orig_name}.xlsx"
            st.download_button(
                label="📥 Aktualisierte Datei herunterladen",
                data=buffer.getvalue(),
                file_name=export_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Fehler: {e}")
