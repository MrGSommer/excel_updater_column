import streamlit as st
import pandas as pd
import io
from datetime import datetime
import openpyxl
from openpyxl.styles import PatternFill

# 📁 App-Layout
st.set_page_config(page_title="Excel Vergleich", layout="wide")
st.title("🔍 Excel Vergleichstool für IFC-Räume")

# 📤 Datei-Uploads
uploaded_original = st.file_uploader("Original Excel hochladen", type=["xlsx"])
uploaded_update = st.file_uploader("Update Excel hochladen", type=["xlsx"])

if uploaded_original and uploaded_update:
    # 📑 Einlesen der Daten
    df_original = pd.read_excel(uploaded_original)
    df_update = pd.read_excel(uploaded_update)

    common_cols = list(set(df_original.columns) & set(df_update.columns))
    common_cols.sort()

    st.subheader("🔍 Vorschau: Original-Datei")
    st.dataframe(df_original.head(5), use_container_width=True)

    st.subheader("🔍 Vorschau: Update-Datei")
    st.dataframe(df_update.head(5), use_container_width=True)
    
    if "GUID" not in common_cols:
        st.error("Beide Dateien müssen die Spalte 'GUID' enthalten.")
        st.stop()

    match_columns = st.multiselect("🔗 Spalten für Matching auswählen", options=common_cols, default=["GUID"])
    overwrite_columns = st.multiselect("📝 Spalten zum Überschreiben bei Match", options=common_cols, default=["Fläche"])
    additional_columns = st.multiselect("➕ Spalten aus Update ergänzen", options=[col for col in df_update.columns if col not in df_original.columns])

    if st.button("Vergleich starten"):
        df_original_copy = df_original.copy()
        df_update_copy = df_update.copy()

        # Index auf Matching-Spalten setzen
        df_update_copy.set_index(match_columns, inplace=True)
        df_original_copy["GUID"] = df_original_copy["GUID"].astype(str)
        changed_cells = set()

        unchanged_count = 0
        updated_count = 0
        new_guids = []

        for idx, row in df_original_copy.iterrows():
            key = tuple(row[col] for col in match_columns)
            try:
                update_row = df_update_copy.loc[key]
                if isinstance(update_row, pd.DataFrame):
                    update_row = update_row.iloc[0]
                any_change = False

                for col in overwrite_columns:
                    if pd.notna(update_row[col]) and row[col] != update_row[col]:
                        df_original_copy.at[idx, f"{col} zuvor"] = row[col]
                        df_original_copy.at[idx, col] = update_row[col]
                        changed_cells.add((idx + 1, df_original_copy.columns.get_loc(col)))
                        any_change = True

                for col in additional_columns:
                    df_original_copy.at[idx, col] = update_row[col]

                if any_change:
                    updated_count += 1
                else:
                    unchanged_count += 1

            except KeyError:
                continue  # Kein Match – bleibt unverändert

        # Neue GUIDs finden
        original_keys = set(tuple(row[col] for col in match_columns) for _, row in df_original.iterrows())
        update_keys = set(df_update_copy.index)
        missing_keys = update_keys - original_keys

        for key in missing_keys:
            try:
                new_row = df_update_copy.loc[key]
                if isinstance(new_row, pd.DataFrame):
                    new_row = new_row.iloc[0]
                new_dict = new_row.to_dict()

                for col in overwrite_columns:
                    new_dict[f"{col} zuvor"] = None

                df_original_copy = pd.concat([df_original_copy, pd.DataFrame([new_dict])], ignore_index=True)
                new_guids.append(key)

                # Färbe komplette neue Zeile
                for i, col in enumerate(df_original_copy.columns):
                    changed_cells.add((df_original_copy.shape[0], i))
            except Exception:
                continue

        # 💾 Export vorbereiten
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_original_copy.to_excel(writer, sheet_name="Vergleich", index=False)
            workbook = writer.book
            worksheet = writer.sheets["Vergleich"]

            fill_green = PatternFill(start_color="CCFFCC", end_color="CCFFCC", fill_type="solid")

            for row_idx, col_idx in changed_cells:
                cell = worksheet.cell(row=row_idx + 1, column=col_idx + 1)
                cell.fill = fill_green

        # 📊 Zusammenfassung
        st.markdown("### Zusammenfassung:")
        st.success(f"🔁 Aktualisierte Einträge: {updated_count}")
        st.info(f"✅ Exakte Matches ohne Änderung: {unchanged_count}")
        st.warning(f"➕ Neue Einträge ergänzt: {len(new_guids)}")

        # ⬇️ Download-Link
        today_str = datetime.today().strftime("%y.%m.%d")
        original_name = uploaded_original.name.replace(".xlsx", "")
        file_name = f"Updated_{today_str}_{original_name}.xlsx"
        st.download_button(
            label="📥 Aktualisierte Datei herunterladen",
            data=output.getvalue(),
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
