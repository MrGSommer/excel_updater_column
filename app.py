import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
import os

st.title("Excel Vergleich: Original vs. Update")

# --- Datei Upload ---
original_file = st.file_uploader("Lade das Original Excel hoch", type=["xlsx"])
update_file = st.file_uploader("Lade das Update Excel hoch", type=["xlsx"])

if original_file and update_file:
    try:
        df_original = pd.read_excel(original_file)
        df_update = pd.read_excel(update_file)

        common_columns = [col for col in df_original.columns if col in df_update.columns]
        if not common_columns:
            st.error("Keine gemeinsamen Spalten für Matching gefunden.")
        else:
            st.success("Dateien erfolgreich geladen.")

            match_columns = st.multiselect("Wähle Matching-Spalten", common_columns)
            overwrite_candidates = [col for col in common_columns if df_original[col].dtype == df_update[col].dtype]
            overwrite_columns = st.multiselect("Spalten zum Überschreiben mit '... zuvor'", overwrite_candidates)

            additional_update_columns = [col for col in df_update.columns if col not in df_original.columns]
            additional_columns = st.multiselect("Zusätzliche Spalten aus dem Update einfügen", additional_update_columns)

            if st.button("Vergleich starten"):
                df_merged = df_original.copy()
                df_update["_merge_key"] = df_update[match_columns].astype(str).agg("_".join, axis=1)
                df_merged["_merge_key"] = df_merged[match_columns].astype(str).agg("_".join, axis=1)

                update_dict = df_update.set_index("_merge_key").to_dict(orient="index")
                changed_cells = set()

                for idx, row in df_merged.iterrows():
                    key = row["_merge_key"]
                    if key in update_dict:
                        update_row = update_dict[key]
                        for col in overwrite_columns:
                            original_value = row[col]
                            new_value = update_row.get(col)
                            if pd.notna(new_value) and original_value != new_value:
                                df_merged.at[idx, col + " zuvor"] = original_value
                                df_merged.at[idx, col] = new_value
                                changed_cells.add((idx + 1, df_merged.columns.get_loc(col)))  # Excel-Zeile +1
                        for col in additional_columns:
                            df_merged.at[idx, col] = update_row.get(col)

                df_merged.drop(columns="_merge_key", inplace=True)

                # --- Anzeige ---
                st.dataframe(df_merged)

                # --- Export ---
                towrite = BytesIO()
                with pd.ExcelWriter(towrite, engine='xlsxwriter') as writer:
                    df_merged.to_excel(writer, index=False, sheet_name='Vergleich')
                    workbook = writer.book
                    worksheet = writer.sheets['Vergleich']
                    green_format = workbook.add_format({'bg_color': '#C6EFCE'})
                    for row, col in changed_cells:
                        worksheet.write(row, col, df_merged.iloc[row - 1, col], green_format)

                    worksheet.freeze_panes(1, 0)
                    worksheet.autofilter(0, 0, len(df_merged), len(df_merged.columns) - 1)

                towrite.seek(0)

                # --- Dateiname ---
                date_str = datetime.now().strftime("%y.%m.%d")
                original_name = os.path.splitext(original_file.name)[0]
                export_filename = f"Updated_{date_str}_{original_name}.xlsx"

                st.download_button(
                    label="Vergleich herunterladen mit Formatierung",
                    data=towrite,
                    file_name=export_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    except Exception as e:
        st.error(f"Fehler: {e}")
