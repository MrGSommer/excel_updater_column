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
        # 📑 Daten einlesen
        df_original = pd.read_excel(uploaded_original)
        df_update   = pd.read_excel(uploaded_update)

        # — Sidebar: Schnellüberblick —
        st.sidebar.subheader("Datenüberblick")
        for df in (df_original, df_update):
            if 'Fläche' in df.columns:
                df['Fläche'] = pd.to_numeric(df['Fläche'], errors='coerce')
        st.sidebar.write(f"Original: {df_original.shape[0]} Zeilen, {df_original['GUID'].nunique()} GUIDs")
        st.sidebar.write(f"Update:   {df_update.shape[0]} Zeilen, {df_update['GUID'].nunique()} GUIDs")
        st.sidebar.write(f"Summe Fläche Original: {df_original['Fläche'].sum():.2f}")
        st.sidebar.write(f"Summe Fläche Update:   {df_update['Fläche'].sum():.2f}")

        # — Vorschau —
        st.subheader("🔍 Vorschau der Datensätze")
        c1, c2 = st.columns(2)
        with c1:
            st.write("**Original (erste 5 Zeilen)**")
            st.dataframe(df_original.head(5), use_container_width=True)
        with c2:
            st.write("**Update (erste 5 Zeilen)**")
            st.dataframe(df_update.head(5),   use_container_width=True)

        # — Gemeinsame Spalten prüfen —
        common_cols = sorted(set(df_original.columns) & set(df_update.columns))
        if "GUID" not in common_cols:
            st.error("Beide Dateien müssen die Spalte 'GUID' enthalten.")
            st.stop()

        # — Flächensummen nach Gruppe —
        st.subheader("📊 Flächen-Summen nach Gruppe")
        group_col = st.selectbox(
            "Spalte zum Gruppieren auswählen",
            options=common_cols,
            index=common_cols.index('Haus') if 'Haus' in common_cols else 0
        )
        if 'Fläche' in common_cols:
            orig_sum = df_original.groupby(group_col)['Fläche']\
                                  .sum().reset_index(name='Summe Original')
            upd_sum  = df_update.groupby(group_col)['Fläche']\
                                 .sum().reset_index(name='Summe Update')
            df_sum   = pd.merge(orig_sum, upd_sum, on=group_col, how='outer').fillna(0)
            st.dataframe(df_sum, use_container_width=True)
        else:
            st.warning("Spalte 'Fläche' nicht gefunden; Summen nicht verfügbar.")

        # — Matching-Parameter —
        match_columns      = st.multiselect(
            "🔗 Spalten für GUID-Matching",
            options=common_cols, default=["GUID"]
        )
        overwrite_columns  = st.multiselect(
            "📝 Spalten zum Überschreiben bei Match",
            options=common_cols, default=["Fläche"]
        )
        additional_columns = st.multiselect(
            "➕ Spalten aus Update ergänzen",
            options=[c for c in df_update.columns if c not in df_original.columns]
        )
        # Fallback-Attribute & Toleranz
        fallback_columns = st.multiselect(
            "🔁 Spalten für Fallback-Matching (optional)",
            options=[c for c in common_cols if c not in match_columns],
            help="Wenn kein GUID-Match erfolgte."
        )
        tolerance = st.number_input(
            "🔒 Flächen-Toleranz für letzten Fallback (m²)",
            min_value=0.0, value=0.5, step=0.1
        )

        # — Vergleich starten —
        if st.button("Vergleich starten"):
            # Kopien für Verarbeitung
            df_orig_copy = df_original.copy()
            df_upd_copy  = df_update.copy()

            # Index-Maps
            full_map     = df_upd_copy.set_index(match_columns, drop=False)
            if fallback_columns:
                fb_map = df_upd_copy.set_index(fallback_columns, drop=False)
            else:
                fb_map = None

            # Tracking-Sets & Counters
            matched_update_keys    = set()
            matched_guid_indices   = set()
            matched_fallback_idx   = set()
            green_cells   = set()  # GUID-Matches
            blue_cells    = set()  # Fallback-Matches
            orange_cells  = set()  # Toleranz-Matches
            yellow_cells  = set()  # Neue Einträge
            updated_guid     = 0
            updated_fallback = 0
            updated_area     = 0
            unchanged        = 0

            # Progress Bar
            initial_len = df_orig_copy.shape[0]
            total_steps = initial_len + len(full_map)
            progress    = st.progress(0)
            step = 0

            # — 1) GUID- und Fallback-Matching —
            for idx in range(initial_len):
                row = df_orig_copy.loc[idx]
                key = tuple(row[c] for c in match_columns)

                # GUID-Match?
                if key in full_map.index:
                    upd_row = full_map.loc[key]
                    if isinstance(upd_row, pd.DataFrame):
                        upd_row = upd_row.iloc[0]
                    matched_update_keys.add(key)
                    matched_guid_indices.add(idx)

                    # Überschreiben
                    changed = False
                    for col in overwrite_columns:
                        new_val  = upd_row[col]
                        orig_val = row[col]
                        if pd.notna(new_val) and orig_val != new_val:
                            df_orig_copy.at[idx, f"{col} zuvor"] = orig_val
                            df_orig_copy.at[idx, col]             = new_val
                            r = idx + 2
                            c = df_orig_copy.columns.get_loc(col) + 1
                            green_cells.add((r, c))
                            changed = True
                    # Zusätzliche
                    for col in additional_columns:
                        df_orig_copy.at[idx, col] = upd_row[col]

                    if changed:
                        updated_guid += 1
                    else:
                        unchanged += 1

                # Fallback-Attribute?
                elif fb_map is not None:
                    fb_key = tuple(row[c] for c in fallback_columns)
                    if fb_key in fb_map.index:
                        fb_row = fb_map.loc[fb_key]
                        if isinstance(fb_row, pd.DataFrame) and len(fb_row) == 1:
                            fb_row = fb_row.iloc[0]
                        elif isinstance(fb_row, pd.DataFrame):
                            fb_row = None
                        if fb_row is not None:
                            # bestimme Voll-Key für Update-Tracking
                            full_key = tuple(fb_row[c] for c in match_columns)
                            if full_key in full_map.index and full_key not in matched_update_keys:
                                matched_update_keys.add(full_key)
                                matched_fallback_idx.add(idx)

                                changed_fb = False
                                for col in overwrite_columns:
                                    nv = fb_row[col]
                                    ov = row[col]
                                    if pd.notna(nv) and ov != nv:
                                        df_orig_copy.at[idx, f"{col} zuvor"] = ov
                                        df_orig_copy.at[idx, col]             = nv
                                        r = idx + 2
                                        c = df_orig_copy.columns.get_loc(col) + 1
                                        blue_cells.add((r, c))
                                        changed_fb = True
                                for col in additional_columns:
                                    df_orig_copy.at[idx, col] = fb_row[col]

                                if changed_fb:
                                    updated_fallback += 1
                                else:
                                    unchanged += 1

                step += 1
                progress.progress(min(1.0, step/total_steps))

            # — 2) Flächen-Toleranz-Matching (Orange) —
            # Original-Indizes ohne Update
            orig_unmatched = {
                idx for idx in range(initial_len)
                if idx not in matched_guid_indices and idx not in matched_fallback_idx
            }
            # Update-Keys ohne Match
            upd_unmatched = [k for k in full_map.index if k not in matched_update_keys]

            # Sammle Paare
            area_pairs = []
            for idx in orig_unmatched:
                o = df_orig_copy.loc[idx]
                for key in upd_unmatched:
                    u = full_map.loc[key]
                    if isinstance(u, pd.DataFrame):
                        u = u.iloc[0]
                    # gleiche Fallback-Attribute
                    if all(o[c] == u[c] for c in fallback_columns):
                        diff = abs(o['Fläche'] - u['Fläche'])
                        if diff <= tolerance:
                            area_pairs.append((idx, key, diff))
            # Greedy nach kleinstem Diff
            area_pairs.sort(key=lambda x: x[2])
            for idx, key, _ in area_pairs:
                if idx not in matched_guid_indices \
                and idx not in matched_fallback_idx \
                and key not in matched_update_keys:
                    u = full_map.loc[key]
                    if isinstance(u, pd.DataFrame):
                        u = u.iloc[0]
                    # Überschreiben
                    for col in overwrite_columns:
                        ov = df_orig_copy.at[idx, col]
                        df_orig_copy.at[idx, f"{col} zuvor"] = ov
                        df_orig_copy.at[idx, col]             = u[col]
                        r = idx + 2
                        c = df_orig_copy.columns.get_loc(col) + 1
                        orange_cells.add((r, c))
                    for col in additional_columns:
                        df_orig_copy.at[idx, col] = u[col]

                    matched_update_keys.add(key)
                    updated_area += 1

            # Fortschritt abschliessen
            for _ in range(len(upd_unmatched)):
                step += 1
                progress.progress(min(1.0, step/total_steps))

            # — 3) Neue GUIDs anhängen (Gelb) —
            for key in full_map.index.unique():
                if key not in matched_update_keys:
                    u = full_map.loc[key]
                    if isinstance(u, pd.DataFrame):
                        u = u.iloc[0]
                    d = u.to_dict()
                    # "... zuvor" Felder
                    for col in overwrite_columns:
                        d[f"{col} zuvor"] = None
                    df_orig_copy = pd.concat([df_orig_copy, pd.DataFrame([d])], ignore_index=True)
                    new_idx = df_orig_copy.shape[0] - 1
                    # komplette Zeile gelb markieren
                    for col_i in range(len(df_orig_copy.columns)):
                        yellow_cells.add((new_idx + 2, col_i + 1))
                    matched_update_keys.add(key)

            # — Spalten neu anordnen: "... zuvor" direkt rechts —
            orig_cols = list(df_original.columns)
            new_order = []
            for col in orig_cols:
                new_order.append(col)
                prev = f"{col} zuvor"
                if prev in df_orig_copy.columns:
                    new_order.append(prev)
            # Rest-Spalten anhängen
            for col in df_orig_copy.columns:
                if col not in new_order:
                    new_order.append(col)
            df_orig_copy = df_orig_copy[new_order]

            # — 4) Excel-Export mit Styling —
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df_orig_copy.to_excel(writer, sheet_name="Vergleich", index=False)
                wb = writer.book
                ws = writer.sheets["Vergleich"]
                ws.freeze_panes = "A2"
                max_c = df_orig_copy.shape[1]
                max_r = df_orig_copy.shape[0] + 1
                ws.auto_filter.ref = f"A1:{get_column_letter(max_c)}{max_r}"

                # Farben
                green  = PatternFill(start_color="CCFFCC", end_color="CCFFCC", fill_type="solid")
                blue   = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")
                orange = PatternFill(start_color="FFD966", end_color="FFD966", fill_type="solid")
                yellow = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

                for (r, c) in green_cells:
                    ws.cell(row=r, column=c).fill = green
                for (r, c) in blue_cells:
                    ws.cell(row=r, column=c).fill = blue
                for (r, c) in orange_cells:
                    ws.cell(row=r, column=c).fill = orange
                for (r, c) in yellow_cells:
                    ws.cell(row=r, column=c).fill = yellow

            buffer.seek(0)

            # — Zusammenfassung —
            st.markdown("### Zusammenfassung")
            st.success (f"🔁 GUID-Updates: {updated_guid}")
            st.info    (f"🔷 Fallback-Updates: {updated_fallback}")
            st.info    (f"🔶 Flächen-Toleranz-Updates: {updated_area}")
            st.info    (f"✅ Unverändert: {unchanged}")
            st.warning(f"➕ Neue Einträge ergänzt: {len([k for k in full_map.index if k not in matched_update_keys])}")

            # — Download —
            date_str   = datetime.now().strftime("%y.%m.%d")
            orig_name  = uploaded_original.name.replace(".xlsx", "")
            export_name= f"Updated_{date_str}_{orig_name}.xlsx"
            st.download_button(
                label="📥 Datei herunterladen",
                data=buffer.getvalue(),
                file_name=export_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Fehler: {e}")
