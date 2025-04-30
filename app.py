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

        # — Sidebar: Übersicht mit Metrics & Deltas —
        for df in (df_original, df_update):
            if 'Fläche' in df.columns:
                df['Fläche'] = pd.to_numeric(df['Fläche'], errors='coerce')

        n_orig     = df_original.shape[0]
        n_upd      = df_update.shape[0]
        orig_guids = df_original['GUID'].nunique() if 'GUID' in df_original.columns else 0
        upd_guids  = df_update['GUID'].nunique()   if 'GUID' in df_update.columns   else 0
        orig_total = df_original['Fläche'].sum() if 'Fläche' in df_original.columns else 0
        upd_total  = df_update['Fläche'].sum()   if 'Fläche' in df_update.columns   else 0

        st.sidebar.header("📊 Kurzübersicht")
        st.sidebar.metric("Zeilen Original", n_orig, delta=n_upd - n_orig)
        st.sidebar.metric("GUIDs Original", orig_guids, delta=upd_guids - orig_guids)
        st.sidebar.metric("Summe Fläche Original", f"{orig_total:.2f}", delta=f"{(upd_total - orig_total):.2f}")

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
            df_orig_copy = df_original.copy()
            df_upd_copy  = df_update.copy()

            full_map   = df_upd_copy.set_index(match_columns, drop=False)
            fb_map     = df_upd_copy.set_index(fallback_columns, drop=False) if fallback_columns else None

            matched_keys          = set()
            matched_guid_idx      = set()
            matched_fallback_idx  = set()

            green_cells   = set()  # GUID-Matches
            blue_cells    = set()  # Fallback-Matches
            orange_cells  = set()  # Toleranz-Matches
            yellow_cells  = set()  # Neue Einträge

            cnt_guid    = cnt_fb = cnt_area = cnt_unchanged = 0

            # Progress Bar
            initial_len = df_orig_copy.shape[0]
            total_steps = initial_len + len(full_map)
            progress    = st.progress(0)
            step = 0

            # 1) GUID- & Fallback-Matching
            for idx in range(initial_len):
                row = df_orig_copy.loc[idx]
                key = tuple(row[c] for c in match_columns)

                if key in full_map.index:
                    upd = full_map.loc[key]
                    if isinstance(upd, pd.DataFrame): upd = upd.iloc[0]
                    matched_keys.add(key); matched_guid_idx.add(idx)

                    changed=False
                    for col in overwrite_columns:
                        nv, ov = upd[col], row[col]
                        if pd.notna(nv) and ov!=nv:
                            df_orig_copy.at[idx, f"{col} zuvor"] = ov
                            df_orig_copy.at[idx, col]             = nv
                            r, c = idx+2, df_orig_copy.columns.get_loc(col)+1
                            green_cells.add((r,c)); changed=True
                    for col in additional_columns:
                        df_orig_copy.at[idx,col]=upd[col]
                    cnt_guid += 1 if changed else (cnt_unchanged:=cnt_unchanged+1)

                elif fb_map is not None:
                    fb_key = tuple(row[c] for c in fallback_columns)
                    if fb_key in fb_map.index:
                        fb = fb_map.loc[fb_key]
                        if isinstance(fb, pd.DataFrame) and len(fb)==1: fb=fb.iloc[0]
                        elif isinstance(fb, pd.DataFrame): fb=None
                        if fb is not None:
                            full_key = tuple(fb[c] for c in match_columns)
                            if full_key in full_map.index and full_key not in matched_keys:
                                matched_keys.add(full_key); matched_fallback_idx.add(idx)
                                changed_fb=False
                                for col in overwrite_columns:
                                    nv, ov = fb[col], row[col]
                                    if pd.notna(nv) and ov!=nv:
                                        df_orig_copy.at[idx, f"{col} zuvor"]=ov
                                        df_orig_copy.at[idx, col]            =nv
                                        r,c=idx+2, df_orig_copy.columns.get_loc(col)+1
                                        blue_cells.add((r,c)); changed_fb=True
                                for col in additional_columns:
                                    df_orig_copy.at[idx,col]=fb[col]
                                cnt_fb += 1 if changed_fb else (cnt_unchanged:=cnt_unchanged+1)

                step += 1
                progress.progress(min(1.0, step/total_steps))

            # 2) Flächen-Toleranz-Matching
            orig_unm = {i for i in range(initial_len) if i not in matched_guid_idx|matched_fallback_idx}
            upd_unm  = [k for k in full_map.index if k not in matched_keys]
            area_pairs=[]
            for i in orig_unm:
                o=df_orig_copy.loc[i]
                for k in upd_unm:
                    u=full_map.loc[k]; u=u.iloc[0] if isinstance(u,pd.DataFrame) else u
                    if all(o[c]==u[c] for c in fallback_columns):
                        d=abs(o['Fläche']-u['Fläche'])
                        if d<=tolerance: area_pairs.append((i,k,d))
            area_pairs.sort(key=lambda x: x[2])
            for i,k,_ in area_pairs:
                if i not in matched_guid_idx|matched_fallback_idx and k not in matched_keys:
                    u=full_map.loc[k]; u=u.iloc[0] if isinstance(u,pd.DataFrame) else u
                    for col in overwrite_columns:
                        ov=df_orig_copy.at[i,col]
                        df_orig_copy.at[i,f"{col} zuvor"]=ov
                        df_orig_copy.at[i,col]=u[col]
                        r,c=i+2, df_orig_copy.columns.get_loc(col)+1
                        orange_cells.add((r,c))
                    for col in additional_columns:
                        df_orig_copy.at[i,col]=u[col]
                    matched_keys.add(k); cnt_area+=1
            for _ in upd_unm: 
                step+=1; progress.progress(min(1.0, step/total_steps))

            # 3) Neue GUIDs anhängen
            for k in full_map.index.unique():
                if k not in matched_keys:
                    u=full_map.loc[k]; u=u.iloc[0] if isinstance(u,pd.DataFrame) else u
                    d=u.to_dict()
                    for col in overwrite_columns: d[f"{col} zuvor"]=None
                    df_orig_copy=pd.concat([df_orig_copy,pd.DataFrame([d])],ignore_index=True)
                    idx_new=df_orig_copy.shape[0]-1
                    for ci in range(len(df_orig_copy.columns)):
                        yellow_cells.add((idx_new+2,ci+1))
                    matched_keys.add(k)

            # Spalten neu anordnen („... zuvor“ neben Original)
            base_cols=list(df_original.columns)
            new_order=[]
            for c in base_cols:
                new_order.append(c)
                p=f"{c} zuvor"
                if p in df_orig_copy.columns: new_order.append(p)
            for c in df_orig_copy.columns:
                if c not in new_order: new_order.append(c)
            df_orig_copy=df_orig_copy[new_order]

            # 4) Excel-Export mit Styling
            buffer=io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df_orig_copy.to_excel(writer, sheet_name="Vergleich", index=False)
                wb=writer.book; ws=writer.sheets["Vergleich"]
                ws.freeze_panes="A2"
                max_c, max_r=df_orig_copy.shape[1], df_orig_copy.shape[0]+1
                ws.auto_filter.ref=f"A1:{get_column_letter(max_c)}{max_r}"
                green = PatternFill(start_color="CCFFCC", end_color="CCFFCC", fill_type="solid")
                blue  = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")
                orange= PatternFill(start_color="FFD966", end_color="FFD966", fill_type="solid")
                yellow= PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
                for (r,c) in green_cells:  ws.cell(r,c).fill=green
                for (r,c) in blue_cells:   ws.cell(r,c).fill=blue
                for (r,c) in orange_cells: ws.cell(r,c).fill=orange
                for (r,c) in yellow_cells: ws.cell(r,c).fill=yellow
            buffer.seek(0)

            # — Zusammenfassung —
            st.markdown("### Zusammenfassung")
            st.success    (f"🔁 GUID-Updates: {cnt_guid}")
            st.info       (f"🔷 Fallback-Updates: {cnt_fb}")
            st.info       (f"🔶 Flächen-Toleranz-Updates: {cnt_area}")
            st.info       (f"✅ Unverändert: {cnt_unchanged}")
            st.warning    (f"➕ Neue Einträge ergänzt: {len([k for k in full_map.index if k not in matched_keys])}")

            # — Download —
            date_str   = datetime.now().strftime("%y.%m.%d")
            orig_name  = uploaded_original.name.replace(".xlsx","")
            export_name= f"Updated_{date_str}_{orig_name}.xlsx"
            st.download_button(
                label="📥 Datei herunterladen",
                data=buffer.getvalue(),
                file_name=export_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Fehler: {e}")
