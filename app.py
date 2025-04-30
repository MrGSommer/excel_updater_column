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
        # 📑 Basis-Daten einlesen
        df_original_base = pd.read_excel(uploaded_original)
        df_update_base   = pd.read_excel(uploaded_update)

        # — Filter: Ausmasscode etc. —
        candidate_cols = df_original_base.select_dtypes(exclude=["number"]).columns.tolist()
        if candidate_cols:
            filter_col = st.selectbox("🗂️ Filter-Spalte auswählen", options=candidate_cols)
            values     = df_original_base[filter_col].dropna().unique().tolist()
            selected   = st.multiselect(f"🔎 {filter_col}-Werte filtern", options=values, default=values)
            df_original = df_original_base[df_original_base[filter_col].isin(selected)].reset_index(drop=True)
            df_update   = df_update_base  [df_update_base  [filter_col].isin(selected)].reset_index(drop=True)
        else:
            df_original = df_original_base.copy()
            df_update   = df_update_base.copy()

        # — Sidebar: Kennzahlen & Deltas —
        for df in (df_original, df_update):
            if "Fläche" in df.columns:
                df["Fläche"] = pd.to_numeric(df["Fläche"], errors="coerce")

        n_orig     = df_original.shape[0]
        n_upd      = df_update.shape[0]
        orig_guids = df_original["GUID"].nunique()
        upd_guids  = df_update["GUID"].nunique()
        orig_sum   = df_original["Fläche"].sum() if "Fläche" in df_original.columns else 0
        upd_sum    = df_update["Fläche"].sum()   if "Fläche" in df_update.columns   else 0

        st.sidebar.header("📊 Kurzübersicht")
        st.sidebar.metric("Zeilen (O/U)", f"{n_orig}", delta=n_upd-n_orig)
        st.sidebar.metric("GUIDs (O/U)", f"{orig_guids}", delta=upd_guids-orig_guids)
        st.sidebar.metric("Summe Fläche (O/U)", f"{orig_sum:.2f}", delta=f"{(upd_sum-orig_sum):.2f}")

        # — Vorschau —
        st.subheader("🔍 Vorschau")
        col1, col2 = st.columns(2)
        with col1:
            st.write("Original (5 Zeilen)")
            st.dataframe(df_original.head(5), use_container_width=True)
        with col2:
            st.write("Update (5 Zeilen)")
            st.dataframe(df_update.head(5),   use_container_width=True)

        # — Gemeinsame Spalten prüfen —
        common_cols = sorted(set(df_original.columns) & set(df_update.columns))
        if "GUID" not in common_cols:
            st.error("Beide Dateien müssen eine 'GUID'-Spalte haben.")
            st.stop()

        # — Flächensummen nach Gruppe —
        st.subheader("📊 Gruppen-Summen")
        group_col = st.selectbox("Gruppieren nach", options=common_cols, index=common_cols.index("Haus") if "Haus" in common_cols else 0)
        if "Fläche" in common_cols:
            sum_o = df_original.groupby(group_col)["Fläche"].sum().reset_index(name="Summe O")
            sum_u = df_update  .groupby(group_col)["Fläche"].sum().reset_index(name="Summe U")
            st.dataframe(pd.merge(sum_o, sum_u, on=group_col, how="outer").fillna(0), use_container_width=True)

        # — Matching-Parameter —
        match_columns      = st.multiselect("🔗 Matching-Spalten",       options=common_cols, default=["GUID"])
        overwrite_columns  = st.multiselect("📝 Überschreiben-Spalten", options=common_cols, default=["Fläche"])
        additional_columns = st.multiselect("➕ Ergänzen-Spalten",       options=[c for c in df_update.columns if c not in df_original.columns])

        # — Fallback- & Toleranz-Stufe —
        fallback_columns = st.multiselect("🔁 Fallback-Matching-Spalten", options=[c for c in common_cols if c not in match_columns])
        tolerance        = st.number_input("🔒 Flächen-Toleranz (m²)", min_value=0.0, value=0.5, step=0.1)

        # — Existenzprüfung für neue Räume —
        existence_attrs = st.multiselect(
            "🔍 Attribute für Existenzprüfung",
            options=[c for c in common_cols if c != "GUID"],
            help="Kombination muss im Original vorhanden sein, sonst interaktive Entscheidung."
        )
        adopt_list = []
        if existence_attrs:
            # fehlende Kombinationen erkennen
            orig_set = {
                tuple(r[a] for a in existence_attrs)
                for _, r in df_original_base.iterrows()
            }
            upd_sub = df_update_base.drop_duplicates(existence_attrs)
            missing_combos = upd_sub[~upd_sub.apply(lambda r: tuple(r[a] for a in existence_attrs) in orig_set, axis=1)]
            missing_combos = missing_combos[existence_attrs].drop_duplicates().reset_index(drop=True)
            st.subheader("⚠️ Fehlende Kombinationen")
            for i, combo in missing_combos.iterrows():
                label = " — ".join(f"{a}: {combo[a]}" for a in existence_attrs)
                if st.checkbox(f"Übernehmen: {label}", key=f"adopt_{i}"):
                    adopt_list.append(tuple(combo[a] for a in existence_attrs))

        # — Vergleich starten —
        if st.button("Vergleich starten"):
            df_orig = df_original.copy()
            full_map = df_update.set_index(match_columns, drop=False)
            fb_map   = df_update.set_index(fallback_columns, drop=False) if fallback_columns else None

            matched     = set()
            gui_idx     = set()
            fb_idx      = set()
            green, blue, orange, yellow = set(), set(), set(), set()
            cnt_guid=cnt_fb=cnt_area=cnt_unc=0

            # Progress
            total = len(df_orig) + len(full_map)
            prog  = st.progress(0); step=0

            # 1) GUID & Fallback
            for idx in range(len(df_orig)):
                row = df_orig.loc[idx]
                key = tuple(row[c] for c in match_columns)
                updated=False

                # GUID
                if key in full_map.index:
                    upd = full_map.loc[key]
                    if isinstance(upd, pd.DataFrame): upd=upd.iloc[0]
                    matched.add(key); gui_idx.add(idx)
                    for col in overwrite_columns:
                        nv, ov = upd[col], row[col]
                        if pd.notna(nv) and nv!=ov:
                            df_orig.at[idx, f"{col} zuvor"]=ov
                            df_orig.at[idx, col]            =nv
                            r,c=idx+2,df_orig.columns.get_loc(col)+1
                            green.add((r,c)); updated=True
                    for col in additional_columns:
                        df_orig.at[idx,col]=upd[col]
                    cnt_guid += 1 if updated else (cnt_unc:=cnt_unc+1)

                # Fallback
                elif fb_map is not None:
                    fb_key = tuple(row[c] for c in fallback_columns)
                    if fb_key in fb_map.index:
                        fb = fb_map.loc[fb_key]
                        if isinstance(fb,pd.DataFrame) and len(fb)==1: fb=fb.iloc[0]
                        elif isinstance(fb,pd.DataFrame): fb=None
                        if fb is not None:
                            fullk = tuple(fb[c] for c in match_columns)
                            if fullk in full_map.index and fullk not in matched:
                                matched.add(fullk); fb_idx.add(idx)
                                updated_fb=False
                                for col in overwrite_columns:
                                    nv, ov = fb[col], row[col]
                                    if pd.notna(nv) and nv!=ov:
                                        df_orig.at[idx, f"{col} zuvor"]=ov
                                        df_orig.at[idx, col]            =nv
                                        r,c=idx+2,df_orig.columns.get_loc(col)+1
                                        blue.add((r,c)); updated_fb=True
                                for col in additional_columns:
                                    df_orig.at[idx,col]=fb[col]
                                cnt_fb += 1 if updated_fb else (cnt_unc:=cnt_unc+1)

                step+=1; prog.progress(min(1.0, step/total))

            # 2) Toleranz-Matching
            orig_unm = {i for i in range(len(df_orig)) if i not in gui_idx|fb_idx}
            upd_unm  = [k for k in full_map.index if k not in matched]
            pairs=[]
            for i in orig_unm:
                o=df_orig.loc[i]
                for k in upd_unm:
                    u=full_map.loc[k]
                    u=u.iloc[0] if isinstance(u,pd.DataFrame) else u
                    if all(o[c]==u[c] for c in fallback_columns):
                        d=abs(o["Fläche"]-u["Fläche"])
                        if d<=tolerance: pairs.append((i,k,d))
            pairs.sort(key=lambda x: x[2])
            for i,k,_ in pairs:
                if i not in gui_idx|fb_idx and k not in matched:
                    u=full_map.loc[k]; u=u.iloc[0] if isinstance(u,pd.DataFrame) else u
                    for col in overwrite_columns:
                        ov=df_orig.at[i,col]
                        df_orig.at[i,f"{col} zuvor"]=ov
                        df_orig.at[i,col]=u[col]
                        r,c=i+2,df_orig.columns.get_loc(col)+1
                        orange.add((r,c))
                    for col in additional_columns:
                        df_orig.at[i,col]=u[col]
                    matched.add(k); cnt_area+=1
            for _ in upd_unm:
                step+=1; prog.progress(min(1.0, step/total))

            # 3) Neueinträge mit Existenzprüfung
            new_keys=[] 
            for k in full_map.index.unique():
                if k not in matched:
                    u=full_map.loc[k]
                    u=u.iloc[0] if isinstance(u,pd.DataFrame) else u
                    # Existenz prüfen
                    if existence_attrs:
                        ek=tuple(u[a] for a in existence_attrs)
                        if ek not in adopt_list:
                            continue
                    d=u.to_dict()
                    for col in overwrite_columns:
                        d[f"{col} zuvor"]=None
                    df_orig=pd.concat([df_orig,pd.DataFrame([d])],ignore_index=True)
                    ni=len(df_orig)-1
                    for ci in range(len(df_orig.columns)):
                        yellow.add((ni+2,ci+1))
                    matched.add(k); new_keys.append(k)

            # Spalten neu anordnen
            base_cols=list(df_original_base.columns)
            order=[]
            for c in base_cols:
                order.append(c)
                p=f"{c} zuvor"
                if p in df_orig.columns: order.append(p)
            for c in df_orig.columns:
                if c not in order: order.append(c)
            df_orig=df_orig[order]

            # 4) Export mit Styling
            buf=io.BytesIO()
            with pd.ExcelWriter(buf, engine="openpyxl") as writer:
                df_orig.to_excel(writer, sheet_name="Vergleich", index=False)
                wb=writer.book; ws=writer.sheets["Vergleich"]
                ws.freeze_panes="A2"
                mc,mr=len(df_orig.columns),len(df_orig)+1
                ws.auto_filter.ref=f"A1:{get_column_letter(mc)}{mr}"
                fmt= {
                    "green": PatternFill(start_color="CCFFCC",end_color="CCFFCC",fill_type="solid"),
                    "blue":  PatternFill(start_color="ADD8E6",end_color="ADD8E6",fill_type="solid"),
                    "orange":PatternFill(start_color="FFD966",end_color="FFD966",fill_type="solid"),
                    "yellow":PatternFill(start_color="FFFF00",end_color="FFFF00",fill_type="solid"),
                }
                for (r,c) in green:  ws.cell(r,c).fill=fmt["green"]
                for (r,c) in blue:   ws.cell(r,c).fill=fmt["blue"]
                for (r,c) in orange: ws.cell(r,c).fill=fmt["orange"]
                for (r,c) in yellow: ws.cell(r,c).fill=fmt["yellow"]
            buf.seek(0)

            # — Zusammenfassung —
            st.markdown("### Zusammenfassung")
            st.success (f"🔁 GUID-Updates: {cnt_guid}")
            st.info    (f"🔷 Fallback-Updates: {cnt_fb}")
            st.info    (f"🔶 Toleranz-Updates: {cnt_area}")
            st.info    (f"✅ Unverändert: {cnt_unc}")
            st.warning(f"➕ Neu übernommen: {len(new_keys)}")

            # — Download —
            ds = datetime.now().strftime("%y.%m.%d")
            name=uploaded_original.name.replace(".xlsx","")
            st.download_button("📥 Datei herunterladen", buf.getvalue(),
                               file_name=f"Updated_{ds}_{name}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        st.error(f"Fehler: {e}")
