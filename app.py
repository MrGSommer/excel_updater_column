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
orig_file = st.file_uploader("Original Excel hochladen", type=["xlsx"])
upd_file  = st.file_uploader("Update Excel hochladen",   type=["xlsx"])

if orig_file and upd_file:
    try:
        df_orig_base = pd.read_excel(orig_file)
        df_upd_base  = pd.read_excel(upd_file)

        # — Filter —
        cols = df_orig_base.select_dtypes(exclude=["number"]).columns.tolist()
        if cols:
            fc   = st.selectbox("Filter-Spalte", options=cols)
            vals = df_orig_base[fc].dropna().unique().tolist()
            sel  = st.multiselect("Filter-Werte", options=vals, default=vals)
            df_orig = df_orig_base[df_orig_base[fc].isin(sel)].reset_index(drop=True)
            df_upd  = df_upd_base [df_upd_base [fc].isin(sel)].reset_index(drop=True)
        else:
            df_orig = df_orig_base.copy()
            df_upd  = df_upd_base.copy()

        # — Sidebar Metriken —
        for df in (df_orig, df_upd):
            if "Fläche" in df.columns:
                df["Fläche"] = pd.to_numeric(df["Fläche"], errors="coerce")
        n_o, n_u = len(df_orig), len(df_upd)
        g_o, g_u = df_orig["GUID"].nunique(), df_upd["GUID"].nunique()
        s_o = df_orig["Fläche"].sum() if "Fläche" in df_orig.columns else 0
        s_u = df_upd["Fläche"].sum()   if "Fläche" in df_upd.columns   else 0
        st.sidebar.header("📊 Kurzübersicht")
        st.sidebar.metric("Zeilen O/U", n_o, delta=n_u-n_o)
        st.sidebar.metric("GUIDs O/U", g_o, delta=g_u-g_o)
        st.sidebar.metric("Fläche O/U", f"{s_o:.2f}", delta=f"{(s_u-s_o):.2f}")

        # — Vorschau —
        st.subheader("🔍 Vorschau")
        c1, c2 = st.columns(2)
        c1.dataframe(df_orig.head(5), use_container_width=True)
        c2.dataframe(df_upd.head(5),  use_container_width=True)

        # — Spalten prüfen —
        common = sorted(set(df_orig.columns) & set(df_upd.columns))
        if "GUID" not in common:
            st.error("Spalte 'GUID' fehlt."); st.stop()

        # — Gruppensummen —
        st.subheader("📊 Gruppen-Summen")
        grp = st.selectbox("Gruppieren nach", common, index=common.index("Haus") if "Haus" in common else 0)
        if "Fläche" in common:
            sum_o = df_orig.groupby(grp)["Fläche"].sum().reset_index(name="Summe O")
            sum_u = df_upd .groupby(grp)["Fläche"].sum().reset_index(name="Summe U")
            st.dataframe(pd.merge(sum_o, sum_u, on=grp, how="outer").fillna(0), use_container_width=True)

        # — Parameter —
        match_cols      = st.multiselect("GUID-Matching", common, default=["GUID"])
        overwrite_cols  = st.multiselect("Überschreiben", common, default=["Fläche"])
        add_cols        = st.multiselect("Ergänzen", [c for c in df_upd.columns if c not in df_orig.columns])
        fb_cols         = st.multiselect("Fallback-Spalten", [c for c in common if c not in match_cols])
        tol             = st.number_input("Toleranz (m²)", 0.0, 5.0, 0.5, 0.1)

        # — Existenzprüfung —
        exist_cols = st.multiselect("Existenz-Attribute", [c for c in common if c!="GUID"])
        adopt      = []
        if exist_cols:
            orig_set = {tuple(r[c] for c in exist_cols) for _,r in df_orig.iterrows()}
            upd_sub  = df_upd.drop_duplicates(exist_cols)
            miss     = upd_sub[~upd_sub.apply(lambda r: tuple(r[c] for c in exist_cols) in orig_set, axis=1)]
            if not miss.empty:
                st.subheader("⚠️ Fehlende Kombinationen")
                for i,r in miss.iterrows():
                    lbl = " | ".join(f"{c}:{r[c]}" for c in exist_cols)
                    if st.checkbox(f"Übernehmen {lbl}", key=f"cb{i}"):
                        adopt.append(tuple(r[c] for c in exist_cols))

        # — Vergleich —
        if st.button("Vergleich starten"):
            df = df_orig.copy()
            full = df_upd.set_index(match_cols, drop=False)
            fb   = df_upd.set_index(fb_cols,   drop=False) if fb_cols else None

            matched, idx_guid, idx_fb = set(), set(), set()
            green, blue, orange, yellow, red = set(), set(), set(), set(), set()
            c_guid = c_fb = c_area = c_unc = c_new = c_susp = 0
            total = len(df)+len(full)
            pbar  = st.progress(0); step=0

            # 1) GUID & Fallback
            for i in range(len(df)):
                row = df.loc[i]; key=tuple(row[c] for c in match_cols)
                upd = full.loc[key] if key in full.index else None
                used = False
                if isinstance(upd, pd.DataFrame): upd=upd.iloc[0]
                if upd is not None:
                    matched.add(key); idx_guid.add(i)
                    for c in overwrite_cols:
                        nv,ov=upd[c],row[c]
                        if pd.notna(nv) and nv!=ov:
                            df.at[i,f"{c} zuvor"]=ov; df.at[i,c]=nv
                            green.add((i+2, df.columns.get_loc(c)+1)); used=True
                    for c in add_cols: df.at[i,c]=upd[c]
                    c_guid+=1 if used else (c_unc:=c_unc+1)
                elif fb and fb_cols:
                    fk=tuple(row[c] for c in fb_cols)
                    fbr=fb.loc[fk] if fk in fb.index else None
                    if isinstance(fbr,pd.DataFrame) and len(fbr)==1: fbr=fbr.iloc[0]
                    if isinstance(fbr,pd.Series):
                        fk2=tuple(fbr[c] for c in match_cols)
                        if fk2 in full.index and fk2 not in matched:
                            matched.add(fk2); idx_fb.add(i)
                            used=False
                            for c in overwrite_cols:
                                nv,ov=fbr[c],row[c]
                                if pd.notna(nv) and nv!=ov:
                                    df.at[i,f"{c} zuvor"]=ov; df.at[i,c]=nv
                                    blue.add((i+2, df.columns.get_loc(c)+1)); used=True
                            for c in add_cols: df.at[i,c]=fbr[c]
                            c_fb+=1 if used else (c_unc:=c_unc+1)
                step+=1; pbar.progress(min(1,step/total))

            # 2) Toleranz-Matching
            orig_un = {i for i in range(len(df)) if i not in idx_guid|idx_fb}
            upd_un  = [k for k in full.index if k not in matched]
            pairs=[]
            for i in orig_un:
                o=df.loc[i]
                for k in upd_un:
                    u=full.loc[k]
                    u=u.iloc[0] if isinstance(u,pd.DataFrame) else u
                    if fb_cols and all(o[c]==u[c] for c in fb_cols):
                        d=abs(o["Fläche"]-u["Fläche"])
                        if d<=tol: pairs.append((i,k,d))
                        else:   # zu große Abweichung
                            red.add((i+2, df.columns.get_loc("Fläche")+1))
            pairs.sort(key=lambda x:x[2])
            for i,k,_ in pairs:
                if i not in idx_guid|idx_fb and k not in matched:
                    u=full.loc[k]; u=u.iloc[0] if isinstance(u,pd.DataFrame) else u
                    for c in overwrite_cols:
                        ov=df.at[i,c]; df.at[i,f"{c} zuvor"]=ov; df.at[i,c]=u[c]
                        orange.add((i+2, df.columns.get_loc(c)+1))
                    for c in add_cols: df.at[i,c]=u[c]
                    matched.add(k); c_area+=1
            step+=len(upd_un); pbar.progress(1.0)

            # 3) Neueinträge & Existenz
            for k in full.index.unique():
                if k not in matched:
                    u=full.loc[k]
                    u=u.iloc[0] if isinstance(u,pd.DataFrame) else u
                    if exist_cols:
                        ek=tuple(u[c] for c in exist_cols)
                        if ek not in adopt: continue
                    d=u.to_dict()
                    for c in overwrite_cols: d[f"{c} zuvor"]=None
                    df=pd.concat([df,pd.DataFrame([d])],ignore_index=True)
                    ni=len(df)-1
                    for ci in range(len(df.columns)):
                        yellow.add((ni+2,ci+1))
                    matched.add(k); c_new+=1

            # Spalten ordnen
            base=list(df_orig_base.columns)
            order=[]
            for c in base:
                order.append(c)
                p=f"{c} zuvor"
                if p in df.columns: order.append(p)
            for c in df.columns:
                if c not in order: order.append(c)
            df=df[order]

            # 4) Export
            buf=io.BytesIO()
            with pd.ExcelWriter(buf, engine="openpyxl") as w:
                df.to_excel(w, sheet_name="Vergleich", index=False)
                wb,ws=w.book,w.sheets["Vergleich"]
                ws.freeze_panes="A2"
                mc,mr=len(df.columns),len(df)+1
                ws.auto_filter.ref=f"A1:{get_column_letter(mc)}{mr}"
                fmt = {
                    "g":PatternFill(start_color="CCFFCC",fill_type="solid"),
                    "b":PatternFill(start_color="ADD8E6",fill_type="solid"),
                    "o":PatternFill(start_color="FFD966",fill_type="solid"),
                    "y":PatternFill(start_color="FFFF00",fill_type="solid"),
                    "r":PatternFill(start_color="FFC7CE",fill_type="solid")
                }
                for r,c in green:  ws.cell(r,c).fill=fmt["g"]
                for r,c in blue:   ws.cell(r,c).fill=fmt["b"]
                for r,c in orange: ws.cell(r,c).fill=fmt["o"]
                for r,c in yellow: ws.cell(r,c).fill=fmt["y"]
                for r,c in red:    ws.cell(r,c).fill=fmt["r"]
            buf.seek(0)

            # Zusammenfassung
            st.markdown("### Zusammenfassung")
            st.success(f"🔁 GUID: {c_guid}")
            st.info   (f"🔷 Fallback: {c_fb}")
            st.info   (f"🔶 Toleranz: {c_area}")
            st.warning(f"⚠️ Suspicious (Toleranz überschritten): {len(red)}")
            st.info   (f"➕ Neu: {c_new}")
            st.info   (f"✅ Unverändert: {c_unc}")

            # Download
            ds=datetime.now().strftime("%y.%m.%d")
            name=orig_file.name.replace(".xlsx","")
            st.download_button("📥 Herunterladen", buf.getvalue(),
                               file_name=f"Updated_{ds}_{name}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        st.error(f"Fehler: {e}")
