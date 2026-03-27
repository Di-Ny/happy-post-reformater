import streamlit as st
import fitz  # PyMuPDF
import io
import os
import tempfile
import pandas as pd
from datetime import date as date_cls
try:
    from generate_import import (
        parse_orders_from_pdf_text, parse_orders_from_tsv_bytes,
        generate_import_file, generate_gls_csv, detect_product_info, COUNTRY_MAP,
    )
except Exception as e:
    import streamlit as _st
    _st.error(f"Erreur import generate_import: {type(e).__name__}: {e}")
    raise

st.set_page_config(
    page_title="Happy Post - Outils d'expedition",
    page_icon="📦",
    layout="wide",
)

st.title("📦 Happy Post — Outils d'expédition")


# -- Helpers de layout --
def get_layout_config(labels_per_page):
    """Retourne la config de grille pour 4 ou 6 etiquettes par page."""
    A4_W, A4_H = 595.28, 841.89
    if labels_per_page == 6:
        # Avery L7166 : 99.1mm x 93.1mm, grille 2x3
        cell_w_mm, cell_h_mm = 99.1, 93.1
        margin_x = (210 - 2 * cell_w_mm) / 2  # ~5.9mm
        margin_y = (297 - 3 * cell_h_mm) / 2  # ~8.85mm
        # Conversion mm -> points (1mm = 2.8346pt)
        cell_w = cell_w_mm * 2.8346
        cell_h = cell_h_mm * 2.8346
        mx = margin_x * 2.8346
        my = margin_y * 2.8346
        cols, rows = 2, 3
        positions = []
        for r in range(rows):
            for c in range(cols):
                positions.append((mx + c * cell_w, my + r * cell_h))
    else:
        margin = 12
        gap = 6
        usable_w = A4_W - 2 * margin
        usable_h = A4_H - 2 * margin
        cell_w = (usable_w - gap) / 2
        cell_h = (usable_h - gap) / 2
        mx, my = margin, margin
        cols, rows = 2, 2
        positions = [
            (mx, my),
            (mx + cell_w + gap, my),
            (mx, my + cell_h + gap),
            (mx + cell_w + gap, my + cell_h + gap),
        ]
    return {
        "A4_W": A4_W, "A4_H": A4_H,
        "cell_w": cell_w, "cell_h": cell_h,
        "positions": positions,
        "cols": cols, "rows": rows,
        "margin_x": mx, "margin_y": my,
        "labels_per_page": labels_per_page,
    }


def best_orientation(src_w, src_h, cell_w, cell_h):
    """Retourne (ratio, rotate) pour l'orientation qui remplit le mieux la cellule."""
    # Sans rotation (portrait)
    ratio_p = src_w / src_h
    cell_ratio = cell_w / cell_h
    if ratio_p > cell_ratio:
        fill_p = cell_w * (cell_w / ratio_p)
    else:
        fill_p = cell_h * (cell_h * ratio_p)
    # Avec rotation 90deg : w et h echanges
    ratio_r = src_h / src_w
    if ratio_r > cell_ratio:
        fill_r = cell_w * (cell_w / ratio_r)
    else:
        fill_r = cell_h * (cell_h * ratio_r)
    if fill_r >= fill_p:
        return ratio_r, 90
    return ratio_p, 0


def smart_crop(src_page):
    """Detecte la zone utile de l'etiquette Happy Post en trouvant le cadre principal."""
    src_rect = src_page.rect
    # Chercher le plus grand rectangle dessine (= cadre principal de l'etiquette)
    best_rect = None
    best_area = 0
    for p in src_page.get_drawings():
        r = p["rect"]
        area = r.width * r.height
        if area > best_area and r.width > 200 and r.height > 150:
            best_rect = r
            best_area = area

    if best_rect:
        # Etendre: a gauche pour le logo, en bas pour les codes-barres sous le cadre
        margin_pt = 8
        x0 = max(best_rect.x0 - 50, 0)            # logo HAPPY-POST a gauche du cadre
        y0 = max(best_rect.y0 - margin_pt, 0)      # petite marge en haut
        x1 = min(best_rect.x1 + margin_pt, src_rect.width)
        y1 = min(best_rect.y1 + 25, src_rect.height)  # codes-barres + tracking en dessous
        return fitz.Rect(x0, y0, x1, y1)

    # Fallback: crop ratio fixe si pas de rectangle detecte
    return fitz.Rect(
        src_rect.x0, src_rect.y0,
        src_rect.x1, src_rect.y0 + src_rect.height * 0.47,
    )


def render_label_in_cell(page, src_page, cfg, idx):
    """Rend une etiquette dans la cellule idx de la page."""
    clip = smart_crop(src_page)
    cell_w, cell_h = cfg["cell_w"], cfg["cell_h"]
    this_src_w = clip.width
    this_src_h = clip.height
    this_rot_ratio, rotate = best_orientation(this_src_w, this_src_h, cell_w, cell_h)

    zoom = 3.0
    mat = fitz.Matrix(zoom, zoom).prerotate(rotate)
    pix = src_page.get_pixmap(matrix=mat, clip=clip)

    cell_ratio = cell_w / cell_h
    if this_rot_ratio > cell_ratio:
        img_w = cell_w
        img_h = cell_w / this_rot_ratio
    else:
        img_h = cell_h
        img_w = cell_h * this_rot_ratio

    cx, cy = cfg["positions"][idx]
    offset_x = (cell_w - img_w) / 2
    offset_y = (cell_h - img_h) / 2
    target = fitz.Rect(cx + offset_x, cy + offset_y,
                       cx + offset_x + img_w, cy + offset_y + img_h)
    page.insert_image(target, pixmap=pix)


def draw_cut_guides(page, cfg):
    """Dessine les lignes de decoupe et reperes de coins."""
    A4_W, A4_H = cfg["A4_W"], cfg["A4_H"]
    cell_w, cell_h = cfg["cell_w"], cfg["cell_h"]
    mx, my = cfg["margin_x"], cfg["margin_y"]
    cols, rows = cfg["cols"], cfg["rows"]
    gray = (0.65, 0.65, 0.65)
    dash = "[3 3] 0"
    shape = page.new_shape()

    # Lignes horizontales entre les rangees
    for r in range(1, rows):
        y = my + r * cell_h
        shape.draw_line(fitz.Point(mx, y), fitz.Point(A4_W - mx, y))
        shape.finish(color=gray, width=0.4, dashes=dash)

    # Lignes verticales entre les colonnes
    for c in range(1, cols):
        x = mx + c * cell_w
        shape.draw_line(fitz.Point(x, my), fitz.Point(x, A4_H - my))
        shape.finish(color=gray, width=0.4, dashes=dash)

    # Reperes de coins
    mark = 6
    for corner_x in [mx, mx + cols * cell_w]:
        for corner_y in [my, my + rows * cell_h]:
            dx = mark if corner_x == mx else -mark
            dy = mark if corner_y == my else -mark
            shape.draw_line(fitz.Point(corner_x, corner_y),
                            fitz.Point(corner_x + dx, corner_y))
            shape.finish(color=gray, width=0.3)
            shape.draw_line(fitz.Point(corner_x, corner_y),
                            fitz.Point(corner_x, corner_y + dy))
            shape.finish(color=gray, width=0.3)
    shape.commit()



tab_etiquettes, tab_multi_etiquettes, tab_import = st.tabs([
    "✂️ Reformater les étiquettes",
    "✂️ Reformater (multi-fichiers)",
    "📋 Générer le fichier d'import",
])

# =============================================================================
# ONGLET 1 : Reformatage des étiquettes
# =============================================================================
with tab_etiquettes:
    _, col1, _ = st.columns([1, 2, 1])
    with col1:
        st.markdown(
            "Transformez vos étiquettes Happy Post : **plusieurs par feuille A4** au lieu d'une seule. "
            "Économisez du papier."
        )

        st.image("preview.png", use_container_width=True)

        format_choice_1 = st.radio(
            "Format de sortie",
            ["4 par page (2x2)", "6 par page — Avery L7166 (2x3)"],
            horizontal=True,
            key="format_tab1",
        )
        lpp_1 = 6 if "6" in format_choice_1 else 4

        uploaded_labels = st.file_uploader(
            "Glissez ici votre PDF d'étiquettes Happy Post",
            type=["pdf"],
            key="labels_uploader",
        )

        if uploaded_labels:
            pdf_bytes = uploaded_labels.read()
            src = fitz.open(stream=pdf_bytes, filetype="pdf")
            n_pages = len(src)

            st.info(f"📄 {n_pages} page(s) détectée(s) dans **{uploaded_labels.name}**")

            cfg = get_layout_config(lpp_1)
            all_pages = list(range(n_pages))
            dst = fitz.open()

            for batch_start in range(0, len(all_pages), lpp_1):
                batch = all_pages[batch_start:batch_start + lpp_1]
                page = dst.new_page(width=cfg["A4_W"], height=cfg["A4_H"])

                for idx, src_page_num in enumerate(batch):
                    render_label_in_cell(page, src[src_page_num], cfg, idx)

                draw_cut_guides(page, cfg)

            out_buf = io.BytesIO()
            dst.save(out_buf)
            dst.close()
            src.close()
            out_buf.seek(0)

            n_sheets = -(-n_pages // lpp_1)
            out_name = uploaded_labels.name.replace(".pdf", f"_{lpp_1}par_page.pdf")

            st.success(
                f"✅ {n_pages} étiquettes → **{n_sheets} feuille(s) A4**  \n"
                f"Économie : **{n_pages - n_sheets} feuille(s)** en moins !"
            )

            st.download_button(
                label=f"⬇️ Télécharger {out_name}",
                data=out_buf,
                file_name=out_name,
                mime="application/pdf",
                type="primary",
            )

# =============================================================================
# ONGLET 2 : Reformatage multi-fichiers
# =============================================================================
with tab_multi_etiquettes:
    _, col2, _ = st.columns([1, 2, 1])
    with col2:
        st.markdown(
            "Vous avez **plusieurs fichiers PDF** (1 étiquette par fichier) ? "
            "Combinez-les en un seul PDF avec **plusieurs étiquettes par feuille A4**."
        )

        format_choice_2 = st.radio(
            "Format de sortie",
            ["4 par page (2x2)", "6 par page — Avery L7166 (2x3)"],
            horizontal=True,
            key="format_tab2",
        )
        lpp_2 = 6 if "6" in format_choice_2 else 4

        uploaded_multi = st.file_uploader(
            "Glissez ici vos fichiers PDF d'étiquettes",
            type=["pdf"],
            accept_multiple_files=True,
            key="multi_labels_uploader",
        )

        if uploaded_multi:
            n_files = len(uploaded_multi)
            st.info(f"📄 **{n_files} fichier(s)** sélectionné(s)")

            sources = []
            for uf in uploaded_multi:
                pdf_bytes = uf.read()
                sources.append(fitz.open(stream=pdf_bytes, filetype="pdf"))

            cfg = get_layout_config(lpp_2)
            dst = fitz.open()

            for batch_start in range(0, n_files, lpp_2):
                batch = sources[batch_start:batch_start + lpp_2]
                page = dst.new_page(width=cfg["A4_W"], height=cfg["A4_H"])

                for idx, src_doc in enumerate(batch):
                    render_label_in_cell(page, src_doc[0], cfg, idx)

                draw_cut_guides(page, cfg)

            for s in sources:
                s.close()

            out_buf = io.BytesIO()
            dst.save(out_buf)
            dst.close()
            out_buf.seek(0)

            n_sheets = -(-n_files // lpp_2)
            out_name = f"etiquettes_{lpp_2}par_page.pdf"
            st.success(
                f"✅ {n_files} étiquettes → **{n_sheets} feuille(s) A4**  \n"
                f"Économie : **{n_files - n_sheets} feuille(s)** en moins !"
            )

            st.download_button(
                label=f"⬇️ Télécharger {out_name}",
                data=out_buf,
                file_name=out_name,
                mime="application/pdf",
                type="primary",
                key="multi_download",
            )

# =============================================================================
# ONGLET 3 : Génération du fichier d'import
# =============================================================================
with tab_import:
    # --- En-tête centré ---
    _, col_top, _ = st.columns([1, 2, 1])
    with col_top:
        st.markdown(
            "Générez le fichier d'import Happy Post (.xlsx) à partir de vos commandes Amazon. "
            "Supporte les **PDF** (bons de commande) et les **rapports TSV** (Unshipped Orders)."
        )

        transporteur = st.radio(
            "Transporteur",
            ["📦 Happy Post (xlsx)", "🚛 GLS (csv)"],
            horizontal=True,
            key="transporteur_choice",
        )
        is_gls = "GLS" in transporteur

        uploaded_amazon = st.file_uploader(
            "Glissez ici votre fichier Amazon (PDF ou rapport TXT/TSV)",
            type=["pdf", "txt", "tsv"],
            key="amazon_uploader",
        )

    if uploaded_amazon:
        file_bytes = uploaded_amazon.read()
        file_ext = uploaded_amazon.name.rsplit(".", 1)[-1].lower()

        # Parser selon le format
        if file_ext == "pdf":
            with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
                tmp.write(file_bytes)
                tmp_path = tmp.name
            try:
                orders = parse_orders_from_pdf_text(tmp_path)
            finally:
                os.unlink(tmp_path)
            is_tsv = False
        else:
            orders = parse_orders_from_tsv_bytes(file_bytes)
            is_tsv = True

        if not orders:
            _, col_warn, _ = st.columns([1, 2, 1])
            with col_warn:
                st.warning("⚠️ Aucune commande trouvée dans ce fichier.")
        else:
            # --- Indicateurs d'urgence (centrés) ---
            _, col_metrics, _ = st.columns([1, 2, 1])
            with col_metrics:
                if is_tsv:
                    today = date_cls.today()
                    nb_late = sum(1 for o in orders if (o.get("days_past_promise") or 0) > 0)
                    nb_today = sum(1 for o in orders if o.get("promise_date") == today)
                    nb_tomorrow = sum(1 for o in orders
                                      if o.get("promise_date") and o["promise_date"] > today
                                      and (o["promise_date"] - today).days == 1)

                    m1, m2, m3, m4 = st.columns(4)
                    m1.metric("Total", len(orders))
                    m2.metric("🔴 En retard", nb_late)
                    m3.metric("🟠 Aujourd'hui", nb_today)
                    m4.metric("🟢 Demain", nb_tomorrow)
                else:
                    st.success(f"✅ **{len(orders)} commande(s)** extraites")

            # --- Tri par urgence AVANT construction du DataFrame ---
            if is_tsv:
                orders = sorted(
                    orders,
                    key=lambda o: o.get("days_past_promise") if o.get("days_past_promise") is not None else -999,
                    reverse=True,
                )

            # --- Construction du DataFrame (depuis orders déjà triées) ---
            weight_map = {"x1": 0.31, "x2": 0.32, "x3": 0.35, "Multi": 0.34, "Multi x3": 0.50}

            edit_data = []
            for o in orders:
                pays_code = o.get("pays_code", "BE")
                type_label = o.get("type_label", "")
                if not type_label:
                    label_map = {0.31: "x1", 0.32: "x2", 0.35: "x3", 0.34: "Multi", 0.50: "Multi x3"}
                    type_label = label_map.get(o["poids"], "autre")

                qty = o.get("quantite", 1)
                ppu = o.get("pieces_par_unite", 0)
                total_p = ppu * qty if ppu > 0 else 0

                urgence = ""
                if is_tsv:
                    dp = o.get("days_past_promise")
                    if dp is not None:
                        if dp > 0:
                            urgence = f"🔴 {dp}j retard"
                        elif dp == 0:
                            urgence = "🟠 Aujourd'hui"
                        elif dp == -1:
                            urgence = "🟠 Demain"
                        else:
                            urgence = f"🟢 J{dp}"

                edit_data.append({
                    "Exporter": pays_code != "FR",
                    "Urgence": urgence,
                    "Pays": pays_code,
                    "Nom": o["nom"],
                    "Prénom": o["prenom"],
                    "Adresse": o["adresse"],
                    "Complément": o["complement"],
                    "CP": o["code_postal"],
                    "Ville": o["ville"],
                    "Téléphone": o["telephone"],
                    "Email": o.get("email", ""),
                    "Type": type_label,
                    "Qté": qty,
                    "Pièces": total_p if total_p > 0 else None,
                    "Poids (kg)": o["poids"],
                    "Commande": o["commande"],
                })

            df = pd.DataFrame(edit_data)

            # Warnings centrés
            missing_rows = []
            warn_rows = []
            default_email = "contact@furgo.fr"
            for i, row in df.iterrows():
                missing = []
                if not row["Adresse"]:
                    missing.append("Adresse")
                if not row["Ville"]:
                    missing.append("Ville")
                if not row["CP"]:
                    missing.append("CP")
                if not row["Téléphone"]:
                    missing.append("Téléphone")
                if not row["Email"] or row["Email"] == default_email:
                    missing.append("Email")
                if missing:
                    label = f"**Ligne {i+1}** ({row['Nom']} {row['Prénom']}) : {', '.join(missing)}"
                    # Adresse/Ville/CP manquants = erreur, le reste = avertissement
                    has_critical = any(f in missing for f in ["Adresse", "Ville", "CP"])
                    if has_critical:
                        missing_rows.append(label)
                    else:
                        warn_rows.append(label)

            _, col_miss, _ = st.columns([1, 2, 1])
            with col_miss:
                if missing_rows:
                    st.error("🔴 **Champs critiques manquants :**\n" + "\n".join(f"- {r}" for r in missing_rows))
                if warn_rows:
                    st.warning("⚠️ **Téléphone ou email manquant/par défaut :**\n" + "\n".join(f"- {r}" for r in warn_rows))

            # --- Tableau PLEINE LARGEUR ---
            edited_df = st.data_editor(
                df,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Exporter": st.column_config.CheckboxColumn("✅", default=True),
                    "Urgence": st.column_config.TextColumn("Urgence", disabled=True),
                    "Pays": st.column_config.TextColumn("Pays", width="small"),
                    "Type": st.column_config.SelectboxColumn(
                        "Type",
                        options=["x1", "x2", "x3", "Multi", "Multi x3", "autre"],
                    ),
                    "Qté": st.column_config.NumberColumn("Qté", disabled=True, width="small"),
                    "Pièces": st.column_config.NumberColumn("Pièces", disabled=True, width="small"),
                    "Poids (kg)": st.column_config.NumberColumn(
                        "Poids (kg)", min_value=0.01, max_value=30.0, step=0.01, format="%.2f",
                    ),
                    "Email": st.column_config.TextColumn("Email"),
                    "Commande": st.column_config.TextColumn("Commande", disabled=True),
                },
                num_rows="fixed",
                key="import_editor",
            )

            # --- Résumé et download centrés ---
            selected_mask = edited_df["Exporter"].fillna(False)
            selected_df = edited_df[selected_mask]
            selected_indices = selected_df.index.tolist()

            export_orders = []
            for idx in selected_indices:
                o = orders[idx].copy()
                row = edited_df.iloc[idx]
                o["nom"] = row["Nom"] or ""
                o["prenom"] = row["Prénom"] or ""
                o["adresse"] = row["Adresse"] or ""
                o["complement"] = str(row["Complément"]) if row["Complément"] else ""
                o["code_postal"] = str(row["CP"]) if row["CP"] else ""
                o["ville"] = row["Ville"] or ""
                o["telephone"] = str(row["Téléphone"]) if row["Téléphone"] else ""
                o["email"] = row["Email"] or ""
                o["poids"] = float(row["Poids (kg)"])

                pays_code = str(row["Pays"]).strip().upper()
                o["pays"] = COUNTRY_MAP.get(pays_code, pays_code)
                o["pays_code"] = pays_code

                export_orders.append(o)

            _, col_bottom, _ = st.columns([1, 2, 1])
            with col_bottom:
                if export_orders:
                    total_pieces = sum(
                        o.get("total_pieces", 0) for o in export_orders
                    )
                    st.markdown(
                        f"**Export :** {len(export_orders)} / {len(orders)} colis sélectionnés"
                        + (f" — {total_pieces} pièges" if total_pieces > 0 else "")
                    )

                    today_str = date_cls.today().strftime("%Y-%m-%d")

                    if is_gls:
                        # GLS : CSV semicolon
                        csv_content = generate_gls_csv(export_orders)
                        csv_bytes = csv_content.encode("utf-8")
                        csv_name = f"import_gls_{today_str}.csv"

                        st.download_button(
                            label=f"⬇️ Télécharger {csv_name}",
                            data=csv_bytes,
                            file_name=csv_name,
                            mime="text/csv",
                            type="primary",
                        )
                    else:
                        # Happy Post : Excel
                        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp_xlsx:
                            tmp_xlsx_path = tmp_xlsx.name

                        try:
                            generate_import_file(export_orders, tmp_xlsx_path)
                            with open(tmp_xlsx_path, "rb") as f:
                                xlsx_bytes = f.read()
                        finally:
                            os.unlink(tmp_xlsx_path)

                        xlsx_name = f"import_happypost_{today_str}.xlsx"

                        st.download_button(
                            label=f"⬇️ Télécharger {xlsx_name}",
                            data=xlsx_bytes,
                            file_name=xlsx_name,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            type="primary",
                        )
                else:
                    st.info("Aucune commande sélectionnée pour l'export.")

st.divider()
st.caption(
    "Outil gratuit par [FURGO](https://shop.furgo.fr) — "
    "Fonctionne avec les PDF d'étiquettes générés par "
    "[happy-post.com](https://happy-post.com)"
)
