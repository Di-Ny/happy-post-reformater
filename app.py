import streamlit as st
import fitz  # PyMuPDF
import io
import os

st.set_page_config(
    page_title="Happy Post - Outils d'expedition",
    page_icon="📦",
    layout="centered",
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
    st.markdown(
        "Générez le fichier d'import Happy Post (.xlsx) à partir de vos bons de commande Amazon (PDF). "
        "Seules les commandes **Belgique** sont extraites."
    )

    uploaded_amazon = st.file_uploader(
        "Glissez ici votre PDF de commandes Amazon",
        type=["pdf"],
        key="amazon_uploader",
    )

    if uploaded_amazon:
        # Sauvegarder temporairement le PDF pour pdfplumber
        import tempfile
        from generate_import import parse_orders_from_pdf_text, generate_import_file

        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
            tmp.write(uploaded_amazon.read())
            tmp_path = tmp.name

        try:
            orders = parse_orders_from_pdf_text(tmp_path)
        finally:
            os.unlink(tmp_path)

        if not orders:
            st.warning("⚠️ Aucune commande Belgique trouvée dans ce PDF.")
        else:
            st.success(f"✅ **{len(orders)} commande(s) Belgique** extraites")

            # Tableau éditable
            import pandas as pd
            label_map = {0.31: "x1", 0.32: "x2", 0.35: "x3", 0.34: "Multi", 0.50: "Multi x3"}
            weight_map = {"x1": 0.31, "x2": 0.32, "x3": 0.35, "Multi": 0.34, "Multi x3": 0.50}

            edit_data = []
            for i, o in enumerate(orders):
                edit_data.append({
                    "Nom": o["nom"],
                    "Prénom": o["prenom"],
                    "Entreprise": o["entreprise"],
                    "Adresse": o["adresse"],
                    "Complément": o["complement"],
                    "CP": o["code_postal"],
                    "Ville": o["ville"],
                    "Province": o["province"],
                    "Téléphone": o["telephone"],
                    "Type": label_map.get(o["poids"], "?"),
                })

            df = pd.DataFrame(edit_data)

            # Signaler les champs vides importants
            missing_rows = []
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
                if missing:
                    missing_rows.append(f"**Ligne {i+1}** ({row['Nom']} {row['Prénom']}) : {', '.join(missing)}")

            if missing_rows:
                st.warning("⚠️ **Champs manquants** (corrigez ci-dessous) :\n" + "\n".join(f"- {r}" for r in missing_rows))

            edited_df = st.data_editor(
                df,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Type": st.column_config.SelectboxColumn(
                        options=["x1", "x2", "x3", "Multi", "Multi x3"],
                    ),
                },
                num_rows="fixed",
            )

            # Réinjecter les modifications dans orders
            for i, o in enumerate(orders):
                row = edited_df.iloc[i]
                o["nom"] = row["Nom"] or ""
                o["prenom"] = row["Prénom"] or ""
                o["entreprise"] = row["Entreprise"] or ""
                o["adresse"] = row["Adresse"] or ""
                o["complement"] = row["Complément"] or ""
                o["code_postal"] = row["CP"] or ""
                o["ville"] = row["Ville"] or ""
                o["province"] = row["Province"] or ""
                o["telephone"] = row["Téléphone"] or ""
                o["poids"] = weight_map.get(row["Type"], o["poids"])

            # Compter par type
            from collections import Counter
            types = Counter(edited_df["Type"])
            total_pieges = sum(
                count * {"x1": 1, "x2": 2, "x3": 3, "Multi": 4, "Multi x3": 6}.get(t, 1)
                for t, count in types.items()
            )

            st.markdown(f"**Résumé :** {len(orders)} colis — ~{total_pieges} pièges au total")

            # Generer le fichier Excel en memoire
            template_path = os.path.join(os.path.dirname(__file__), "templates", "template_import_colis.xlsx")
            with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp_xlsx:
                tmp_xlsx_path = tmp_xlsx.name

            try:
                generate_import_file(orders, tmp_xlsx_path)
                with open(tmp_xlsx_path, "rb") as f:
                    xlsx_bytes = f.read()
            finally:
                os.unlink(tmp_xlsx_path)

            from datetime import date
            today = date.today().strftime("%Y-%m-%d")
            xlsx_name = f"import_happypost_{today}.xlsx"

            st.download_button(
                label=f"⬇️ Télécharger {xlsx_name}",
                data=xlsx_bytes,
                file_name=xlsx_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
            )

st.divider()
st.caption(
    "Outil gratuit par [FURGO](https://shop.furgo.fr) — "
    "Fonctionne avec les PDF d'étiquettes générés par "
    "[happy-post.com](https://happy-post.com)"
)
