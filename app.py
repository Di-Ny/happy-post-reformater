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

tab_etiquettes, tab_import = st.tabs(["✂️ Reformater les étiquettes", "📋 Générer le fichier d'import"])

# =============================================================================
# ONGLET 1 : Reformatage des étiquettes
# =============================================================================
with tab_etiquettes:
    st.markdown(
        "Transformez vos étiquettes Happy Post : **4 par feuille A4** au lieu d'une seule. "
        "Économisez **~70% de papier**."
    )

    st.image("preview.png", use_container_width=True)

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

        # --- Parametres ---
        A4_W, A4_H = 595.28, 841.89
        crop_ratio = 0.48
        margin = 12
        gap = 6
        usable_w = A4_W - 2 * margin
        usable_h = A4_H - 2 * margin
        cell_w = (usable_w - gap) / 2
        cell_h = (usable_h - gap) / 2
        cell_ratio = cell_w / cell_h

        all_pages = list(range(n_pages))
        dst = fitz.open()

        positions = [
            (margin, margin),
            (margin + cell_w + gap, margin),
            (margin, margin + cell_h + gap),
            (margin + cell_w + gap, margin + cell_h + gap),
        ]

        for batch_start in range(0, len(all_pages), 4):
            batch = all_pages[batch_start:batch_start + 4]
            page = dst.new_page(width=A4_W, height=A4_H)

            for idx, src_page_num in enumerate(batch):
                src_page = src[src_page_num]
                src_rect = src_page.rect

                clip = fitz.Rect(
                    src_rect.x0,
                    src_rect.y0,
                    src_rect.x1,
                    src_rect.y0 + src_rect.height * crop_ratio,
                )

                zoom = 3.0
                mat = fitz.Matrix(zoom, zoom).prerotate(90)
                pix = src_page.get_pixmap(matrix=mat, clip=clip)

                cx, cy = positions[idx]
                this_src_w = src_rect.width
                this_src_h = src_rect.height * crop_ratio
                this_rot_ratio = this_src_h / this_src_w
                if this_rot_ratio > cell_ratio:
                    this_img_w = cell_w
                    this_img_h = cell_w / this_rot_ratio
                else:
                    this_img_h = cell_h
                    this_img_w = cell_h * this_rot_ratio

                offset_x = (cell_w - this_img_w) / 2
                offset_y = (cell_h - this_img_h) / 2
                target = fitz.Rect(
                    cx + offset_x,
                    cy + offset_y,
                    cx + offset_x + this_img_w,
                    cy + offset_y + this_img_h,
                )
                page.insert_image(target, pixmap=pix)

            # Lignes de decoupe
            mid_y = margin + cell_h + gap / 2
            mid_x = margin + cell_w + gap / 2
            gray = (0.65, 0.65, 0.65)
            dash = "[3 3] 0"
            shape = page.new_shape()
            shape.draw_line(fitz.Point(margin, mid_y), fitz.Point(A4_W - margin, mid_y))
            shape.finish(color=gray, width=0.4, dashes=dash)
            shape.draw_line(fitz.Point(mid_x, margin), fitz.Point(mid_x, A4_H - margin))
            shape.finish(color=gray, width=0.4, dashes=dash)

            mark = 6
            for corner_x in [margin, A4_W - margin]:
                for corner_y in [margin, A4_H - margin]:
                    dx = mark if corner_x == margin else -mark
                    dy = mark if corner_y == margin else -mark
                    shape.draw_line(fitz.Point(corner_x, corner_y), fitz.Point(corner_x + dx, corner_y))
                    shape.finish(color=gray, width=0.3)
                    shape.draw_line(fitz.Point(corner_x, corner_y), fitz.Point(corner_x, corner_y + dy))
                    shape.finish(color=gray, width=0.3)
            shape.commit()

        # Generer le PDF en memoire
        out_buf = io.BytesIO()
        dst.save(out_buf)
        dst.close()
        src.close()
        out_buf.seek(0)

        n_sheets = -(-n_pages // 4)
        out_name = uploaded_labels.name.replace(".pdf", "_4par_page.pdf")

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
# ONGLET 2 : Génération du fichier d'import
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

            # Tableau recapitulatif
            label_map = {0.31: "x1", 0.32: "x2", 0.35: "x3", 0.34: "Multi", 0.50: "Multi x3"}
            recap_data = []
            for i, o in enumerate(orders, 1):
                recap_data.append({
                    "#": i,
                    "Nom": f"{o['nom']} {o['prenom']}",
                    "Adresse": o["adresse"],
                    "CP": o["code_postal"],
                    "Ville": o["ville"],
                    "Type": label_map.get(o["poids"], "?"),
                    "Poids": f"{o['poids']} kg",
                })

            st.dataframe(recap_data, use_container_width=True, hide_index=True)

            # Compter par type
            from collections import Counter
            types = Counter(label_map.get(o["poids"], "?") for o in orders)
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

            st.warning("⚠️ **Vérifiez toujours le fichier avant de l'importer sur Happy Post !**")

st.divider()
st.caption(
    "Outil gratuit par [FURGO](https://shop.furgo.fr) — "
    "Fonctionne avec les PDF d'étiquettes générés par "
    "[happy-post.com](https://happy-post.com)"
)
