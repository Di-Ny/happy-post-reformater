import streamlit as st
import fitz  # PyMuPDF
import io

st.set_page_config(
    page_title="Happy Post - Reformatage etiquettes",
    page_icon="📦",
    layout="centered",
)

st.title("📦 Happy Post — Reformatage d'étiquettes")
st.markdown(
    "Transformez vos étiquettes Happy Post : **4 par feuille A4** au lieu d'une seule. "
    "Économisez **~70% de papier**."
)

st.image("preview.png", use_container_width=True)

uploaded = st.file_uploader(
    "Glissez ici votre PDF d'étiquettes Happy Post",
    type=["pdf"],
)

if uploaded:
    pdf_bytes = uploaded.read()
    src = fitz.open(stream=pdf_bytes, filetype="pdf")
    n_pages = len(src)

    st.info(f"📄 {n_pages} page(s) détectée(s) dans **{uploaded.name}**")

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
    out_name = uploaded.name.replace(".pdf", "_4par_page.pdf")

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

st.divider()
st.caption(
    "Outil gratuit par [FURGO](https://shop.furgo.fr) — "
    "Fonctionne avec les PDF d'étiquettes générés par "
    "[happy-post.com](https://happy-post.com)"
)
