#!/usr/bin/env python3
"""
Reformate les etiquettes Happy Post :
- Extrait la partie haute (etiquette de transport) de chaque page
- Tourne les etiquettes de 90 degres pour mieux remplir l'espace
- Produit un PDF avec 4 etiquettes par feuille A4 (grille 2x2)
- Lignes de decoupe en pointilles

Usage: python reformat_etiquettes.py <fichier_etiquettes.pdf> [output.pdf]
"""

import sys
import os
import fitz  # PyMuPDF


def reformat_labels(input_pdf, output_pdf=None):
    if output_pdf is None:
        base = os.path.splitext(input_pdf)[0]
        output_pdf = f"{base}_4par_page.pdf"

    src = fitz.open(input_pdf)

    # Dimensions A4 en points
    A4_W, A4_H = 595.28, 841.89

    # Toutes les pages sont des etiquettes (y compris page 0 MasterColis)
    all_pages = list(range(len(src)))

    if not all_pages:
        print("Aucune page trouvee.")
        return

    dst = fitz.open()

    # Page 0 (MasterColis) : crop_ratio different (la partie haute est plus grande)
    # Pages 1+ (colis simples) : crop_ratio ~48%
    def get_crop_ratio(page_idx):
        if page_idx == 0:
            return 0.48  # MasterColis : partie haute similaire
        return 0.48

    # Marge et espacement
    margin = 12
    gap = 6

    # Zone utile
    usable_w = A4_W - 2 * margin
    usable_h = A4_H - 2 * margin

    # Cellule 2x2
    cell_w = (usable_w - gap) / 2
    cell_h = (usable_h - gap) / 2

    # L'etiquette source est en portrait (~595 x 403 apres crop)
    # On la tourne de 90 degres -> elle devient ~403 x 595 (paysage dans la cellule)
    # La cellule fait environ 283 x 412
    # Apres rotation, le ratio de l'etiquette = hauteur_originale / largeur_originale
    sample_page = src[1] if len(src) > 1 else src[0]
    src_w = sample_page.rect.width   # ~595
    src_h = sample_page.rect.height * 0.48  # ~403

    # Apres rotation 90 CW : nouvelle largeur = src_h, nouvelle hauteur = src_w
    rot_w = src_h  # ~403
    rot_h = src_w  # ~595
    rot_ratio = rot_w / rot_h  # ~0.68

    # Ajuster dans la cellule en gardant les proportions
    cell_ratio = cell_w / cell_h
    if rot_ratio > cell_ratio:
        img_w = cell_w
        img_h = cell_w / rot_ratio
    else:
        img_h = cell_h
        img_w = cell_h * rot_ratio

    # Regrouper par 4
    for batch_start in range(0, len(all_pages), 4):
        batch = all_pages[batch_start:batch_start + 4]

        page = dst.new_page(width=A4_W, height=A4_H)

        positions = [
            (margin, margin),
            (margin + cell_w + gap, margin),
            (margin, margin + cell_h + gap),
            (margin + cell_w + gap, margin + cell_h + gap),
        ]

        for idx, src_page_num in enumerate(batch):
            src_page = src[src_page_num]
            src_rect = src_page.rect
            crop_r = get_crop_ratio(src_page_num)

            # Zone a extraire : partie haute
            clip = fitz.Rect(
                src_rect.x0,
                src_rect.y0,
                src_rect.x1,
                src_rect.y0 + src_rect.height * crop_r
            )

            # Render avec rotation 90 degres (sens horaire)
            # Matrix : zoom * rotation
            zoom = 3.0
            mat = fitz.Matrix(zoom, zoom).prerotate(90)
            pix = src_page.get_pixmap(matrix=mat, clip=clip)

            # Centrer dans la cellule
            cx, cy = positions[idx]

            # Recalculer img dims pour cette etiquette specifique
            this_src_w = src_rect.width
            this_src_h = src_rect.height * crop_r
            this_rot_ratio = this_src_h / this_src_w  # apres rotation
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
                cy + offset_y + this_img_h
            )

            page.insert_image(target, pixmap=pix)

        # Lignes de decoupe
        mid_y = margin + cell_h + gap / 2
        mid_x = margin + cell_w + gap / 2
        gray = (0.65, 0.65, 0.65)
        dash = "[3 3] 0"

        shape = page.new_shape()

        # Ligne horizontale
        shape.draw_line(fitz.Point(margin, mid_y), fitz.Point(A4_W - margin, mid_y))
        shape.finish(color=gray, width=0.4, dashes=dash)

        # Ligne verticale
        shape.draw_line(fitz.Point(mid_x, margin), fitz.Point(mid_x, A4_H - margin))
        shape.finish(color=gray, width=0.4, dashes=dash)

        # Reperes de coupe aux coins
        mark = 6
        for cx in [margin, A4_W - margin]:
            for cy in [margin, A4_H - margin]:
                dx = mark if cx == margin else -mark
                dy = mark if cy == margin else -mark
                shape.draw_line(fitz.Point(cx, cy), fitz.Point(cx + dx, cy))
                shape.finish(color=gray, width=0.3)
                shape.draw_line(fitz.Point(cx, cy), fitz.Point(cx, cy + dy))
                shape.finish(color=gray, width=0.3)

        shape.commit()

    dst.save(output_pdf)
    dst.close()
    src.close()

    n_labels = len(all_pages)
    n_sheets = -(-n_labels // 4)
    print(f"OK: {n_labels} etiquettes -> {n_sheets} feuille(s) A4")
    print(f"Fichier: {output_pdf}")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python reformat_etiquettes.py <fichier.pdf> [output.pdf]")
        sys.exit(1)

    input_file = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) > 2 else None
    reformat_labels(input_file, output_file)
