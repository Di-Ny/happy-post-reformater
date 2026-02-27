#!/usr/bin/env python3
"""
Reformate plusieurs fichiers PDF d'etiquettes Happy Post (1 etiquette par fichier) :
- Extrait la partie haute (etiquette de transport) de chaque PDF
- Tourne les etiquettes de 90 degres pour mieux remplir l'espace
- Produit un PDF avec 4 etiquettes par feuille A4 (grille 2x2)
- Lignes de decoupe en pointilles

Usage: python reformat_multi_etiquettes.py <fichier1.pdf> <fichier2.pdf> ... [--output output.pdf]
       python reformat_multi_etiquettes.py <dossier_contenant_les_pdfs> [--output output.pdf]
"""

import sys
import os
import glob
import fitz  # PyMuPDF


def reformat_multi_labels(pdf_paths, output_pdf=None):
    if not pdf_paths:
        print("Aucun fichier PDF fourni.")
        return

    if output_pdf is None:
        # Nommer d'apres le dossier parent ou le premier fichier
        base_dir = os.path.dirname(pdf_paths[0]) or "."
        output_pdf = os.path.join(base_dir, "etiquettes_4par_page.pdf")

    # Dimensions A4 en points
    A4_W, A4_H = 595.28, 841.89
    crop_ratio = 0.48
    margin = 12
    gap = 6

    usable_w = A4_W - 2 * margin
    usable_h = A4_H - 2 * margin
    cell_w = (usable_w - gap) / 2
    cell_h = (usable_h - gap) / 2
    cell_ratio = cell_w / cell_h

    positions = [
        (margin, margin),
        (margin + cell_w + gap, margin),
        (margin, margin + cell_h + gap),
        (margin + cell_w + gap, margin + cell_h + gap),
    ]

    dst = fitz.open()

    # Regrouper par 4
    for batch_start in range(0, len(pdf_paths), 4):
        batch = pdf_paths[batch_start:batch_start + 4]
        page = dst.new_page(width=A4_W, height=A4_H)

        for idx, pdf_path in enumerate(batch):
            src = fitz.open(pdf_path)
            src_page = src[0]
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
            src.close()

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

    dst.save(output_pdf)
    dst.close()

    n_labels = len(pdf_paths)
    n_sheets = -(-n_labels // 4)
    print(f"OK: {n_labels} etiquettes -> {n_sheets} feuille(s) A4")
    print(f"Fichier: {output_pdf}")
    return output_pdf


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python reformat_multi_etiquettes.py <fichier1.pdf> [fichier2.pdf ...] [--output out.pdf]")
        print("       python reformat_multi_etiquettes.py <dossier> [--output out.pdf]")
        sys.exit(1)

    output = None
    args = sys.argv[1:]
    if "--output" in args:
        oi = args.index("--output")
        output = args[oi + 1]
        args = args[:oi] + args[oi + 2:]

    # Si un seul argument et c'est un dossier, prendre tous les PDF dedans
    if len(args) == 1 and os.path.isdir(args[0]):
        pdf_files = sorted(glob.glob(os.path.join(args[0], "*.pdf")))
    else:
        pdf_files = args

    if not pdf_files:
        print("Aucun fichier PDF trouve.")
        sys.exit(1)

    reformat_multi_labels(pdf_files, output)
