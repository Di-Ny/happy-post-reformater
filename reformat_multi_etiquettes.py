#!/usr/bin/env python3
"""
Reformate plusieurs fichiers PDF d'etiquettes Happy Post (1 etiquette par fichier) :
- Extrait la partie haute (etiquette de transport) de chaque PDF
- Oriente les etiquettes automatiquement pour mieux remplir l'espace
- Produit un PDF avec 4 ou 6 etiquettes par feuille A4
- Lignes de decoupe en pointilles

Usage: python reformat_multi_etiquettes.py <fichier1.pdf> <fichier2.pdf> ... [--output output.pdf] [--format 4|6]
       python reformat_multi_etiquettes.py <dossier_contenant_les_pdfs> [--output output.pdf] [--format 6]
"""

import sys
import os
import glob
import fitz  # PyMuPDF


def get_layout_config(labels_per_page):
    """Retourne la config de grille pour 4 ou 6 etiquettes par page."""
    A4_W, A4_H = 595.28, 841.89
    if labels_per_page == 6:
        # Avery L7166 : 99.1mm x 93.1mm, grille 2x3
        cell_w_mm, cell_h_mm = 99.1, 93.1
        margin_x = (210 - 2 * cell_w_mm) / 2
        margin_y = (297 - 3 * cell_h_mm) / 2
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
    ratio_p = src_w / src_h
    cell_ratio = cell_w / cell_h
    if ratio_p > cell_ratio:
        fill_p = cell_w * (cell_w / ratio_p)
    else:
        fill_p = cell_h * (cell_h * ratio_p)
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
    best_rect = None
    best_area = 0
    for p in src_page.get_drawings():
        r = p["rect"]
        area = r.width * r.height
        if area > best_area and r.width > 200 and r.height > 150:
            best_rect = r
            best_area = area

    if best_rect:
        margin_pt = 8
        x0 = max(best_rect.x0 - 50, 0)
        y0 = max(best_rect.y0 - margin_pt, 0)
        x1 = min(best_rect.x1 + margin_pt, src_rect.width)
        y1 = min(best_rect.y1 + 25, src_rect.height)
        return fitz.Rect(x0, y0, x1, y1)

    return fitz.Rect(
        src_rect.x0, src_rect.y0,
        src_rect.x1, src_rect.y0 + src_rect.height * 0.47,
    )


def draw_cut_guides(page, cfg):
    """Dessine les lignes de decoupe et reperes de coins."""
    A4_W, A4_H = cfg["A4_W"], cfg["A4_H"]
    cell_w, cell_h = cfg["cell_w"], cfg["cell_h"]
    mx, my = cfg["margin_x"], cfg["margin_y"]
    cols, rows = cfg["cols"], cfg["rows"]
    gray = (0.65, 0.65, 0.65)
    dash = "[3 3] 0"
    shape = page.new_shape()

    for r in range(1, rows):
        y = my + r * cell_h
        shape.draw_line(fitz.Point(mx, y), fitz.Point(A4_W - mx, y))
        shape.finish(color=gray, width=0.4, dashes=dash)

    for c in range(1, cols):
        x = mx + c * cell_w
        shape.draw_line(fitz.Point(x, my), fitz.Point(x, A4_H - my))
        shape.finish(color=gray, width=0.4, dashes=dash)

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


def reformat_multi_labels(pdf_paths, output_pdf=None, labels_per_page=4):
    if not pdf_paths:
        print("Aucun fichier PDF fourni.")
        return

    if output_pdf is None:
        base_dir = os.path.dirname(pdf_paths[0]) or "."
        output_pdf = os.path.join(base_dir, f"etiquettes_{labels_per_page}par_page.pdf")

    cfg = get_layout_config(labels_per_page)
    dst = fitz.open()

    for batch_start in range(0, len(pdf_paths), labels_per_page):
        batch = pdf_paths[batch_start:batch_start + labels_per_page]
        page = dst.new_page(width=cfg["A4_W"], height=cfg["A4_H"])

        for idx, pdf_path in enumerate(batch):
            src = fitz.open(pdf_path)
            src_page = src[0]

            clip = smart_crop(src_page)
            cell_w, cell_h = cfg["cell_w"], cfg["cell_h"]
            this_rot_ratio, rotate = best_orientation(clip.width, clip.height, cell_w, cell_h)

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
            src.close()

        draw_cut_guides(page, cfg)

    dst.save(output_pdf)
    dst.close()

    n_labels = len(pdf_paths)
    n_sheets = -(-n_labels // labels_per_page)
    print(f"OK: {n_labels} etiquettes -> {n_sheets} feuille(s) A4 ({labels_per_page} par page)")
    print(f"Fichier: {output_pdf}")
    return output_pdf


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python reformat_multi_etiquettes.py <fichier1.pdf> [fichier2.pdf ...] [--output out.pdf] [--format 4|6]")
        print("       python reformat_multi_etiquettes.py <dossier> [--output out.pdf] [--format 6]")
        sys.exit(1)

    output = None
    fmt = 4
    args = sys.argv[1:]
    if "--output" in args:
        oi = args.index("--output")
        output = args[oi + 1]
        args = args[:oi] + args[oi + 2:]
    if "--format" in args:
        fi = args.index("--format")
        fmt = int(args[fi + 1])
        args = args[:fi] + args[fi + 2:]
        if fmt not in (4, 6):
            print("Erreur: --format doit etre 4 ou 6")
            sys.exit(1)

    # Si un seul argument et c'est un dossier, prendre tous les PDF dedans
    if len(args) == 1 and os.path.isdir(args[0]):
        pdf_files = sorted(glob.glob(os.path.join(args[0], "*.pdf")))
    else:
        pdf_files = args

    if not pdf_files:
        print("Aucun fichier PDF trouve.")
        sys.exit(1)

    reformat_multi_labels(pdf_files, output, fmt)
