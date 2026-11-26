#!/usr/bin/env python3
"""
DocConvert Local Engine — Conversor de Documentos de Alta Fidelidade (PyMuPDF + python-docx)
===========================================================================================
Complemento ao `index.html` usando bibliotecas nativas já instaladas:
  - PyMuPDF (`pymupdf`) para extração fiel de PDF (geometria, recuo de parágrafos,
    entrelinha real, 2 colunas, negrito/itálico e imagens reais na posição exata)
  - python-docx para leitura/escrita nativa de Word (.docx) com imagens embutidas
  - BeautifulSoup4 + lxml para XHTML estrito (EPUB 3 + toc.ncx)
  - Flask para servir como motor local conectado automaticamente ao `index.html`

Uso via CLI:
  python converter_local.py entrada.pdf saida.epub
  python converter_local.py entrada.epub saida.docx
  python converter_local.py entrada.pdf saida.md --title "Meu Livro" --author "Autor"

Uso como servidor local conectado ao index.html:
  python converter_local.py --serve
"""

import argparse
import base64
import html as html_lib
import io
import json
import re
import statistics
import sys
import uuid
import zipfile
from datetime import datetime, timezone
from pathlib import Path

import pymupdf as fitz  # PyMuPDF
from bs4 import BeautifulSoup
from docx import Document as DocxDocument
from docx.shared import Inches


# =============================================================================
# 1. UTILITÁRIOS E PARSERS -> IR CANÔNICO {title, author, lang, chapters[], images[], cover_image}
# =============================================================================

CHAPTER_RE = re.compile(
    r"^(cap[ií]tulo|chapter|parte|part|se[çc][ãa]o|pr[oó]logo|ep[ií]logo|pref[áa]cio|introdu[çc][ãa]o|conclus[ãa]o|ap[êe]ndice)\b",
    re.IGNORECASE,
)
DIALOGUE_RE = re.compile(r"^[—–―]\s*\S|^-\s+[A-ZÁÉÍÓÚÂÊÔÃÕÇ\"“«]")
BULLET_RE = re.compile(r"^([•●▪◦‣\*]|\d{1,3}[.)]|[a-z][)])\s+")
TERMINAL_PUNCT_RE = re.compile(r"[.!?:»”\"…)\]]$")


def _percentile(sorted_vals, p: float, default: float = 0.0) -> float:
    if not sorted_vals:
        return default
    idx = min(len(sorted_vals) - 1, max(0, int(len(sorted_vals) * p)))
    return sorted_vals[idx]


def _split_html_chapters(full_html: str, fallback_title: str, split_chapters: bool = True):
    if not split_chapters:
        return [{"id": "chap_1", "title": fallback_title, "html": full_html}]

    soup = BeautifulSoup(full_html, "html.parser")
    h1s = soup.find_all("h1")
    split_tag = "h1" if len(h1s) >= 2 else ("h2" if len(soup.find_all("h2")) >= 2 else None)

    if not split_tag:
        first_h = soup.find(["h1", "h2"])
        return [{
            "id": "chap_1",
            "title": first_h.get_text(strip=True) if first_h else fallback_title,
            "html": full_html,
        }]

    chapters = []
    cur_title = fallback_title
    cur_parts = []

    def _chunk_stats(parts):
        body_chars = 0
        has_img = False
        last_page = None
        for p in parts:
            name = getattr(p, "name", None)
            if not name:
                continue
            if p.get("data-page"):
                last_page = str(p.get("data-page"))
            if name in ("img", "figure") or p.find("img"):
                has_img = True
            if not re.match(r"^h[1-6]$", name):
                body_chars += len(p.get_text(strip=True))
        return body_chars, has_img, last_page

    for node in soup.children:
        if getattr(node, "name", None) == split_tag:
            body_chars, has_img, last_page = _chunk_stats(cur_parts)
            node_page = str(node.get("data-page")) if node.get("data-page") else None
            has_page_info = bool(last_page and node_page)
            same_page = has_page_info and last_page == node_page
            diff_page = has_page_info and last_page != node_page
            only_headings = body_chars < 25 and not has_img

            should_split = bool(cur_parts) and (
                diff_page or (not same_page and not only_headings) or (same_page and body_chars >= 80)
            )
            if should_split:
                chunk = "\n".join(str(p) for p in cur_parts).strip()
                s_chunk = BeautifulSoup(chunk, "html.parser")
                if s_chunk.get_text(strip=True) or s_chunk.find("img"):
                    chapters.append({"id": f"chap_{len(chapters)+1}", "title": cur_title, "html": chunk})
                cur_title = node.get_text(strip=True) or f"Capítulo {len(chapters)+1}"
                cur_parts = [node]
            else:
                if not cur_parts:
                    cur_title = node.get_text(strip=True) or cur_title
                cur_parts.append(node)
        else:
            cur_parts.append(node)

    if cur_parts:
        chunk = "\n".join(str(p) for p in cur_parts).strip()
        s_chunk = BeautifulSoup(chunk, "html.parser")
        if s_chunk.get_text(strip=True) or s_chunk.find("img"):
            chapters.append({"id": f"chap_{len(chapters)+1}", "title": cur_title, "html": chunk})

    return chapters or [{"id": "chap_1", "title": fallback_title, "html": full_html}]


def _extract_pdf_page_images(doc, page, page_num: int, text_chars: int, img_counter: list) -> list:
    """Extrai imagens reais da página PDF em sua posição vertical exata (y0)."""
    page_w = page.rect.width or 1.0
    page_h = page.rect.height or 1.0
    page_area = page_w * page_h
    extracted = []
    seen_rects = []

    for img_info in page.get_images(full=True):
        xref = img_info[0]
        rects = page.get_image_rects(xref)
        if not rects:
            continue

        # Extrai pixmap tratando CMYK e SMask (transparência) corretamente
        try:
            smask = img_info[1]
            pix = fitz.Pixmap(doc, xref)
            if smask > 0:
                mask = fitz.Pixmap(doc, smask)
                if pix.n >= 4 and pix.colorspace != fitz.csRGB:
                    pix = fitz.Pixmap(fitz.csRGB, pix)
                pix = fitz.Pixmap(pix, mask)
                mask = None
            if pix.colorspace and pix.colorspace.n > 3:
                pix = fitz.Pixmap(fitz.csRGB, pix)
            if pix.alpha:
                # Compõe sobre fundo branco para JPEG limpo
                white = fitz.Pixmap(fitz.csRGB, pix.irect, False)
                white.clear_with(255)
                white.copy(pix, pix.irect)
                img_bytes = pix.tobytes("png")
                mime, ext = "image/png", "png"
            else:
                img_bytes = pix.tobytes("jpeg", jpg_quality=88)
                mime, ext = "image/jpeg", "jpg"
            w_px, h_px = pix.width, pix.height
            pix = None
        except Exception:
            continue

        if w_px < 40 or h_px < 40:
            continue

        for rect in rects:
            if rect.width < 36 or rect.height < 36:
                continue
            area_ratio = (rect.width * rect.height) / page_area
            # Ignora fundo/marca d'água de página inteira se houver texto normal
            if rect.width > page_w * 0.92 and rect.height > page_h * 0.92 and text_chars > 80:
                continue
            # Evita duplicar exatamente o mesmo retângulo na mesma página
            key = (round(rect.x0, 1), round(rect.y0, 1), round(rect.x1, 1), round(rect.y1, 1))
            if key in seen_rects:
                continue
            seen_rects.append(key)

            img_counter[0] += 1
            img_id = f"img_{img_counter[0]:03d}"
            fname = f"{img_id}.{ext}"
            extracted.append({
                "id": img_id,
                "filename": fname,
                "mime": mime,
                "bytes": img_bytes,
                "width": w_px,
                "height": h_px,
                "page": page_num,
                "y_top": rect.y0,
                "area_ratio": area_ratio,
                "alt": f"Figura {img_counter[0]} (Página {page_num})",
            })

    # Fallback para página escaneada sem texto
    if not extracted and text_chars < 25:
        pix = page.get_pixmap(dpi=150)
        img_bytes = pix.tobytes("jpeg", jpg_quality=86)
        img_counter[0] += 1
        img_id = f"img_{img_counter[0]:03d}"
        extracted.append({
            "id": img_id,
            "filename": f"{img_id}.jpg",
            "mime": "image/jpeg",
            "bytes": img_bytes,
            "width": pix.width,
            "height": pix.height,
            "page": page_num,
            "y_top": 0.0,
            "area_ratio": 1.0,
            "alt": f"Página {page_num}",
        })

    return extracted


def _order_columns(lines: list, page_w: float) -> list:
    if len(lines) < 8:
        return [{**l, "col": 0} for l in lines]
    mid = page_w / 2.0
    left = [l for l in lines if l["x1"] <= mid + 8]
    right = [l for l in lines if l["x0"] >= mid - 8]
    span = [l for l in lines if l["x1"] > mid + 8 and l["x0"] < mid - 8]
    if len(left) >= 4 and len(right) >= 4 and len(span) <= max(3, int(len(lines) * 0.22)):
        top_y = min(left[0]["y_top"], right[0]["y_top"])
        bot_y = max(left[-1]["y_top"], right[-1]["y_top"])
        top_span = [{**l, "col": 0} for l in span if l["y_top"] < top_y]
        mid_span = [{**l, "col": 0} for l in span if top_y <= l["y_top"] <= bot_y]
        bot_span = [{**l, "col": 0} for l in span if l["y_top"] > bot_y]
        return top_span + [{**l, "col": 1} for l in left] + [{**l, "col": 2} for l in right] + mid_span + bot_span
    return [{**l, "col": 0} for l in lines]


def parse_pdf(data: bytes, filename: str, opts: dict) -> dict:
    doc = fitz.open(stream=data, filetype="pdf")
    meta = doc.metadata or {}

    font_sizes = []
    pages_data = []
    all_images = []
    img_counter = [0]
    cover_image = None
    extract_imgs = opts.get("extractImages", True)

    for page_idx, page in enumerate(doc, start=1):
        page_w = page.rect.width or 595.0
        page_h = page.rect.height or 842.0
        blocks = page.get_text("dict", sort=True).get("blocks", [])
        lines_out = []
        page_chars = 0

        for b in blocks:
            if b.get("type") != 0:
                continue
            for line in b.get("lines", []):
                spans = line.get("spans", [])
                plain_parts = []
                html_parts = []
                max_size = 0.0
                bold_chars = 0
                total_chars = 0
                for s in spans:
                    t = s.get("text", "")
                    if not t:
                        continue
                    plain_parts.append(t)
                    flags = s.get("flags", 0)
                    fname_s = (s.get("font") or "").lower()
                    is_b = bool(flags & 16) or any(k in fname_s for k in ("bold", "black", "heavy", "demi"))
                    is_i = bool(flags & 2) or any(k in fname_s for k in ("italic", "oblique"))
                    esc = html_lib.escape(t)
                    if is_b and is_i:
                        html_parts.append(f"<strong><em>{esc}</em></strong>")
                    elif is_b:
                        html_parts.append(f"<strong>{esc}</strong>")
                    elif is_i:
                        html_parts.append(f"<em>{esc}</em>")
                    else:
                        html_parts.append(esc)
                    max_size = max(max_size, float(s.get("size", 11.0)))
                    l_t = len(t.strip())
                    total_chars += l_t
                    if is_b:
                        bold_chars += l_t

                text = re.sub(r"\s+", " ", "".join(plain_parts)).strip()
                if not text:
                    continue
                page_chars += len(text)
                bbox = line.get("bbox", [0, 0, 0, 0])
                max_size = max_size or 11.0
                font_sizes.append(round(max_size, 1))
                line_html = re.sub(r"</strong>(\s*)<strong>", r"\1", "".join(html_parts))
                line_html = re.sub(r"</em>(\s*)<em>", r"\1", line_html)
                line_html = re.sub(r"\s+", " ", line_html).strip()

                lines_out.append({
                    "kind": "line",
                    "page": page_idx,
                    "text": text,
                    "html": line_html,
                    "size": max_size,
                    "is_bold": total_chars > 0 and (bold_chars / total_chars) > 0.72,
                    "x0": bbox[0],
                    "x1": bbox[2],
                    "width": max(0.0, bbox[2] - bbox[0]),
                    "y_top": bbox[1],
                    "norm_y": bbox[1] / page_h,
                })

        lines_out.sort(key=lambda l: (l["y_top"], l["x0"]))
        ordered_lines = _order_columns(lines_out, page_w)

        page_imgs = []
        if extract_imgs:
            page_imgs = _extract_pdf_page_images(doc, page, page_idx, page_chars, img_counter)
            if page_idx == 1 and not cover_image and page_imgs and page_imgs[0]["area_ratio"] > 0.40:
                cover_image = page_imgs[0]
            all_images.extend(page_imgs)

        pages_data.append({"page_num": page_idx, "page_w": page_w, "page_h": page_h, "lines": ordered_lines, "images": page_imgs})

    # Se não havia capa imagem na página 1 mas genCover=True, renderiza página 1 como capa
    cover_bytes = None
    if opts.get("genCover", True) and not cover_image and len(doc) > 0:
        pix = doc[0].get_pixmap(dpi=130)
        cover_bytes = pix.tobytes("jpeg", jpg_quality=86)

    median_font = statistics.median(font_sizes) if font_sizes else 11.0

    # Mede entrelinha real (leading) e margens reais
    leadings, x0s, x1s = [], [], []
    for pd in pages_data:
        prev = None
        for l in pd["lines"]:
            is_body = abs(l["size"] - median_font) <= max(1.2, median_font * 0.14)
            if is_body and 0.08 < l["norm_y"] < 0.92 and len(l["text"]) > 20:
                x0s.append(l["x0"])
                x1s.append(l["x1"])
                if prev and prev["col"] == l["col"]:
                    dy = l["y_top"] - prev["y_top"]
                    if median_font * 0.75 <= dy <= median_font * 2.35:
                        leadings.append(dy)
                prev = l
            else:
                prev = None

    leadings.sort()
    x0s.sort()
    x1s.sort()
    median_leading = _percentile(leadings, 0.5, median_font * 1.32)
    global_left = _percentile(x0s, 0.20, 50.0)
    global_right = _percentile(x1s, 0.85, 520.0)
    global_width = max(120.0, global_right - global_left)

    # Detectar cabeçalhos/rodapés repetidos
    repeated_edges = set()
    if opts.get("stripHF", True) and len(doc) >= 3:
        freq = {}
        for pd in pages_data:
            seen = set()
            for l in pd["lines"]:
                if l["norm_y"] < 0.095 or l["norm_y"] > 0.905:
                    key = re.sub(r"\d+", "#", l["text"]).strip().lower()
                    if key and len(key) < 95 and key not in seen:
                        seen.add(key)
                        freq[key] = freq.get(key, 0) + 1
        threshold = max(2, int(len(doc) * 0.35))
        repeated_edges = {k for k, v in freq.items() if v >= threshold}

    html_parts = []
    cur_text = ""
    cur_html = ""
    prev_line = None

    def flush():
        nonlocal cur_text, cur_html
        if cur_html.strip():
            html_parts.append(f"<p>{cur_html.strip()}</p>")
            cur_text = ""
            cur_html = ""

    for pd in pages_data:
        # Margens por coluna nesta página (suporta margens espelhadas par/ímpar)
        col_margins = {}
        for cid in (0, 1, 2):
            clines = [l for l in pd["lines"] if l["col"] == cid and abs(l["size"] - median_font) <= 1.5 and 0.09 < l["norm_y"] < 0.91 and len(l["text"]) > 18]
            if len(clines) >= 3:
                col_margins[cid] = (
                    _percentile(sorted(l["x0"] for l in clines), 0.20, global_left),
                    _percentile(sorted(l["x1"] for l in clines), 0.85, global_right),
                )
            else:
                col_margins[cid] = (global_left, global_right)

        has_two_cols = any(l["col"] == 2 for l in pd["lines"])
        if not has_two_cols:
            items = list(pd["lines"]) + [{"kind": "image", "y_top": im["y_top"], "img": im} for im in pd["images"]]
            items.sort(key=lambda it: it["y_top"])
        else:
            s_imgs = sorted([{"kind": "image", "y_top": im["y_top"], "img": im} for im in pd["images"]], key=lambda x: x["y_top"])
            items, idx_im = [], 0
            for l in pd["lines"]:
                while idx_im < len(s_imgs) and s_imgs[idx_im]["y_top"] <= l["y_top"] and l["col"] != 2:
                    items.append(s_imgs[idx_im])
                    idx_im += 1
                items.append(l)
            items.extend(s_imgs[idx_im:])

        for it in items:
            if it["kind"] == "image":
                im = it["img"]
                if cover_image and im["id"] == cover_image["id"] and opts.get("genCover", True) and len(pd["lines"]) <= 4:
                    continue
                flush()
                prev_line = None
                html_parts.append(f'<figure class="doc-figure"><img src="images/{im["filename"]}" data-img-id="{im["id"]}" alt="{html_lib.escape(im["alt"])}"/></figure>')
                continue

            l = it
            txt = l["text"]

            if opts.get("stripHF", True):
                is_edge = l["norm_y"] < 0.095 or l["norm_y"] > 0.905
                is_page_num = bool(re.match(r"^[-–—\s]*(p[áa]g\.?|page)?\s*\d+\s*[-–—\s]*$", txt, re.I))
                key = re.sub(r"\d+", "#", txt).strip().lower()
                if is_edge and (is_page_num or key in repeated_edges):
                    continue

            col_left_pre, col_right_pre = col_margins.get(l["col"], (global_left, global_right))
            col_w_pre = max(100.0, col_right_pre - col_left_pre)
            line_centered = l["width"] < col_w_pre * 0.78 and abs((l["x0"] - col_left_pre) - (col_right_pre - l["x1"])) < 22 and (l["x0"] - col_left_pre) > 18
            is_author_credit = bool(re.match(r"^(by|por|written by|escrito por|tradu[çc][ãa]o de)\s+", txt, re.I))

            h_lvl = 0
            is_sub = False
            if opts.get("detectHeadings", True) and len(txt) < 150:
                if is_author_credit and l["size"] >= median_font * 1.05:
                    h_lvl = 2
                    is_sub = True
                elif l["size"] >= median_font * 1.38 or (CHAPTER_RE.match(txt) and l["size"] >= median_font * 1.04 and len(txt) < 100):
                    h_lvl = 1
                elif l["size"] >= median_font * 1.16 or (l["is_bold"] and len(txt) < 85 and not txt.endswith((".", ",", ";")) and l["width"] < global_width * 0.80):
                    h_lvl = 2

            if h_lvl:
                flush()
                classes = []
                if is_sub:
                    classes.append("subtitle")
                if line_centered:
                    classes.append("centered")
                cls_attr = f' class="{" ".join(classes)}"' if classes else ""
                inner = l["html"] if is_sub else html_lib.escape(txt)
                html_parts.append(f'<h{h_lvl}{cls_attr} data-page="{l["page"]}">{inner}</h{h_lvl}>')
                prev_line = None
                continue

            col_left, col_right = col_margins.get(l["col"], (global_left, global_right))
            col_w = max(100.0, col_right - col_left)

            if BULLET_RE.match(txt):
                flush()
                cur_text = txt
                cur_html = l["html"]
                prev_line = l
                continue

            if not opts.get("unwrap", True):
                flush()
                html_parts.append(f"<p>{l['html']}</p>")
                prev_line = l
                continue

            # Decisão precisa de quebra de parágrafo (mesmos 5 sinais do motor JS!)
            start_new = False
            if not cur_text or not prev_line:
                start_new = True
            else:
                same_pc = (prev_line["page"] == l["page"] and prev_line["col"] == l["col"])
                dy = (l["y_top"] - prev_line["y_top"]) if same_pc else 0.0
                has_vgap = same_pc and dy > (median_leading * 1.24)

                indent_delta = l["x0"] - col_left
                is_indented = max(median_font * 0.72, 7.5) <= indent_delta <= col_w * 0.35

                first_word = txt.split()[0] if txt.split() else ""
                approx_fw = len(first_word) * median_font * 0.50
                right_gap = col_right - prev_line["x1"]
                prev_short = right_gap > max(approx_fw + median_font * 1.6, col_w * 0.14)
                prev_ended = bool(TERMINAL_PUNCT_RE.search(prev_line["text"]))

                is_dialogue = bool(DIALOGUE_RE.match(txt))
                left_sp = l["x0"] - col_left
                right_sp = col_right - l["x1"]
                is_centered = l["width"] < col_w * 0.72 and abs(left_sp - right_sp) < 16 and left_sp > 24

                cross_page = (not same_pc) and (prev_ended or is_indented or is_dialogue or prev_short)
                font_shift = abs(l["size"] - prev_line["size"]) >= 1.6

                if (
                    has_vgap
                    or is_dialogue
                    or is_centered
                    or cross_page
                    or font_shift
                    or (is_indented and (prev_ended or prev_short or indent_delta >= median_font * 1.0))
                    or (prev_short and prev_ended)
                    or (prev_line["width"] < col_w * 0.55 and prev_ended)
                ):
                    start_new = True

            if start_new:
                flush()
                cur_text = txt
                cur_html = l["html"]
            elif opts.get("dehyphen", True) and re.search(r"[a-záéíóúâêôãõç]-$", cur_text, re.I) and re.match(r"^[a-záéíóúâêôãõç]", txt, re.I):
                cur_text = cur_text[:-1] + txt
                cur_html = re.sub(r"-(</(?:strong|em)>)*$", r"\1", cur_html) + txt
            else:
                cur_text += " " + txt
                cur_html += " " + l["html"]

            prev_line = l
    flush()

    title = opts.get("title") or (meta.get("title") or "").strip() or Path(filename).stem
    author = opts.get("author") or (meta.get("author") or "").strip() or "Autor Desconhecido"
    chapters = _split_html_chapters("\n".join(html_parts), title, opts.get("splitChapters", True))
    doc.close()

    return {
        "title": title,
        "author": author,
        "lang": opts.get("lang") or "pt-BR",
        "chapters": chapters,
        "images": all_images,
        "cover_image": cover_image,
        "cover_bytes": cover_bytes,
    }


def parse_epub(data: bytes, filename: str, opts: dict) -> dict:
    with zipfile.ZipFile(io.BytesIO(data), "r") as zf:
        opf_path = None
        if "META-INF/container.xml" in zf.namelist():
            c_soup = BeautifulSoup(zf.read("META-INF/container.xml"), "xml")
            rf = c_soup.find("rootfile")
            if rf:
                opf_path = rf.get("full-path")
        if not opf_path:
            opf_path = next((n for n in zf.namelist() if n.endswith(".opf")), None)
        if not opf_path:
            raise ValueError("EPUB inválido: arquivo .opf não encontrado.")

        base_dir = opf_path.rsplit("/", 1)[0] + "/" if "/" in opf_path else ""
        opf = BeautifulSoup(zf.read(opf_path), "xml")

        def meta_txt(tag):
            el = opf.find(tag)
            return el.get_text(strip=True) if el else ""

        title = opts.get("title") or meta_txt("title") or Path(filename).stem
        author = opts.get("author") or meta_txt("creator") or "Autor Desconhecido"
        lang = opts.get("lang") or meta_txt("language") or "pt-BR"

        manifest = {}
        images = []
        img_href_map = {}
        for it in opf.find_all("item"):
            iid, href, mtype = it.get("id"), it.get("href") or "", it.get("media-type") or ""
            if iid:
                manifest[iid] = href
            if mtype.startswith("image/") and "svg" not in mtype:
                target = base_dir + href if (base_dir + href) in zf.namelist() else href
                if target in zf.namelist():
                    idx = len(images) + 1
                    ext = "png" if "png" in mtype else "jpg"
                    im_obj = {
                        "id": f"img_{idx:03d}",
                        "filename": f"img_{idx:03d}.{ext}",
                        "mime": mtype,
                        "bytes": zf.read(target),
                        "alt": f"Imagem {idx}",
                    }
                    images.append(im_obj)
                    img_href_map[href] = im_obj
                    img_href_map[href.split("/")[-1]] = im_obj

        spine_ids = [ir.get("idref") for ir in opf.find_all("itemref") if ir.get("idref")]
        chapters = []
        for sid in spine_ids:
            href = manifest.get(sid)
            if not href:
                continue
            target = base_dir + href if (base_dir + href) in zf.namelist() else href
            if target not in zf.namelist():
                continue
            soup = BeautifulSoup(zf.read(target).decode("utf-8", errors="replace"), "html.parser")
            body = soup.body or soup
            for s in body.find_all(["script", "style"]):
                s.decompose()
            for im in body.find_all("img"):
                src = (im.get("src") or "").lstrip("./")
                found = img_href_map.get(src) or img_href_map.get(src.split("/")[-1])
                if found:
                    im["src"] = f"images/{found['filename']}"
                    im["data-img-id"] = found["id"]
            if not body.get_text(strip=True) and not body.find("img"):
                continue
            h = body.find(["h1", "h2", "h3"])
            chapters.append({
                "id": f"chap_{len(chapters)+1}",
                "title": h.get_text(strip=True) if h else f"Seção {len(chapters)+1}",
                "html": body.decode_contents(),
            })

    return {"title": title, "author": author, "lang": lang, "chapters": chapters, "images": images, "cover_image": None, "cover_bytes": None}


def parse_docx(data: bytes, filename: str, opts: dict) -> dict:
    doc = DocxDocument(io.BytesIO(data))
    parts = []
    for p in doc.paragraphs:
        txt = p.text.strip()
        if not txt:
            continue
        style = (p.style.name or "").lower()
        if "heading 1" in style or "título 1" in style:
            parts.append(f"<h1>{html_lib.escape(txt)}</h1>")
        elif "heading 2" in style or "título 2" in style:
            parts.append(f"<h2>{html_lib.escape(txt)}</h2>")
        elif "heading" in style or "título" in style:
            parts.append(f"<h3>{html_lib.escape(txt)}</h3>")
        else:
            parts.append(f"<p>{html_lib.escape(txt)}</p>")

    title = opts.get("title") or (doc.core_properties.title or "").strip() or Path(filename).stem
    author = opts.get("author") or (doc.core_properties.author or "").strip() or "Autor Desconhecido"
    chapters = _split_html_chapters("\n".join(parts), title, opts.get("splitChapters", True))
    return {"title": title, "author": author, "lang": opts.get("lang") or "pt-BR", "chapters": chapters, "images": [], "cover_image": None, "cover_bytes": None}


def parse_text_like(data: bytes, filename: str, ext: str, opts: dict) -> dict:
    text = data.decode("utf-8", errors="replace")
    if ext in ("html", "htm"):
        soup = BeautifulSoup(text, "html.parser")
        body = soup.body or soup
        for s in body.find_all(["script", "style"]):
            s.decompose()
        full_html = body.decode_contents()
    elif ext in ("md", "markdown"):
        parts = []
        for l in text.splitlines():
            s = l.strip()
            if not s:
                continue
            if s.startswith("### "):
                parts.append(f"<h3>{html_lib.escape(s[4:])}</h3>")
            elif s.startswith("## "):
                parts.append(f"<h2>{html_lib.escape(s[3:])}</h2>")
            elif s.startswith("# "):
                parts.append(f"<h1>{html_lib.escape(s[2:])}</h1>")
            else:
                parts.append(f"<p>{html_lib.escape(s)}</p>")
        full_html = "\n".join(parts)
    else:
        blocks = [b.strip() for b in re.split(r"\n\s*\n+", text.replace("\r\n", "\n")) if b.strip()]
        full_html = "\n".join(f"<p>{html_lib.escape(re.sub(r'\s+', ' ', b))}</p>" for b in blocks)

    title = opts.get("title") or Path(filename).stem
    author = opts.get("author") or "Autor Desconhecido"
    chapters = _split_html_chapters(full_html, title, opts.get("splitChapters", True))
    return {"title": title, "author": author, "lang": opts.get("lang") or "pt-BR", "chapters": chapters, "images": [], "cover_image": None, "cover_bytes": None}


# =============================================================================
# 2. WRITERS: IR CANÔNICO -> EPUB 3 / DOCX / PDF / HTML / MD / TXT (COM IMAGENS)
# =============================================================================

def _to_strict_xhtml(fragment: str) -> str:
    soup = BeautifulSoup(fragment, "html.parser")
    for el in soup.find_all(True):
        for attr in ("data-img-id", "data-page"):
            if el.has_attr(attr):
                del el[attr]
    out = soup.decode(formatter="minimal")
    out = out.replace("&nbsp;", "&#160;")
    out = re.sub(r"<(br|hr|img)(\s[^>]*)?(?<!/)>", r"<\1\2/>", out)
    return out


def _inline_images_html(fragment: str, images: list) -> str:
    if not images:
        return fragment
    by_id = {im["id"]: im for im in images}
    by_path = {f"images/{im['filename']}": im for im in images}
    soup = BeautifulSoup(fragment, "html.parser")
    for im in soup.find_all("img"):
        found = by_id.get(im.get("data-img-id")) or by_path.get(im.get("src"))
        if found:
            b64 = base64.b64encode(found["bytes"]).decode("ascii")
            im["src"] = f"data:{found['mime']};base64,{b64}"
    return str(soup)


def write_epub(ir: dict, opts: dict) -> bytes:
    buf = io.BytesIO()
    uid = f"urn:uuid:{uuid.uuid4()}"
    modified = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    images = ir.get("images") or []

    with zipfile.ZipFile(buf, "w") as zf:
        zf.writestr("mimetype", "application/epub+zip", compress_type=zipfile.ZIP_STORED)
        zf.writestr("META-INF/container.xml", """<?xml version="1.0" encoding="UTF-8"?>
<container version="1.0" xmlns="urn:oasis:names:tc:opendocument:xmlns:container">
  <rootfiles><rootfile full-path="OEBPS/content.opf" media-type="application/oebps-package+xml"/></rootfiles>
</container>""", compress_type=zipfile.ZIP_DEFLATED)

        zf.writestr("OEBPS/styles.css", """body { font-family: Georgia, serif; line-height: 1.65; margin: 5%; color: #1d1d1f; }
h1, h2, h3 { line-height: 1.25; margin-top: 1.4em; margin-bottom: 0.6em; page-break-after: avoid; }
h1 { font-size: 1.6em; border-bottom: 1px solid #e5e5ea; padding-bottom: 0.25em; }
h1.centered { border-bottom: none; margin-bottom: 0.3em; }
h2.subtitle { font-size: 1.15em; font-weight: normal; font-style: italic; color: #555; margin-top: 0.2em; margin-bottom: 1.2em; }
.centered { text-align: center; text-indent: 0; }
p { margin: 0 0 0.75em 0; text-align: justify; text-indent: 1.25em; }
h1 + p, h2 + p, h3 + p, figure + p { text-indent: 0; }
figure.doc-figure { margin: 1.2em auto; text-align: center; page-break-inside: avoid; }
img { max-width: 100%; height: auto; }""", compress_type=zipfile.ZIP_DEFLATED)

        manifest = [
            '<item id="nav" href="nav.xhtml" media-type="application/xhtml+xml" properties="nav"/>',
            '<item id="ncx" href="toc.ncx" media-type="application/x-dtbncx+xml"/>',
            '<item id="css" href="styles.css" media-type="text/css"/>',
        ]
        spine, nav_links, ncx_pts = [], [], []

        cover_img = ir.get("cover_image")
        for im in images:
            zf.writestr(f"OEBPS/images/{im['filename']}", im["bytes"], compress_type=zipfile.ZIP_DEFLATED)
            props = ' properties="cover-image"' if (cover_img and im["id"] == cover_img["id"] and opts.get("genCover", True)) else ""
            manifest.append(f'<item id="{im["id"]}" href="images/{im["filename"]}" media-type="{im["mime"]}"{props}/>')

        if opts.get("genCover", True):
            if cover_img:
                zf.writestr("OEBPS/cover.xhtml", f"""<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="{html_lib.escape(ir['lang'])}">
<head><title>Capa</title><link rel="stylesheet" type="text/css" href="styles.css"/></head>
<body style="text-align:center"><img src="images/{cover_img['filename']}" alt="Capa"/></body></html>""", compress_type=zipfile.ZIP_DEFLATED)
                manifest.append('<item id="cover-page" href="cover.xhtml" media-type="application/xhtml+xml"/>')
                spine.append('<itemref idref="cover-page"/>')
            elif ir.get("cover_bytes"):
                zf.writestr("OEBPS/cover.jpg", ir["cover_bytes"], compress_type=zipfile.ZIP_DEFLATED)
                zf.writestr("OEBPS/cover.xhtml", f"""<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="{html_lib.escape(ir['lang'])}">
<head><title>Capa</title><link rel="stylesheet" type="text/css" href="styles.css"/></head>
<body style="text-align:center"><img src="cover.jpg" alt="Capa"/></body></html>""", compress_type=zipfile.ZIP_DEFLATED)
                manifest.append('<item id="cover-img" href="cover.jpg" media-type="image/jpeg" properties="cover-image"/>')
                manifest.append('<item id="cover-page" href="cover.xhtml" media-type="application/xhtml+xml"/>')
                spine.append('<itemref idref="cover-page"/>')

        for i, chap in enumerate(ir["chapters"], start=1):
            fname = f"chap_{i:03d}.xhtml"
            body = _to_strict_xhtml(chap["html"])
            xhtml = f"""<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="{html_lib.escape(ir['lang'])}">
<head><meta charset="UTF-8"/><title>{html_lib.escape(chap['title'])}</title><link rel="stylesheet" type="text/css" href="styles.css"/></head>
<body>\n{body}\n</body></html>"""
            zf.writestr(f"OEBPS/{fname}", xhtml, compress_type=zipfile.ZIP_DEFLATED)
            manifest.append(f'<item id="chap{i}" href="{fname}" media-type="application/xhtml+xml"/>')
            spine.append(f'<itemref idref="chap{i}"/>')
            nav_links.append(f'<li><a href="{fname}">{html_lib.escape(chap["title"])}</a></li>')
            ncx_pts.append(f'<navPoint id="np-{i}" playOrder="{i}"><navLabel><text>{html_lib.escape(chap["title"])}</text></navLabel><content src="{fname}"/></navPoint>')

        zf.writestr("OEBPS/nav.xhtml", f"""<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml" xmlns:epub="http://www.idpf.org/2007/ops" xml:lang="{html_lib.escape(ir['lang'])}">
<head><title>Sumário</title></head>
<body><nav epub:type="toc" id="toc"><h1>Sumário</h1><ol>{"".join(nav_links)}</ol></nav></body></html>""", compress_type=zipfile.ZIP_DEFLATED)

        zf.writestr("OEBPS/toc.ncx", f"""<?xml version="1.0" encoding="UTF-8"?>
<ncx xmlns="http://www.daisy.org/z3986/2005/ncx/" version="2005-1">
<head><meta name="dtb:uid" content="{uid}"/><meta name="dtb:depth" content="1"/></head>
<docTitle><text>{html_lib.escape(ir['title'])}</text></docTitle>
<navMap>{"".join(ncx_pts)}</navMap></ncx>""", compress_type=zipfile.ZIP_DEFLATED)

        zf.writestr("OEBPS/content.opf", f"""<?xml version="1.0" encoding="UTF-8"?>
<package xmlns="http://www.idpf.org/2007/opf" version="3.0" unique-identifier="bookid">
  <metadata xmlns:dc="http://purl.org/dc/elements/1.1/">
    <dc:identifier id="bookid">{uid}</dc:identifier>
    <dc:title>{html_lib.escape(ir['title'])}</dc:title>
    <dc:creator>{html_lib.escape(ir['author'])}</dc:creator>
    <dc:language>{html_lib.escape(ir['lang'])}</dc:language>
    <meta property="dcterms:modified">{modified}</meta>
  </metadata>
  <manifest>{"".join(manifest)}</manifest>
  <spine toc="ncx">{"".join(spine)}</spine>
</package>""", compress_type=zipfile.ZIP_DEFLATED)

    return buf.getvalue()


def write_docx(ir: dict, opts: dict) -> bytes:
    doc = DocxDocument()
    doc.core_properties.title = ir["title"]
    doc.core_properties.author = ir["author"]
    images = ir.get("images") or []
    by_id = {im["id"]: im for im in images}
    by_path = {f"images/{im['filename']}": im for im in images}

    for chap in ir["chapters"]:
        soup = BeautifulSoup(chap["html"], "html.parser")
        for el in soup.find_all(["h1", "h2", "h3", "p", "li", "img"]):
            if el.name == "img":
                found = by_id.get(el.get("data-img-id")) or by_path.get(el.get("src"))
                if found:
                    try:
                        doc.add_picture(io.BytesIO(found["bytes"]), width=Inches(5.5))
                    except Exception:
                        pass
                continue
            txt = el.get_text(" ", strip=True)
            if not txt:
                continue
            if el.name == "h1":
                doc.add_heading(txt, level=1)
            elif el.name == "h2":
                doc.add_heading(txt, level=2)
            elif el.name == "h3":
                doc.add_heading(txt, level=3)
            elif el.name == "li":
                doc.add_paragraph(txt, style="List Bullet")
            else:
                doc.add_paragraph(txt)
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def write_pdf(ir: dict, opts: dict) -> bytes:
    images = ir.get("images") or []
    html_body = "\n".join(_inline_images_html(c["html"], images) for c in ir["chapters"])
    full_html = f"""<html><head><style>
      body {{ font-family: serif; font-size: 11pt; line-height: 1.5; }}
      h1 {{ font-size: 18pt; color: #0071e3; margin-top: 16pt; }}
      h2 {{ font-size: 14pt; margin-top: 12pt; }}
      p {{ text-align: justify; margin-bottom: 8pt; }}
      img {{ max-width: 100%; }}
    </style></head><body><h1>{html_lib.escape(ir['title'])}</h1><p><em>{html_lib.escape(ir['author'])}</em></p><hr/>{html_body}</body></html>"""

    story = fitz.Story(html=full_html)
    buf = io.BytesIO()
    writer = fitz.DocumentWriter(buf)
    mediabox = fitz.paper_rect("a4")
    where = mediabox + (50, 54, -50, -54)
    more = 1
    while more:
        dev = writer.begin_page(mediabox)
        more, _ = story.place(where)
        story.draw(dev)
        writer.end_page()
    writer.close()
    return buf.getvalue()


def write_html(ir: dict, opts: dict) -> bytes:
    images = ir.get("images") or []
    body = "\n<hr/>\n".join(_inline_images_html(c["html"], images) for c in ir["chapters"])
    out = f"""<!DOCTYPE html><html lang="{html_lib.escape(ir['lang'])}"><head><meta charset="UTF-8"/><title>{html_lib.escape(ir['title'])}</title>
<style>body{{max-width:740px;margin:48px auto;padding:0 20px;font-family:Georgia,serif;line-height:1.75;color:#1d1d1f}}p{{text-align:justify;text-indent:1.25em}}h1+p,h2+p,figure+p{{text-indent:0}}figure{{text-align:center;margin:20px 0}}img{{max-width:100%;height:auto;border-radius:10px}}</style></head>
<body><h1>{html_lib.escape(ir['title'])}</h1><p style="text-indent:0"><em>{html_lib.escape(ir['author'])}</em></p>{body}</body></html>"""
    return out.encode("utf-8")


def write_md(ir: dict, opts: dict) -> bytes:
    lines = [f"# {ir['title']}", f"*{ir['author']}*", ""]
    for chap in ir["chapters"]:
        soup = BeautifulSoup(chap["html"], "html.parser")
        for el in soup.find_all(["h1", "h2", "h3", "p", "li", "img"]):
            if el.name == "img":
                lines.extend([f"![{el.get('alt', 'Imagem')}]({el.get('src', '')})", ""])
                continue
            txt = el.get_text(" ", strip=True)
            if not txt:
                continue
            if el.name == "h1":
                lines.extend(["", f"# {txt}", ""])
            elif el.name == "h2":
                lines.extend(["", f"## {txt}", ""])
            elif el.name == "h3":
                lines.extend(["", f"### {txt}", ""])
            elif el.name == "li":
                lines.append(f"- {txt}")
            else:
                lines.extend([txt, ""])
    return "\n".join(lines).encode("utf-8")


def write_txt(ir: dict, opts: dict) -> bytes:
    lines = [ir["title"].upper(), ir["author"], "=" * 50, ""]
    for chap in ir["chapters"]:
        soup = BeautifulSoup(chap["html"], "html.parser")
        for el in soup.find_all(["h1", "h2", "h3", "p", "li"]):
            txt = el.get_text(" ", strip=True)
            if txt:
                lines.extend([txt, ""])
    return "\n".join(lines).encode("utf-8")


def convert_bytes(data: bytes, filename: str, target_fmt: str, opts: dict | None = None) -> bytes:
    opts = opts or {}
    src_ext = Path(filename).suffix.lstrip(".").lower()
    if src_ext == "pdf":
        ir = parse_pdf(data, filename, opts)
    elif src_ext == "epub":
        ir = parse_epub(data, filename, opts)
    elif src_ext == "docx":
        ir = parse_docx(data, filename, opts)
    else:
        ir = parse_text_like(data, filename, src_ext, opts)

    dst = target_fmt.lower().lstrip(".")
    if dst == "epub":
        return write_epub(ir, opts)
    if dst == "docx":
        return write_docx(ir, opts)
    if dst == "pdf":
        return write_pdf(ir, opts)
    if dst in ("html", "htm"):
        return write_html(ir, opts)
    if dst in ("md", "markdown"):
        return write_md(ir, opts)
    return write_txt(ir, opts)


# =============================================================================
# 3. SERVIDOR LOCAL FLASK OPCIONAL (CONECTA AUTOMATICAMENTE AO index.html)
# =============================================================================

def run_server(port: int = 5055):
    from flask import Flask, jsonify, request, send_file

    app = Flask(__name__)

    @app.after_request
    def add_cors(resp):
        resp.headers["Access-Control-Allow-Origin"] = "*"
        resp.headers["Access-Control-Allow-Headers"] = "Content-Type"
        resp.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS"
        return resp

    @app.route("/api/health")
    def health():
        return jsonify({"status": "ok", "engine": "PyMuPDF + python-docx", "pymupdf_version": fitz.version[0]})

    @app.route("/api/convert", methods=["POST", "OPTIONS"])
    def api_convert():
        if request.method == "OPTIONS":
            return ("", 204)
        f = request.files.get("file")
        if not f:
            return ("Nenhum arquivo enviado", 400)
        target = request.form.get("target_format", "epub")
        opts = json.loads(request.form.get("options", "{}"))
        out_bytes = convert_bytes(f.read(), f.filename or "documento.pdf", target, opts)
        out_name = f"{Path(f.filename or 'documento').stem}.{target}"
        return send_file(io.BytesIO(out_bytes), as_attachment=True, download_name=out_name)

    @app.route("/")
    def index():
        idx = Path(__file__).parent / "index.html"
        return send_file(idx) if idx.exists() else ("index.html não encontrado", 404)

    print(f"DocConvert Local Engine rodando em http://127.0.0.1:{port}")
    app.run(host="127.0.0.1", port=port, debug=False)


def main():
    parser = argparse.ArgumentParser(description="DocConvert Local Engine (PyMuPDF + python-docx)")
    parser.add_argument("input", nargs="?", help="Arquivo de entrada (pdf, epub, docx, md, html, txt)")
    parser.add_argument("output", nargs="?", help="Arquivo de saída (epub, pdf, docx, md, html, txt)")
    parser.add_argument("--title", default="", help="Título do documento")
    parser.add_argument("--author", default="", help="Autor do documento")
    parser.add_argument("--serve", action="store_true", help="Inicia servidor local na porta 5055 integrado ao index.html")
    parser.add_argument("--port", type=int, default=5055, help="Porta do servidor local (padrão: 5055)")
    args = parser.parse_args()

    if args.serve:
        run_server(args.port)
        return

    if not args.input or not args.output:
        parser.print_help()
        sys.exit(1)

    in_path = Path(args.input)
    out_path = Path(args.output)
    target = out_path.suffix.lstrip(".").lower()
    opts = {"title": args.title, "author": args.author}
    out_bytes = convert_bytes(in_path.read_bytes(), in_path.name, target, opts)
    out_path.write_bytes(out_bytes)
    print(f"Convertido com sucesso: {in_path.name} -> {out_path.name} ({len(out_bytes):,} bytes)")


if __name__ == "__main__":
    main()
