#!/usr/bin/env python3
import os
import io
import sys
import subprocess

import aspose.slides as slides
from aspose.slides.export import SVGOptions
from lxml import etree  # pip install lxml

def remove_aspose_watermark(svg_xml: str) -> str:
    """Strip any <text>…Aspose…</text> elements via XML parsing."""
    parser = etree.XMLParser(ns_clean=True, recover=True)
    tree = etree.fromstring(svg_xml.encode("utf-8"), parser=parser)
    ns = tree.nsmap.copy()
    if None in ns:
        ns["svg"] = ns.pop(None)
    for txt in tree.xpath('//svg:text[contains(normalize-space(string(.)), "Aspose")]', namespaces=ns):
        if txt.getparent() is not None:
            txt.getparent().remove(txt)
    decl = '<?xml version="1.0" encoding="utf-8" standalone="yes"?>\n'
    doctype = '<!DOCTYPE svg PUBLIC "-//W3C//DTD SVG 1.1//EN" ' \
              '"http://www.w3.org/Graphics/SVG/1.1/DTD/svg11.dtd">\n'
    body = etree.tostring(tree, encoding="utf-8").decode("utf-8")
    return decl + doctype + body

def svg_to_png_with_inkscape(svg_path: str, png_path: str, dpi: int = 300):
    """
    Render an SVG file to PNG using Inkscape CLI at the given DPI,
    preserving all font styles via the system’s font engine.
    """
    subprocess.run([
        "inkscape",
        svg_path,
        "--export-type=png",
        "--export-filename", png_path,
        "--export-dpi", str(dpi)
    ], check=True)

def pptx_to_png_via_svg(input_pptx: str, svg_dir: str, png_dir: str):
    # 1) Prepare output dirs
    os.makedirs(svg_dir, exist_ok=True)
    os.makedirs(png_dir, exist_ok=True)

    # 2) Make sure Aspose can embed your macOS fonts
    slides.FontsLoader.load_external_fonts([
        "/Library/Fonts",
        "/System/Library/Fonts",
        os.path.expanduser("~/Library/Fonts")
    ])

    # 3) Configure high-fidelity SVG export
    opts = SVGOptions.wysiwyg
    opts.vectorize_text = False            # keep <text> so we can remove Aspose watermark
    opts.use_frame_size = True
    opts.metafile_rasterization_dpi = 300
    opts.external_fonts_handling = slides.export.SvgExternalFontsHandling.EMBED

    # 4) Loop through slides
    with slides.Presentation(input_pptx) as pres:
        total = pres.slides.length
        for i, slide in enumerate(pres.slides, start=1):
            # a) export to SVG
            buf = io.BytesIO()
            slide.write_as_svg(buf, opts)
            raw_svg = buf.getvalue().decode("utf-8")

            # b) remove watermark
            clean_svg = remove_aspose_watermark(raw_svg)

            # c) save SVG
            svg_path = os.path.join(svg_dir, f"slide_{i}.svg")
            with open(svg_path, "w", encoding="utf-8") as f:
                f.write(clean_svg)

            # d) render to PNG with Inkscape
            png_path = os.path.join(png_dir, f"slide_{i}.png")
            svg_to_png_with_inkscape(svg_path, png_path, dpi=300)

            print(f"[{i}/{total}] → SVG: {svg_path}")
            print(f"[{i}/{total}] → PNG: {png_path}")

if __name__ == "__main__":
    if len(sys.argv) == 4:
        pptx, svg_out, png_out = sys.argv[1:]
    else:
        raise Exception("3 params are not provided")

    # Ensure Inkscape is installed:
    #   brew install --cask inkscape    (macOS)
    pptx_to_png_via_svg(pptx, svg_out, png_out)
