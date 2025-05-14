#!/usr/bin/env python3
import os
import sys
import subprocess
from lxml import etree  # pip install lxml
import aspose.slides as slides
from aspose.slides.export import SVGOptions

def strip_watermark_from_file(svg_path: str):
    """Load the SVG, remove any <text>…Aspose…</text> elements, overwrite the file."""
    parser = etree.XMLParser(ns_clean=True, recover=True)
    tree = etree.parse(svg_path, parser)
    ns = tree.getroot().nsmap.copy()
    if None in ns:
        ns['svg'] = ns.pop(None)

    for txt in tree.xpath('//svg:text[contains(normalize-space(string(.)), "Aspose")]', namespaces=ns):
        parent = txt.getparent()
        if parent is not None:
            parent.remove(txt)

    # write back (including XML decl & DOCTYPE if you need them—ElementTree will omit by default)
    tree.write(svg_path, xml_declaration=True, encoding='utf-8')

def svg_to_png_with_inkscape(svg_path: str, png_path: str, dpi: int = 300):
    subprocess.run([
        "inkscape",
        svg_path,
        "--export-type=png",
        "--export-filename", png_path,
        "--export-dpi", str(dpi)
    ], check=True)

def pptx_to_png_via_svg(input_pptx: str, svg_dir: str, png_dir: str):
    os.makedirs(svg_dir, exist_ok=True)
    os.makedirs(png_dir, exist_ok=True)

    # load fonts so Aspose can embed them
    slides.FontsLoader.load_external_fonts([
        "/Library/Fonts",
        "/System/Library/Fonts",
        os.path.expanduser("~/Library/Fonts")
    ])

    # high-fidelity SVG options
    opts = SVGOptions.wysiwyg
    opts.vectorize_text = False
    opts.use_frame_size = True
    opts.metafile_rasterization_dpi = 300
    opts.external_fonts_handling = slides.export.SvgExternalFontsHandling.EMBED

    pres = slides.Presentation(input_pptx)
    for idx, slide in enumerate(pres.slides, start=1):
        svg_path = os.path.join(svg_dir, f"slide-{idx}.svg")
        # 1) write raw SVG to file
        with open(svg_path, "wb") as f_svg:
            slide.write_as_svg(f_svg, opts)

        # 2) strip the Aspose watermark in-place
        strip_watermark_from_file(svg_path)

        # 3) render cleaned SVG → PNG
        png_path = os.path.join(png_dir, f"slide-{idx}.png")
        svg_to_png_with_inkscape(svg_path, png_path)

        print(f"[{idx}] → SVG: {svg_path}")
        print(f"[{idx}] → PNG: {png_path}")


if __name__ == "__main__":
    if len(sys.argv) != 4:
        print("Usage: slide_export.py input.pptx svg_out_dir png_out_dir")
        sys.exit(1)

    pptx, svg_out, png_out = sys.argv[1], sys.argv[2], sys.argv[3]
    pptx_to_png_via_svg(pptx, svg_out, png_out)
