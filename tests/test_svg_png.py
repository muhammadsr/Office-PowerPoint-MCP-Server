import pytest

from presentation import remove_aspose_watermark, svg_to_png

SVG_HEADER = """<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<!DOCTYPE svg PUBLIC "-//W3C//DTD SVG 1.1//EN" \
"http://www.w3.org/Graphics/SVG/1.1/DTD/svg11.dtd">
"""

@pytest.fixture
def dummy_svg_with_watermark():
    # note: we include a real <text>…Aspose…</text> inside
    body = """
    <svg xmlns="http://www.w3.org/2000/svg" width="200" height="100">
      <rect width="200" height="100" fill="red"/>
      <text x="10" y="20">Created with Aspose.Slides for Python</text>
      <circle cx="50" cy="50" r="10" fill="blue"/>
    </svg>
    """
    return SVG_HEADER + body

def test_remove_aspose_watermark(dummy_svg_with_watermark):
    clean = remove_aspose_watermark(dummy_svg_with_watermark)
    # the Aspose text node must be gone
    assert "Aspose" not in clean
    # other elements remain
    assert "<rect" in clean
    assert "<circle" in clean

def test_svg_to_png_minimal(tmp_path):
    # a minimal valid SVG (no Aspose watermark)
    svg = SVG_HEADER + """
    <svg xmlns="http://www.w3.org/2000/svg" width="50" height="50">
      <rect width="50" height="50" fill="green"/>
    </svg>
    """
    png_bytes = svg_to_png(svg, dpi=72)
    # PNGs always begin with this eight‐byte signature:
    assert png_bytes.startswith(b"\x89PNG\r\n\x1a\n")

    # write out so we can eyeball if needed
    out = tmp_path / "test_min.png"
    out.write_bytes(png_bytes)
    assert out.exists()

