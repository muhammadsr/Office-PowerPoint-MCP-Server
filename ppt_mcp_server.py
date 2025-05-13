#!/usr/bin/env python
"""
MCP Server for PowerPoint manipulation using python-pptx.
"""
from typing import Dict, List, Optional

from mcp.server.fastmcp import FastMCP

from presentation import Presentation

# Initialize the FastMCP server
app = FastMCP(
    name="ppt-mcp-server",
    description="MCP Server for PowerPoint manipulation using python-pptx",
    version="1.0.0"
)

# Singleton session
_session: Optional[Presentation] = None


def get_session() -> Presentation:
    global _session
    if _session is None:
        _session = Presentation()
    return _session


# ---- Needed for multi presentations session ----
# @app.tool()
# def get_presentation_info() -> Dict:
#     return get_session().get_presentation_info()


# ---- Tools ----
@app.tool()
def get_slide_info(slide_index: int) -> Dict:
    return get_session().get_slide_info(slide_index)


@app.tool()
def save_presentation(file_path: str) -> Dict:
    return get_session().save_presentation(file_path)


@app.tool()
def populate_placeholder(
        slide_index: int,
        placeholder_idx: int,
        text: str
) -> Dict:
    return get_session().populate_placeholder(slide_index, placeholder_idx, text)


@app.tool()
def add_bullet_points(
        slide_index: int,
        placeholder_idx: int,
        bullet_points: List[str]
) -> Dict:
    return get_session().add_bullet_points(slide_index, placeholder_idx, bullet_points)


@app.tool()
def add_textbox(
        slide_index: int,
        left: float,
        top: float,
        width: float,
        height: float,
        text: str,
        font_size: Optional[int] = None,
        font_name: Optional[str] = None,
        bold: Optional[bool] = None,
        italic: Optional[bool] = None,
        color: Optional[List[int]] = None,
        alignment: Optional[str] = None
) -> Dict:
    return get_session().add_textbox(
        slide_index,
        left, top, width, height, text,
        font_size=font_size,
        font_name=font_name,
        bold=bold,
        italic=italic,
        color=color,
        alignment=alignment
    )


@app.tool()
def add_image(
        slide_index: int,
        image_path: str,
        left: float,
        top: float,
        width: Optional[float] = None,
        height: Optional[float] = None
) -> Dict:
    return get_session().add_image(slide_index, image_path, left, top, width, height)


@app.tool()
def add_image_from_base64(
        slide_index: int,
        base64_string: str,
        left: float,
        top: float,
        width: Optional[float] = None,
        height: Optional[float] = None
) -> Dict:
    return get_session().add_image_from_base64(slide_index, base64_string, left, top, width, height)


@app.tool()
def add_table(
        slide_index: int,
        rows: int,
        cols: int,
        left: float,
        top: float,
        width: float,
        height: float,
        data: Optional[List[List[str]]] = None
) -> Dict:
    return get_session().add_table(slide_index, rows, cols, left, top, width, height, data)


@app.tool()
def format_table_cell(
        slide_index: int,
        shape_index: int,
        row: int,
        col: int,
        font_size: Optional[int] = None,
        font_name: Optional[str] = None,
        bold: Optional[bool] = None,
        italic: Optional[bool] = None,
        color: Optional[List[int]] = None,
        bg_color: Optional[List[int]] = None,
        alignment: Optional[str] = None,
        vertical_alignment: Optional[str] = None
) -> Dict:
    return get_session().format_table_cell(
        slide_index,
        shape_index, row, col,
        font_size=font_size,
        font_name=font_name,
        bold=bold,
        italic=italic,
        color=color,
        bg_color=bg_color,
        alignment=alignment,
        vertical_alignment=vertical_alignment
    )


@app.tool()
def add_shape(
        slide_index: int,
        shape_type: str,
        left: float,
        top: float,
        width: float,
        height: float,
        fill_color: Optional[List[int]] = None,
        line_color: Optional[List[int]] = None,
        line_width: Optional[float] = None
) -> Dict:
    return get_session().add_shape(slide_index,
                                   shape_type, left, top, width, height, fill_color, line_color, line_width
                                   )


@app.tool()
def add_chart(
        slide_index: int,
        chart_type: str,
        left: float,
        top: float,
        width: float,
        height: float,
        categories: List[str],
        series_names: List[str],
        series_values: List[List[float]],
        has_legend: bool = True,
        legend_position: str = "right",
        has_data_labels: bool = False,
        title: Optional[str] = None
) -> Dict:
    return get_session().add_chart(slide_index,
                                   chart_type, left, top, width, height,
                                   categories, series_names, series_values,
                                   has_legend, legend_position, has_data_labels, title
                                   )


# @app.tool()
# def get_slide_svg() -> Dict:
#     svg = get_session().get_slide_svg()
#     return {"svg": svg}
#
#
# @app.tool()
# def get_slide_png() -> Dict:
#     png = get_session().get_slide_png()
#     return {"png_base64": base64.b64encode(png).decode()}


# ---- Main Execution ----
def main():
    # Start the MCP server (stdio for Claude Desktop, or change transport)
    app.run(transport="stdio")


if __name__ == "__main__":
    main()
