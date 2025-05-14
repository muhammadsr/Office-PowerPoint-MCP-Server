#!/usr/bin/env python
"""
MCP Server for PowerPoint manipulation using python-pptx.
"""
import base64
from typing import Dict, List, Optional

import mcp.types as types
from mcp.server.fastmcp import FastMCP

from presentation import Presentation

SESSION_SLIDE_INDEX = 0

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
def list_layouts() -> Dict:
    """
    Return the available slide layouts (index and name)
    so the LLM can choose one.
    """
    return get_session().get_layouts()

@app.tool()
def create_presentation(layout_index: Optional[int] = None) -> Dict:
    """
    Create a new presentation, using the given layout index (or blank if None).
    """
    return get_session().add_layout_index(layout_index)

@app.tool()
def get_slide_info() -> Dict:
    return get_session().get_slide_info(SESSION_SLIDE_INDEX)


@app.tool()
def save_presentation(file_path: str) -> Dict:
    return get_session().save_presentation(file_path)


@app.tool()
def populate_placeholder(
        placeholder_idx: int,
        text: str
) -> Dict:
    return get_session().populate_placeholder(SESSION_SLIDE_INDEX, placeholder_idx, text)


@app.tool()
def add_bullet_points(
        placeholder_idx: int,
        bullet_points: List[str]
) -> Dict:
    return get_session().add_bullet_points(SESSION_SLIDE_INDEX, placeholder_idx, bullet_points)


@app.tool()
def add_textbox(
        left_inches: float,
        top_inches: float,
        width_inches: float,
        height_inches: float,
        text: str,
        font_size: Optional[int] = None,
        font_name: Optional[str] = None,
        bold: Optional[bool] = None,
        italic: Optional[bool] = None,
        color: Optional[List[int]] = None,
        alignment: Optional[str] = None
) -> Dict:
    return get_session().add_textbox(
        0,
        left=left_inches,
        top=top_inches,
        width=width_inches,
        height=height_inches,
        text=text,
        font_size=font_size,
        font_name=font_name,
        bold=bold,
        italic=italic,
        color=color,
        alignment=alignment
    )


@app.tool()
def add_image(
        image_path: str,
        left: float,
        top: float,
        width: Optional[float] = None,
        height: Optional[float] = None
) -> Dict:
    return get_session().add_image(SESSION_SLIDE_INDEX, image_path, left, top, width, height)


@app.tool()
def add_image_from_base64(
        base64_string: str,
        left: float,
        top: float,
        width: Optional[float] = None,
        height: Optional[float] = None
) -> Dict:
    return get_session().add_image_from_base64(SESSION_SLIDE_INDEX, base64_string, left, top, width, height)


@app.tool()
def add_table(
        rows: int,
        cols: int,
        left: float,
        top: float,
        width: float,
        height: float,
        data: Optional[List[List[str]]] = None
) -> Dict:
    return get_session().add_table(SESSION_SLIDE_INDEX, rows, cols, left, top, width, height, data)


@app.tool()
def format_table_cell(
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
        0,
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
        shape_type: str,
        left: float,
        top: float,
        width: float,
        height: float,
        fill_color: Optional[List[int]] = None,
        line_color: Optional[List[int]] = None,
        line_width: Optional[float] = None
) -> Dict:
    return get_session().add_shape(SESSION_SLIDE_INDEX,
                                   shape_type, left, top, width, height, fill_color, line_color, line_width
                                   )


@app.tool()
def add_chart(
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
    return get_session().add_chart(SESSION_SLIDE_INDEX,
                                   chart_type, left, top, width, height,
                                   categories, series_names, series_values,
                                   has_legend, legend_position, has_data_labels, title
                                   )


@app.tool()
def get_slide_image() -> List[types.ImageContent] | Dict:
    """
    Return the current slide as an MCP image content for inline rendering.
    """
    # 1) Render slide to PNG bytes
    try:
        png_bytes = get_session().get_slide_image()

        # 2) Encode to base64
        b64 = base64.b64encode(png_bytes).decode('ascii')
    except Exception as e:
        return {
            "error": f"Failed to get slide image: {str(e)}"
        }

    # 3) Return as MCP ImageContent list
    return [
        types.ImageContent(
            type="image",
            data=b64,
            mimeType="image/png"
        )
    ]

@app.tool()
def move_element(
        element_or_shape_index: int,
        left: float,
        top: float
) -> Dict:
    return get_session().move_element(element_or_shape_index, left, top)


@app.tool()
def remove_element(shape_index: int) -> Dict:
    """Remove the shape at the given index from the current slide."""
    return get_session().remove_element(SESSION_SLIDE_INDEX, shape_index)



# ---- Main Execution ----
def main():
    # Run the FastMCP server
    app.run(transport='stdio')


if __name__ == "__main__":
    # entrypoint()
    main()
