# Thinking Out Loud
* Not sure what "gets activated" means. I am assuming that if we trigger the MCP to create a presentation
* I kept the ID for now which goes to the LLM. This can be leveraged if we decided to create multiple presentations per session
* I kept slide index to support multi slides but removed presentation_id.
* Ran into an issue with the `aspose` library and I had to install `brew install mono-libgdiplus`
  * Had to put this in the `env` in the config file `"DYLD_FALLBACK_LIBRARY_PATH": "/usr/local/lib:/opt/homebrew/lib"`
* To preserve the variations from SVG to PNG, I am using `inkscape` which is a system tool. For more info https://inkscape-manuals.readthedocs.io/en/latest/index.html. In case, we want to avoid the system tool, we can use `cairosvg` as a fallback. I did have an issue preserving the font style (PPT → SVG) with that library tho.
* I had to change the top,left in add_text_box to include "inches" because the LLM was thinking it's pixels.
* I am using `types.ImageContent` to render the image in the request window. However, I am not sure how to render it in the chat.

# Test Coverage (Partial)

- **Init & Shape Tools**: `Presentation()` creates exactly one blank slide; `add_shape` + `move_element` place shapes as expected. Also, test `get_slide_info` (placeholders/textboxes)
- **SVG & PNG Watermark**: verify SVG export strips Aspose watermark and PNG conversion yields non-empty image.
- **Next**: We are missing coverage for things like:
  - `get_slide_image` due to the `DllNotFoundException` issue
  - `add_table` / `format_table_cell` behavior
  - `add_image` / `add_image_from_base64` insertion
  - `populate_placeholder` and text formatting  