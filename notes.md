# Thinking Out Loud
* Not sure what "gets activated" means. I am assuming that if we trigger the MCP to create a presentation
* I kept the ID for now which goes to the LLM. This can be leveraged if we decided to create multiple presentations per session
* I kept slide index to support multi slides but removed presentation_id.
* Ran into an issue with the `aspose` library and I had to install `brew install mono-libgdiplus`
  * Had to put this in the `env` in the config file `"DYLD_FALLBACK_LIBRARY_PATH": "/usr/local/lib:/opt/homebrew/lib"`
* To preserve the variations from SVG to PNG, I am using `inkscape` which is a system tool. For more info https://inkscape-manuals.readthedocs.io/en/latest/index.html. In case, we want to avoid the system tool, we can use `cairosvg` as a fallback. I did have an issue preserving the font style (PPT → SVG) with that library tho.
* I had to change the top,left in add_text_box to include "inches" because the LLM was thinking it's pixels.
* I am using `types.ImageContent` to render the image in the request window. However, I am not sure how to render it in the chat.
* I was adding a blank slide when I first created the presentation. However, I found that it is much more effective if we let the LLM `list_layouts` and decide the layout template the for the slide (`0 → Title Slide, 1 → Title and Content ...etc`).
* I had to create a new presentation in the session if the LLM is asking for it. This way, we can avoid using the same presentation across chats. 

# Limitations
* `FontsLoader.load_external_fonts` is mainly for macOS. If we need to handle multiple OS, we will need to do further work.
* There is a delay when we `get_slide_image`. I will need to investigate that further.
* In some cases, the PNG we display in the chat is not a 100% perfect match of the PPT.
* If we are working on an unsaved presentation in one chat, and we start another chat, the previous presentation will get destroyed.
* Watermark remove logic is a hack and relies on the watermark including the text "Aspose". If that changes, the watermark remove will fail.

# Test Coverage (Partial `/tests`)
- **Init & Shape Tools**: `Presentation()` creates exactly one blank slide; `add_shape` + `move_element` place shapes as expected. Also, test `get_slide_info` (placeholders/textboxes)
- **SVG & PNG Watermark**: verify SVG export strips Aspose watermark and PNG conversion yields a non-empty image.
- **Next**: We are missing coverage for things like:
  - `get_slide_image` due to the `DllNotFoundException` issue
  - `add_table` / `format_table_cell` behavior
  - `add_image` / `add_image_from_base64` insertion
  - `populate_placeholder` and text formatting

## Improvements
- Promote parameters to typed Pydantic models
- Centralized error / warning schema (i.e. `Result(ok: bool, msg: str, data: Any = None, warnings: list[str] = [])`)
- If we want to maintain presentations between sessions, we might want to introduce sessions. One approach is let the llm `create_session`.


