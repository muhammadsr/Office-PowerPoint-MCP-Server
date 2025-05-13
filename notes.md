* Not sure what "gets activated" means. I am assuming that if we trigger the MCP to create a presentation
* I kept the ID for now which goes to the LLM. This can be leveraged if we decided to create mutliple presentations per session
* I kept slide index to support multi slides but removed presentation_id.
* The first time around, the LLM must get the slide info to figure out what's on the slide.
* Ran into an issue where I had to install `brew install mono-libgdiplus`
  * Had to put this in the `env` in the config file `"DYLD_FALLBACK_LIBRARY_PATH": "/usr/local/lib:/opt/homebrew/lib"`
* To preserve the variations from SVG to PNG, I am using `inkscape` which is a system tool. For more info https://inkscape-manuals.readthedocs.io/en/latest/index.html