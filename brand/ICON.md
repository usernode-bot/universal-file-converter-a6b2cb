# Universal File Converter — brand icons

**Glyph:** one bold document page with a folded top-right corner, and two
arrows chasing each other in a circle on the page. That means "this file
becomes another kind of file". Flat, no text.

**Style:** follows the Homeroom brand framework (https://baked.design/brand/homeroom/):
one mark on a full-bleed square tile, in an allowed brand pairing. The brand
page and its "brand file for AI" download were password protected when this
was made, so the colours, pairings and shape rules written in the request
were used as the rules.

**Colours:** brand Cream `#FFFEEA`, Black `#000000`, Yellow `#FFEE6F`.
Anti-aliased edge pixels blend these and are not counted as new colours.

| File | Size | Pairing |
| --- | --- | --- |
| `icon.svg` / `icon.png` | 1024×1024 | Cream tile, black page, cream arrows |
| `featured-light.svg` / `featured-light.png` | 1600×900 | Cream ground, black page, cream arrows, yellow page for the converted file |
| `featured-dark.svg` / `featured-dark.png` | 1600×900 | Black ground, cream page, black arrows, yellow page |

The app icon is drawn with square corners: Homeroom rounds the tile itself.
The platform fills and crops `icon.png` into the tile, so the mark keeps at
least 160 px of margin on every side.

**Rendering** (one-off, not committed): the PNGs were rendered from the SVGs
with the sandbox's headless Chromium —
`headless_shell --headless --window-size=1024,1024 --screenshot=icon.png icon.svg`
(and 1600,900 for the featured pair). The SVGs use only plain shapes, so any
renderer gives the same result.

**Wiring:** `dapp.json` sets `icon.image` to `brand/icon.png`, read at deploy
time (PNG, under 256 KB); `icon.emoji` 📄 stays as the fallback if the image
can't load. The featured illustration has no manifest field on the platform
yet, so `featured-light.png` / `featured-dark.png` are delivered here in the
repo for the directory's featured card.