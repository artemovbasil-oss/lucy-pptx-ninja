# lucy-pptx-ninja
Turn Figma frames into clean, editable PowerPoint slides. Fast. Precise. No fuss.
***
https://github.com/user-attachments/assets/a6b4aff3-7923-4812-8a67-85a9a442d6ee
***
# Lucy — PPTX Ninja

**Lucy** is a Figma plugin that exports design frames to **editable PowerPoint (PPTX)** files, preserving layout, text, shapes, and visual hierarchy as accurately as possible.

Built by designers, for designers.

---

## v0.7 update

Lucy just got a v0.7 update 🚀

We focused on polishing the UI and improving PPTX export reliability.

What’s new:
- UI polish: updated loading flow, progress feedback, and format-based accents for a clearer export state
- PPTX export fixes: improved handling of clipped backgrounds and wide pill shapes for better slide fidelity

Still building Lucy in public, step by step.  
Try it out, break it, and let me know what you think 👀

---

## Why Lucy

Exporting from Figma to PowerPoint is usually painful:

- text becomes part of an image  
- buttons and pills are flattened  
- icons and UI details disappear into the background  
- everything must be rebuilt manually  

**Lucy fixes this.**

It intelligently separates a Figma frame into:
- a clean background image
- editable text layers
- editable shapes (buttons, pills, cards)
- raster overlays for icons, images, and complex vectors

The result is a **PPTX file that can actually be edited** by designers, managers, and clients.

---

## Current Features (v0.3.5)

### Frame → Slide
- One Figma frame exports to one PowerPoint slide
- Pixel-perfect background rendering

### Text (Editable)
- Fully editable text boxes
- Font size calibrated from Figma (px → pt)
- Bold and italic detection
- Text alignment preserved
- Text opacity preserved

### Shapes (Editable)
- Rectangles, pills, cards, and Auto Layout containers
- Ellipses and lines
- Solid fills and strokes
- Corner radius preserved
- Shape opacity preserved

### Smart Raster Overlays
Lucy automatically rasterizes and layers on top:
- icons and vector graphics
- images
- arrows and complex shapes
- elements with effects or unsupported styles

### Clean Background
- Text and shapes are removed from the background image
- No duplicated or “burned-in” UI elements

### Layer Order
- Z-order preserved (background → shapes → text → overlays)

---

## How Lucy Decides What to Export

Lucy uses a **top-down decision strategy**:

| Figma element | Exported as |
|--------------|------------|
| Text | Editable PowerPoint text |
| Simple shapes (solid, no effects) | Editable PowerPoint shapes |
| Auto Layout frames with solid background | Editable rounded rectangles |
| Icons, vectors, images | Raster overlays (PNG) |
| Complex or effect-heavy elements | Raster overlays |

This approach avoids:
- duplicated layers
- text baked into images
- broken or unpredictable layouts

---

## Known Limitations

- Mixed text styles inside a single text node are simplified (first style is used)
- Gradients, shadows, blur, and blend modes are rasterized
- Masked content is rasterized
- Line height and letter spacing are approximated
- Custom fonts may be substituted in PowerPoint

These trade-offs are intentional to keep exports stable and reliable.

---

### Author

Created and maintained by
Basil Artemov — Senior Product Designer

Portfolio: https://ux.luxury

---

### Lucy — PPTX Ninja
Export once. Edit everywhere.
