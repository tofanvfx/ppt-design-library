# Premade Design Library

This folder contains all premade design templates that automatically appear in the **"Premade"** tab of the PPT Design Library add-in.

## How to Add a New Premade Design

1. **Place your files here:**
   - `yourdesign.pptx` — the PowerPoint file with your design shapes on slide 1
   - `yourdesign.png` — a preview image (thumbnail) of the design

2. **Update `designs.json`** — add a new entry:
   ```json
   {
     "name": "My New Design",
     "category": "Layouts",
     "pptx_url": "https://raw.githubusercontent.com/tofanvfx/ppt-design-library/main/premade/yourdesign.pptx",
     "preview_url": "https://raw.githubusercontent.com/tofanvfx/ppt-design-library/main/premade/yourdesign.png"
   }
   ```

3. **Commit and push** to the `main` branch — the add-in will pick it up immediately (no reinstall needed).

## Category Suggestions

| Category | Description |
|----------|-------------|
| Text     | Title boxes, text frames, callouts |
| Shapes   | Icons, dividers, decorative shapes |
| Layouts  | Full slide layout templates |
| Charts   | Chart placeholders and infographics |
| Logos    | Logo lockups and brand elements |

## File Naming Tips
- Use lowercase with underscores: `hero_title_card.pptx`
- Keep `.pptx` and `.png` names matching (same base name)
- The `.pptx` should have the design shapes on **Slide 1** only
