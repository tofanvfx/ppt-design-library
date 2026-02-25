# 📐 Design Library — PowerPoint Add-in User Guide

> **Save, organize, and reuse your favorite slide designs with one click.**

---

## What Is Design Library?

Design Library is a PowerPoint add-in that lets you **save any shapes, text boxes, or groups** from your slides into a personal library. You can then **insert them into any presentation** anytime — like building blocks for your slides.

**Think of it as:** Copy-paste that works across presentations and never gets lost.

---

## Where to Find It

After installation, open PowerPoint. You'll see a new **"Design Library"** tab in the top ribbon:

```
File | Home | Insert | Draw | Design | ... | Design Library
```

The tab has **4 groups** of buttons:

| Group | What It Does |
|---|---|
| **Insert Design** | Browse and insert your saved designs |
| **Save Selected Shapes** | Save whatever you've selected on the slide |
| **Slide Size** | Quick-resize slide to 20 × 11.25 inches |
| **Advanced Management** | Open the side panel for full library management |

---

## Step-by-Step: How to Use Each Feature

### 1️⃣ Saving a Design

This is how you save shapes from your slide to reuse later.

1. **Select** one or more shapes on your slide (click them, or Ctrl+Click to select multiple)
2. In the **Design Library** tab, find the **"Save Selected Shapes"** group
3. Type a **Name** (e.g., "Blue Header Bar")
4. Type a **Category** (e.g., "Headers") — this helps organize your library
5. Click **"Save Selection"** ✅

**What happens behind the scenes:**
- The selected shapes are saved as a small `.pptx` file
- A `.png` preview image is also generated
- Both are stored in your personal library folder (`%AppData%\PPTDesignLibrary`)

---

### 2️⃣ Inserting a Saved Design

This is how you reuse a design you saved earlier.

1. In the **Design Library** tab, click **"My Designs"** (the big dropdown button)
2. Your designs appear **organized by category** (e.g., Headers, Footers, Icons)
3. Click on any design name to **insert it** onto your current slide

The shapes will be pasted exactly as you saved them — same size, colors, fonts, and positioning.

---

### 3️⃣ Using the Side Panel (Advanced Management)

The Side Panel gives you full control over your library with previews.

1. Click **"Side Panel Manager"** toggle button in the ribbon
2. A panel opens on the **right side** of PowerPoint

**What you can do in the Side Panel:**

| Action | How |
|---|---|
| **Browse designs** | Scroll through the design list |
| **Filter by category** | Use the dropdown at the top |
| **Preview** | Click any design to see a PNG preview image |
| **Insert** | Select a design → click **"Insert Design"** (blue button) |
| **Save new** | Enter a name + category → click **"Save Current Selection"** (green button) |
| **Rename** | Select a design → click **"Rename"** |
| **Delete** | Select a design → click **"Delete"** (with confirmation) |
| **Refresh** | Click "Refresh" to reload the list |

---

### 4️⃣ Resizing Slides

A quick utility button for a specific slide size:

1. Click **"Resize to 20×11.25"** in the **Slide Size** group
2. Your presentation's slide dimensions change to **20 inches × 11.25 inches** (landscape)

This is useful for large-format or custom-size presentations.

---

## Where Are My Designs Stored?

All your saved designs are stored locally at:
```
C:\Users\<YourName>\AppData\Roaming\PPTDesignLibrary\
```

Inside this folder you'll find:
- `designs.txt` — A metadata file listing all designs (name, category, date)
- `*.pptx` files — The actual saved design shapes
- `*.png` files — Preview thumbnails

> **Tip:** You can back up this folder or copy it to another computer to transfer your design library!

---

## Quick Reference Cheat Sheet

| I want to... | Do this |
|---|---|
| Save shapes for reuse | Select shapes → type Name & Category → click **Save Selection** |
| Insert a saved design | Click **My Designs** dropdown → pick from the category menu |
| See previews | Toggle **Side Panel Manager** → click any design |
| Organize designs | Use categories when saving (e.g., "Logos", "Banners", "Charts") |
| Delete a design | Open Side Panel → select design → click **Delete** |
| Rename a design | Open Side Panel → select design → click **Rename** |
| Resize slides | Click **Resize to 20×11.25** |

---

## Troubleshooting

| Problem | Solution |
|---|---|
| "Design Library" tab not visible | Restart PowerPoint. Check **File → Options → Add-ins** → ensure "Design Library" is enabled |
| Saved design won't insert | Make sure you have a slide open in the active presentation |
| Preview image not showing | The PNG export may have failed during save — the design still works, just without preview |

---

## Requirements

- **Windows 10** or later
- **Microsoft PowerPoint** (Office 2016 / 2019 / 365)
- **.NET Framework 4.5+** (pre-installed on Windows 10+)
