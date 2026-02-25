# 📐 PPT Design Library

A PowerPoint COM Add-in that lets you **save, organize, and reuse slide design elements** — like building blocks for your presentations.

> Save any shapes, text boxes, or groups into a personal library and insert them into any presentation with one click.

---

## ✨ Features

- **Save Designs** — Select shapes on your slide, give them a name and category, and save to your library
- **Insert Designs** — Browse your library from the ribbon dropdown and insert saved designs instantly
- **Side Panel Manager** — Visual panel with preview thumbnails, filtering, renaming, and deleting
- **Category Organization** — Organize designs into categories (Headers, Logos, Charts, etc.)
- **Slide Resize** — Quick button to resize slides to 20 × 11.25 inches
- **PNG Previews** — Auto-generated thumbnails for each saved design

---

## 🖥️ Requirements

- **Windows 10** or later
- **Microsoft PowerPoint** (Office 2016 / 2019 / 365)
- **.NET Framework 4.5+** (pre-installed on Windows 10+)

> ⚠️ **Windows only** — COM Add-ins are not supported on macOS.

---

## 📦 Installation

### Option A: Run the Installer
1. Download `DesignLibrary_Setup.exe` from [Releases](https://github.com/tofanvfx/ppt-design-library/releases)
2. Run as **Administrator**
3. Restart PowerPoint — you'll see the **"Design Library"** tab in the ribbon

### Option B: Build from Source
```powershell
# 1. Compile the DLL
.\DesignLibrary_COM_AddIn\build_dll.ps1

# 2. Restart PowerPoint
```

---

## 🔧 Building the Installer (.exe)

1. Install [Inno Setup](https://jrsoftware.org/isdl.php) (free)
2. Open `DesignLibrary_COM_AddIn\installer.iss` in Inno Setup Compiler
3. Press **Ctrl+F9** to compile
4. Find the output at `DesignLibrary_COM_AddIn\Output\DesignLibrary_Setup.exe`

See [BUILD_INSTALLER.md](DesignLibrary_COM_AddIn/BUILD_INSTALLER.md) for detailed instructions.

---

## 📖 How It Works

| Ribbon Group | Action |
|---|---|
| **Insert Design** | Click "My Designs" dropdown → pick a saved design by category |
| **Save Selected Shapes** | Select shapes → enter Name & Category → click "Save Selection" |
| **Slide Size** | One-click resize to 20 × 11.25 inches |
| **Advanced Management** | Toggle the side panel for previews, rename, delete, and filtering |

Designs are stored locally at `%AppData%\PPTDesignLibrary\` as `.pptx` + `.png` files.

See [USER_GUIDE.md](DesignLibrary_COM_AddIn/USER_GUIDE.md) for the full step-by-step guide.

---

## 📁 Project Structure

```
ppt_addin/
├── DesignLibrary_COM_AddIn/
│   ├── src/
│   │   ├── AddIn.cs              # Main COM Add-in (ribbon, callbacks)
│   │   ├── DesignManager.cs      # Save/insert/delete design logic
│   │   ├── TaskPaneControl.cs    # Side panel UI with preview
│   │   ├── LibraryForm.cs        # Standalone library window
│   │   └── Ribbon.xml            # Ribbon tab definition
│   ├── build_dll.ps1             # Build & register script
│   ├── installer.iss             # Inno Setup installer script
│   ├── BUILD_INSTALLER.md        # Installer build instructions
│   └── USER_GUIDE.md             # End-user guide
└── .gitignore
```

---

## 📄 License

This project is provided as-is for educational and personal use.

---

## 👤 Author

**Aveti Learning** — [tofanvfx](https://github.com/tofanvfx)
