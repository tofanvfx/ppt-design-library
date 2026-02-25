# How to Build the Design Library Installer (.exe)

## One-Time Setup

1. **Download Inno Setup** from: https://jrsoftware.org/isdl.php
2. Install it (default settings are fine)

## Building the Installer

### Option A: Using the GUI
1. Open **Inno Setup Compiler** from your Start menu
2. Click **File → Open** and select `installer.iss` from this folder
3. Press **Ctrl+F9** (or click **Build → Compile**)
4. The `.exe` will be created in the `Output` folder

### Option B: Using Command Line
```powershell
& "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" installer.iss
```

## Output

After building, you'll find:
```
DesignLibrary_COM_AddIn\Output\DesignLibrary_Setup.exe
```

## What the Installer Does on the Target Computer

1. Copies `DesignLibraryAddIn.dll` and `Ribbon.xml` to `%LOCALAPPDATA%\DesignLibrary`
2. Registers the COM component (per-user, no admin required)
3. Writes PowerPoint add-in registry keys
4. Creates an uninstaller in Add/Remove Programs

## Requirements for Target Computer

- Windows 10 or later
- .NET Framework 4.5+ (pre-installed on Windows 10+)
- Microsoft Office / PowerPoint
