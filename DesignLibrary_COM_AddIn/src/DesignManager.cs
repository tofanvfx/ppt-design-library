using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace DesignLibraryAddIn
{
    public class DesignItem
    {
        public string FileName { get; set; }
        public string DisplayName { get; set; }
        public string Category { get; set; }
        public string CreatedDate { get; set; }
    }

    public class DesignManager
    {
        private const string METADATA_FILE = "designs.txt";
        private PowerPoint.Application _ppt;

        public DesignManager(PowerPoint.Application pptApp)
        {
            _ppt = pptApp;
        }

        public string GetLibraryPath()
        {
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string libraryPath = Path.Combine(appData, "PPTDesignLibrary");
            if (!Directory.Exists(libraryPath))
            {
                Directory.CreateDirectory(libraryPath);
            }
            return libraryPath;
        }

        private string SanitizeName(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return "Untitled";
            char[] invalidChars = Path.GetInvalidFileNameChars();
            string sanitized = new string(name.Select(c => invalidChars.Contains(c) ? '_' : c).ToArray());
            if (sanitized.Length > 50) sanitized = sanitized.Substring(0, 50);
            return sanitized.Trim();
        }

        public void SaveSelectedAsDesign(string designName, string category)
        {
            if (_ppt.ActiveWindow == null || _ppt.ActiveWindow.Selection == null) return;
            var selection = _ppt.ActiveWindow.Selection;

            string baseName = string.Format("{0}_{1:yyyyMMddHHmmss}", SanitizeName(designName), DateTime.Now);
            string fileName = baseName + ".pptx";
            string filePath = Path.Combine(GetLibraryPath(), fileName);
            string pngPath = Path.Combine(GetLibraryPath(), baseName + ".png");

            try
            {
                selection.ShapeRange.Export(pngPath, PowerPoint.PpShapeFormat.ppShapeFormatPNG);
            }
            catch { /* Ignore if it fails to export image */ }

            selection.ShapeRange.Copy();

            PowerPoint.Presentation tempPres = _ppt.Presentations.Add(MsoTriState.msoFalse);
            tempPres.PageSetup.SlideWidth = _ppt.ActivePresentation.PageSetup.SlideWidth;
            tempPres.PageSetup.SlideHeight = _ppt.ActivePresentation.PageSetup.SlideHeight;

            PowerPoint.Slide tempSlide = tempPres.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank);
            tempSlide.Shapes.Paste();

            tempPres.SaveAs(filePath, PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation);
            tempPres.Close();

            SaveMetadata(fileName, designName, category);
        }

        public void InsertDesign(string fileName)
        {
            string filePath = Path.Combine(GetLibraryPath(), fileName);
            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException(string.Format("Design file not found at {0}", filePath));
            }

            PowerPoint.Presentation designPres = _ppt.Presentations.Open(filePath, MsoTriState.msoTrue, MsoTriState.msoFalse, MsoTriState.msoFalse);
            
            if (designPres.Slides.Count > 0 && designPres.Slides[1].Shapes.Count > 0)
            {
                designPres.Slides[1].Shapes.Range().Copy();
                ((PowerPoint.Slide)_ppt.ActiveWindow.View.Slide).Shapes.Paste();
            }

            designPres.Close();
        }

        public List<DesignItem> GetAllDesigns()
        {
            var designs = new List<DesignItem>();
            string metaPath = Path.Combine(GetLibraryPath(), METADATA_FILE);

            if (!File.Exists(metaPath)) return designs;

            string[] lines = File.ReadAllLines(metaPath);
            foreach (string line in lines)
            {
                if (string.IsNullOrWhiteSpace(line)) continue;
                string[] parts = line.Split('|');
                if (parts.Length >= 4)
                {
                    if (File.Exists(Path.Combine(GetLibraryPath(), parts[0])))
                    {
                        designs.Add(new DesignItem
                        {
                            FileName = parts[0],
                            DisplayName = parts[1],
                            Category = parts[2],
                            CreatedDate = parts[3]
                        });
                    }
                }
            }
            return designs;
        }

        public void DeleteDesign(string fileName)
        {
            string filePath = Path.Combine(GetLibraryPath(), fileName);
            string pngPath = Path.Combine(GetLibraryPath(), fileName.Replace(".pptx", ".png"));

            if (File.Exists(filePath))
            {
                File.Delete(filePath);
            }
            if (File.Exists(pngPath))
            {
                File.Delete(pngPath);
            }
            RemoveMetadata(fileName);
        }

        public void RenameDesign(string fileName, string newName)
        {
            string metaPath = Path.Combine(GetLibraryPath(), METADATA_FILE);
            if (!File.Exists(metaPath)) return;

            var lines = File.ReadAllLines(metaPath).ToList();
            for (int i = 0; i < lines.Count; i++)
            {
                string[] parts = lines[i].Split('|');
                if (parts.Length >= 4 && parts[0] == fileName)
                {
                    lines[i] = string.Format("{0}|{1}|{2}|{3}", parts[0], newName, parts[2], parts[3]);
                    break;
                }
            }
            File.WriteAllLines(metaPath, lines);
            
            // Note: We don't necessarily rename the underlying file (the ID), just the DisplayName.
            // If we did want to rename the file, we would rename both .pptx and .png here.
            // But since the system uses 'fileName' purely as a unique ID backed by timestamp,
            // changing the display name in 'designs.txt' is all that's needed for the user.
        }

        private void SaveMetadata(string fileName, string designName, string category)
        {
            string metaPath = Path.Combine(GetLibraryPath(), METADATA_FILE);
            string line = string.Format("{0}|{1}|{2}|{3:yyyy-MM-dd HH:mm}", fileName, designName, category, DateTime.Now);
            File.AppendAllLines(metaPath, new[] { line });
        }

        private void RemoveMetadata(string fileName)
        {
            string metaPath = Path.Combine(GetLibraryPath(), METADATA_FILE);
            if (!File.Exists(metaPath)) return;

            var lines = File.ReadAllLines(metaPath).Where(l => !l.StartsWith(fileName + "|")).ToList();
            File.WriteAllLines(metaPath, lines);
        }
    }
}
