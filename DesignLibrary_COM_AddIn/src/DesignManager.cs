using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;
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

    public class OnlineDesignItem
    {
        public string Name { get; set; }
        public string Category { get; set; }
        public string PptxUrl { get; set; }
        public string PreviewUrl { get; set; }
    }

    public class DesignManager
    {
        private const string METADATA_FILE = "designs.txt";
        private const string ONLINE_MANIFEST_URL = "https://raw.githubusercontent.com/tofanvfx/ppt-design-library/main/premade/designs.json";
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

        // ─── LOCAL LIBRARY ────────────────────────────────────────────────────

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

        // ─── ONLINE / GITHUB LIBRARY ──────────────────────────────────────────

        /// <summary>
        /// Fetches the premade designs manifest from GitHub and parses it.
        /// Always fetched live — no caching.
        /// </summary>
        public List<OnlineDesignItem> FetchOnlineDesigns()
        {
            var designs = new List<OnlineDesignItem>();

            string json;
            using (var client = new WebClient())
            {
                // Bypass cached CDN versions
                client.Headers.Add("Cache-Control", "no-cache");
                json = client.DownloadString(ONLINE_MANIFEST_URL);
            }

            // Simple JSON parsing — no external library needed.
            // Each design block: { "name": "...", "category": "...", "pptx_url": "...", "preview_url": "..." }
            var blockPattern = new Regex(@"\{[^{}]+\}", RegexOptions.Singleline);
            var fieldPattern = new Regex(@"""(\w+)""\s*:\s*""([^""]+)""");

            foreach (Match block in blockPattern.Matches(json))
            {
                var fields = new Dictionary<string, string>();
                foreach (Match field in fieldPattern.Matches(block.Value))
                {
                    fields[field.Groups[1].Value] = field.Groups[2].Value;
                }

                if (fields.ContainsKey("name") && fields.ContainsKey("pptx_url"))
                {
                    designs.Add(new OnlineDesignItem
                    {
                        Name        = fields.ContainsKey("name")        ? fields["name"]        : "Untitled",
                        Category    = fields.ContainsKey("category")    ? fields["category"]    : "General",
                        PptxUrl     = fields.ContainsKey("pptx_url")    ? fields["pptx_url"]    : "",
                        PreviewUrl  = fields.ContainsKey("preview_url") ? fields["preview_url"] : ""
                    });
                }
            }

            return designs;
        }

        /// <summary>
        /// Downloads the preview PNG from a URL into a MemoryStream (caller owns lifetime).
        /// Returns null if the URL is empty or download fails.
        /// </summary>
        public System.IO.MemoryStream DownloadPreviewImage(string previewUrl)
        {
            if (string.IsNullOrEmpty(previewUrl)) return null;
            try
            {
                using (var client = new WebClient())
                {
                    byte[] data = client.DownloadData(previewUrl);
                    return new System.IO.MemoryStream(data);
                }
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Downloads the .pptx to a temp file, copies shapes from slide 1 to the
        /// active slide, then deletes the temp file. Read-only by design.
        /// </summary>
        public void InsertOnlineDesign(string pptxUrl)
        {
            if (string.IsNullOrEmpty(pptxUrl))
                throw new ArgumentException("No download URL provided.");

            string tempFile = Path.GetTempFileName() + ".pptx";

            try
            {
                using (var client = new WebClient())
                {
                    client.DownloadFile(pptxUrl, tempFile);
                }

                PowerPoint.Presentation designPres = _ppt.Presentations.Open(
                    tempFile,
                    MsoTriState.msoTrue,  // ReadOnly
                    MsoTriState.msoFalse,
                    MsoTriState.msoFalse  // Hidden window
                );

                try
                {
                    if (designPres.Slides.Count > 0 && designPres.Slides[1].Shapes.Count > 0)
                    {
                        designPres.Slides[1].Shapes.Range().Copy();
                        ((PowerPoint.Slide)_ppt.ActiveWindow.View.Slide).Shapes.Paste();
                    }
                }
                finally
                {
                    designPres.Close();
                }
            }
            finally
            {
                try { if (File.Exists(tempFile)) File.Delete(tempFile); } catch { }
            }
        }
    }
}
