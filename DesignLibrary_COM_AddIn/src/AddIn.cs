using System;
using System.IO;
using System.Runtime.InteropServices;
using Extensibility;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.Windows.Forms;
using System.Text;
using System.Linq;

namespace DesignLibraryAddIn
{
    [ComVisible(true)]
    [Guid("CF8DBA7F-EDDE-4A8E-AF10-C3D7BB89EE69")]
    [ProgId("DesignLibraryAddIn.AddIn")]
    public class AddIn : IDTExtensibility2, IRibbonExtensibility, ICustomTaskPaneConsumer
    {
        private PowerPoint.Application _pptApplication;
        private object _addInInstance;
        private CustomTaskPane _taskPane;
        private IRibbonUI _ribbon;
        private string _currentSaveName = "";
        private string _currentSaveCategory = "General";

        // IDTExtensibility2 Methods
        public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            _pptApplication = (PowerPoint.Application)Application;
            _addInInstance = AddInInst;
        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            _pptApplication = null;
        }

        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { }

        // ICustomTaskPaneConsumer Method
        public void CTPFactoryAvailable(ICTPFactory CTPFactoryInst)
        {
            try
            {
                _taskPane = CTPFactoryInst.CreateCTP("DesignLibraryAddIn.TaskPaneControl", "Design Library");
                _taskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
                _taskPane.Width = 320;
                
                TaskPaneControl control = (TaskPaneControl)_taskPane.ContentControl;
                DesignManager manager = new DesignManager(_pptApplication);
                control.Initialize(manager);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error initializing Task Pane: " + ex.Message);
            }
        }

        // IRibbonExtensibility Method
        public string GetCustomUI(string RibbonID)
        {
            try
            {
                // Path to the Ribbon.xml file (will be copied alongside DLL)
                string dllPath = typeof(AddIn).Assembly.Location;
                string xmlPath = Path.Combine(Path.GetDirectoryName(dllPath), "Ribbon.xml");
                if (File.Exists(xmlPath))
                    return File.ReadAllText(xmlPath);
                else
                {
                    MessageBox.Show("Ribbon.xml not found at " + xmlPath);
                    return "";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading Ribbon UI: " + ex.Message);
                return "";
            }
        }

        // Ribbon Callbacks
        public void OnLoad(IRibbonUI ribbon)
        {
            _ribbon = ribbon;
        }

        public void OnNameChange(IRibbonControl control, string text)
        {
            _currentSaveName = text;
        }
        public string GetDesignName(IRibbonControl control)
        {
            return _currentSaveName;
        }

        public void OnCategoryChange(IRibbonControl control, string text)
        {
            _currentSaveCategory = text;
        }
        public string GetDesignCategory(IRibbonControl control)
        {
            return _currentSaveCategory;
        }

        public void OnSaveDesign(IRibbonControl control)
        {
            try
            {
                if (_pptApplication.ActiveWindow == null || _pptApplication.ActiveWindow.Selection == null) return;
                
                var selection = _pptApplication.ActiveWindow.Selection;
                if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes && selection.Type != PowerPoint.PpSelectionType.ppSelectionText)
                {
                    MessageBox.Show("Please select one or more shapes first.", "Design Library", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }

                if (string.IsNullOrWhiteSpace(_currentSaveName))
                {
                    MessageBox.Show("Please type a name in the 'Name' box on the Ribbon first.", "Design Library", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                string cat = string.IsNullOrWhiteSpace(_currentSaveCategory) ? "General" : _currentSaveCategory;

                DesignManager manager = new DesignManager(_pptApplication);
                manager.SaveSelectedAsDesign(_currentSaveName, cat);
                
                _currentSaveName = "";
                if (_ribbon != null)
                {
                    _ribbon.InvalidateControl("txtDesignName");
                    _ribbon.InvalidateControl("dynLibrary");
                }
                
                MessageBox.Show("Design saved!", "Design Library", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error saving design: " + ex.Message, "Design Library", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private string EscapeXml(string s)
        {
            if (string.IsNullOrEmpty(s)) return s;
            return s.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;").Replace("'", "&apos;");
        }

        public string GetLibraryMenuContent(IRibbonControl control)
        {
            DesignManager manager = new DesignManager(_pptApplication);
            var designs = manager.GetAllDesigns();
            
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("<menu xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\">");
            
            var categories = designs.Select(d => d.Category).Distinct().OrderBy(c => c).ToList();
            
            foreach (var cat in categories)
            {
                sb.AppendFormat("<menu id=\"cat_{0}\" label=\"{1}\">\n", Guid.NewGuid().ToString("N"), EscapeXml(cat));
                var catDesigns = designs.Where(d => d.Category == cat).OrderBy(d => d.DisplayName).ToList();
                foreach (var d in catDesigns)
                {
                    sb.AppendFormat("<button id=\"insert_{0}\" label=\"{1}\" onAction=\"OnInsertDesignFromMenu\" tag=\"{2}\" imageMso=\"Paste\" />\n", 
                        Guid.NewGuid().ToString("N"), 
                        EscapeXml(d.DisplayName),
                        EscapeXml(d.FileName));
                }
                sb.AppendLine("</menu>");
            }
            
            if (designs.Count == 0)
            {
                sb.AppendLine("<button id=\"emptyBtn\" label=\"No designs saved yet\" onAction=\"OnDoNothing\" />");
            }
            
            sb.AppendLine("</menu>");
            return sb.ToString();
        }

        public void OnInsertDesignFromMenu(IRibbonControl control)
        {
            try
            {
                string fileName = control.Tag;
                if (!string.IsNullOrEmpty(fileName))
                {
                    DesignManager manager = new DesignManager(_pptApplication);
                    manager.InsertDesign(fileName);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error inserting design: " + ex.Message, "Design Library", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnDoNothing(IRibbonControl control) { }

        public void OnResizeSlide(IRibbonControl control)
        {
            try
            {
                if (_pptApplication.ActivePresentation == null) return;
                
                var presentation = _pptApplication.ActivePresentation;
                
                // 20 inches * 72 points/inch = 1440
                // 11.25 inches * 72 points/inch = 810
                presentation.PageSetup.SlideWidth = 1440f;
                presentation.PageSetup.SlideHeight = 810f;
                
                MessageBox.Show("Slide size updated to 20 x 11.25 inches.", "Design Library", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error resizing slide: " + ex.Message, "Design Library", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnToggleTaskPane(IRibbonControl control, bool isPressed)
        {
            if (_taskPane != null)
            {
                _taskPane.Visible = isPressed;
            }
        }

        public bool GetTaskPaneState(IRibbonControl control)
        {
            if (_taskPane != null)
                return _taskPane.Visible;
            return false;
        }
    }
}
