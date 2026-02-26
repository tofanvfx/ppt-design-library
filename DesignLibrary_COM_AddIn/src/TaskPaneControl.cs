using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace DesignLibraryAddIn
{
    [ComVisible(true)]
    [Guid("5C5E7A22-8CD9-411B-87C3-8B39B2A4B8F6")]
    [ProgId("DesignLibraryAddIn.TaskPaneControl")]
    public class TaskPaneControl : UserControl
    {
        private DesignManager _manager;

        // ── Shared colours ──────────────────────────────────────────────────
        private readonly Color _primaryBlue  = ColorTranslator.FromHtml("#0078D4");
        private readonly Color _primaryGreen = ColorTranslator.FromHtml("#107C41");
        private readonly Color _dangerRed    = ColorTranslator.FromHtml("#A4262C");
        private readonly Color _lightGray    = ColorTranslator.FromHtml("#F3F2F1");
        private readonly Color _borderGray   = ColorTranslator.FromHtml("#EDEBE9");
        private readonly Color _textDark     = ColorTranslator.FromHtml("#323130");
        private readonly Color _textLight    = ColorTranslator.FromHtml("#605E5C");

        // ── My Library tab controls ──────────────────────────────────────────
        private ListBox    _lstDesigns;
        private ComboBox   _cboCategory;
        private TextBox    _txtName;
        private TextBox    _txtCategory;
        private Label      _lblStatus;
        private Button     _btnSave;
        private Button     _btnInsert;
        private Button     _btnDelete;
        private Button     _btnRename;
        private PictureBox _picPreview;

        // ── Premade tab controls ─────────────────────────────────────────────
        private ListBox    _lstOnline;
        private ComboBox   _cboCategoryOnline;
        private PictureBox _picOnlinePreview;
        private Label      _lblOnlineStatus;
        private Button     _btnOnlineInsert;

        // ── Online data ──────────────────────────────────────────────────────
        private List<OnlineDesignItem> _onlineDesigns = new List<OnlineDesignItem>();

        public TaskPaneControl()
        {
            InitializeComponent();
        }

        public void Initialize(DesignManager manager)
        {
            _manager = manager;
            RefreshCategories();
            RefreshList();
        }

        // ════════════════════════════════════════════════════════════════════
        // UI CONSTRUCTION
        // ════════════════════════════════════════════════════════════════════

        private void InitializeComponent()
        {
            this.Size      = new Size(300, 640);
            this.BackColor = Color.White;
            this.AutoScroll = false;
            this.Font      = new Font("Segoe UI", 9F);

            var tabs = new TabControl
            {
                Dock     = DockStyle.Fill,
                Font     = new Font("Segoe UI", 9F),
                Padding  = new Point(8, 4)
            };

            var pageLocal   = new TabPage("My Library")  { BackColor = Color.White, Padding = new Padding(0) };
            var pageOnline  = new TabPage("⭐ Premade")  { BackColor = Color.White, Padding = new Padding(0) };

            pageLocal.Controls.Add(BuildLocalPanel());
            pageOnline.Controls.Add(BuildOnlinePanel());

            tabs.TabPages.Add(pageLocal);
            tabs.TabPages.Add(pageOnline);
            tabs.SelectedIndexChanged += (s, e) =>
            {
                if (tabs.SelectedIndex == 1 && _onlineDesigns.Count == 0)
                    LoadOnlineDesigns();
            };

            this.Controls.Add(tabs);
        }

        // ── LOCAL PANEL ──────────────────────────────────────────────────────

        private Panel BuildLocalPanel()
        {
            var panel  = new Panel { Dock = DockStyle.Fill, AutoScroll = true, BackColor = Color.White };
            int y      = 12;
            int margin = 12;
            int width  = 270;

            // Title
            panel.Controls.Add(new Label
            {
                Text      = "DESIGN LIBRARY",
                Location  = new Point(margin, y),
                Size      = new Size(width, 24),
                Font      = new Font("Segoe UI Semibold", 12F),
                ForeColor = _textDark
            });
            y += 24;
            panel.Controls.Add(new Label
            {
                Text      = "Save and reuse slide objects",
                Location  = new Point(margin, y),
                Size      = new Size(width, 15),
                ForeColor = _textLight
            });
            y += 22;
            panel.Controls.Add(new Label { BackColor = _borderGray, Location = new Point(margin, y), Size = new Size(width, 1) });
            y += 12;

            // Filter row
            panel.Controls.Add(new Label { Text = "Filter:", Location = new Point(margin, y + 4), AutoSize = true, ForeColor = _textDark });
            _cboCategory = new ComboBox { Location = new Point(55, y), Size = new Size(138, 25), DropDownStyle = ComboBoxStyle.DropDownList, FlatStyle = FlatStyle.Flat, BackColor = _lightGray };
            _cboCategory.SelectedIndexChanged += (s, e) => RefreshList();
            panel.Controls.Add(_cboCategory);

            var btnRefresh = MakeButton("Refresh", new Point(203, y - 1), new Size(80, 26), _lightGray, _textDark);
            btnRefresh.Click += (s, e) => { RefreshCategories(); RefreshList(); };
            panel.Controls.Add(btnRefresh);
            y += 34;

            // Design list
            _lstDesigns = new ListBox { Location = new Point(margin, y), Size = new Size(width, 110), Font = new Font("Segoe UI", 9F), BorderStyle = BorderStyle.FixedSingle, BackColor = Color.White, ForeColor = _textDark };
            _lstDesigns.SelectedIndexChanged += LstDesigns_SelectedIndexChanged;
            panel.Controls.Add(_lstDesigns);
            y += 118;

            // Preview
            _picPreview = new PictureBox { Location = new Point(margin, y), Size = new Size(width, 110), SizeMode = PictureBoxSizeMode.Zoom, BorderStyle = BorderStyle.FixedSingle, BackColor = Color.White };
            panel.Controls.Add(_picPreview);
            y += 118;

            // Action buttons
            _btnInsert = MakeButton("Insert Design", new Point(margin, y), new Size(width, 34), _primaryBlue, Color.White, "Segoe UI Semibold", 9.5F);
            _btnInsert.Click += BtnInsert_Click;
            panel.Controls.Add(_btnInsert);
            y += 40;

            _btnRename = MakeButton("Rename", new Point(margin, y), new Size(width, 30), _lightGray, _textDark);
            _btnRename.Click += BtnRename_Click;
            panel.Controls.Add(_btnRename);
            y += 36;

            _btnDelete = MakeButton("Delete", new Point(margin, y), new Size(width, 30), _lightGray, _dangerRed);
            _btnDelete.Click += BtnDelete_Click;
            panel.Controls.Add(_btnDelete);
            y += 42;

            panel.Controls.Add(new Label { BackColor = _borderGray, Location = new Point(margin, y), Size = new Size(width, 1) });
            y += 12;

            // Save section
            panel.Controls.Add(new Label { Text = "SAVE SELECTED SHAPES", Location = new Point(margin, y), Size = new Size(width, 18), Font = new Font("Segoe UI Semibold", 9.5F), ForeColor = _textDark });
            y += 22;

            panel.Controls.Add(new Label { Text = "Name:", Location = new Point(margin, y + 4), AutoSize = true, ForeColor = _textDark });
            _txtName = new TextBox { Location = new Point(75, y), Size = new Size(195, 25), BorderStyle = BorderStyle.FixedSingle };
            panel.Controls.Add(_txtName);
            y += 30;

            panel.Controls.Add(new Label { Text = "Category:", Location = new Point(margin, y + 4), AutoSize = true, ForeColor = _textDark });
            _txtCategory = new TextBox { Location = new Point(75, y), Size = new Size(195, 25), Text = "General", BorderStyle = BorderStyle.FixedSingle };
            panel.Controls.Add(_txtCategory);
            y += 36;

            _btnSave = MakeButton("Save Current Selection", new Point(margin, y), new Size(width, 34), _primaryGreen, Color.White, "Segoe UI Semibold", 9.5F);
            _btnSave.Click += BtnSave_Click;
            panel.Controls.Add(_btnSave);
            y += 42;

            _lblStatus = new Label { Location = new Point(margin, y), Size = new Size(width, 40), ForeColor = _textLight };
            panel.Controls.Add(_lblStatus);

            return panel;
        }

        // ── ONLINE / PREMADE PANEL ───────────────────────────────────────────

        private Panel BuildOnlinePanel()
        {
            var panel  = new Panel { Dock = DockStyle.Fill, AutoScroll = true, BackColor = Color.White };
            int y      = 12;
            int margin = 12;
            int width  = 270;

            // Title
            panel.Controls.Add(new Label
            {
                Text      = "PREMADE TEMPLATES",
                Location  = new Point(margin, y),
                Size      = new Size(width, 24),
                Font      = new Font("Segoe UI Semibold", 12F),
                ForeColor = _textDark
            });
            y += 24;
            panel.Controls.Add(new Label
            {
                Text      = "Ready-made designs from TofanVFX",
                Location  = new Point(margin, y),
                Size      = new Size(width, 15),
                ForeColor = _textLight
            });
            y += 22;
            panel.Controls.Add(new Label { BackColor = _borderGray, Location = new Point(margin, y), Size = new Size(width, 1) });
            y += 12;

            // Filter + Refresh row
            panel.Controls.Add(new Label { Text = "Filter:", Location = new Point(margin, y + 4), AutoSize = true, ForeColor = _textDark });
            _cboCategoryOnline = new ComboBox { Location = new Point(55, y), Size = new Size(138, 25), DropDownStyle = ComboBoxStyle.DropDownList, FlatStyle = FlatStyle.Flat, BackColor = _lightGray };
            _cboCategoryOnline.SelectedIndexChanged += (s, e) => RefreshOnlineList();
            panel.Controls.Add(_cboCategoryOnline);

            var btnRefreshOnline = MakeButton("Refresh", new Point(203, y - 1), new Size(80, 26), _lightGray, _textDark);
            btnRefreshOnline.Click += (s, e) => LoadOnlineDesigns();
            panel.Controls.Add(btnRefreshOnline);
            y += 34;

            // Status label (shows loading / error / count)
            _lblOnlineStatus = new Label
            {
                Location  = new Point(margin, y),
                Size      = new Size(width, 18),
                ForeColor = _textLight,
                Text      = "Click Refresh or switch to this tab to load."
            };
            panel.Controls.Add(_lblOnlineStatus);
            y += 24;

            // Design list
            _lstOnline = new ListBox { Location = new Point(margin, y), Size = new Size(width, 130), Font = new Font("Segoe UI", 9F), BorderStyle = BorderStyle.FixedSingle, BackColor = Color.White, ForeColor = _textDark };
            _lstOnline.SelectedIndexChanged += LstOnline_SelectedIndexChanged;
            panel.Controls.Add(_lstOnline);
            y += 138;

            // Preview
            _picOnlinePreview = new PictureBox { Location = new Point(margin, y), Size = new Size(width, 130), SizeMode = PictureBoxSizeMode.Zoom, BorderStyle = BorderStyle.FixedSingle, BackColor = Color.White };
            panel.Controls.Add(_picOnlinePreview);
            y += 138;

            // Insert button
            _btnOnlineInsert = MakeButton("Insert Design", new Point(margin, y), new Size(width, 34), _primaryBlue, Color.White, "Segoe UI Semibold", 9.5F);
            _btnOnlineInsert.Click += BtnOnlineInsert_Click;
            panel.Controls.Add(_btnOnlineInsert);
            y += 42;

            // Read-only notice
            panel.Controls.Add(new Label
            {
                Text      = "🔒 Premade designs are read-only.\nTo add your own, use \"My Library\" tab.",
                Location  = new Point(margin, y),
                Size      = new Size(width, 36),
                ForeColor = _textLight,
                Font      = new Font("Segoe UI", 8F)
            });

            return panel;
        }

        // ════════════════════════════════════════════════════════════════════
        // HELPERS
        // ════════════════════════════════════════════════════════════════════

        private Button MakeButton(string text, Point location, Size size,
                                  Color backColor, Color foreColor,
                                  string fontFamily = "Segoe UI", float fontSize = 9F)
        {
            var btn = new Button
            {
                Text      = text,
                Location  = location,
                Size      = size,
                FlatStyle = FlatStyle.Flat,
                BackColor = backColor,
                ForeColor = foreColor,
                Font      = new Font(fontFamily, fontSize)
            };
            btn.FlatAppearance.BorderSize = 0;
            return btn;
        }

        // ════════════════════════════════════════════════════════════════════
        // LOCAL LIBRARY EVENT HANDLERS
        // ════════════════════════════════════════════════════════════════════

        private void BtnSave_Click(object sender, EventArgs e)
        {
            if (_manager == null) return;
            try
            {
                string name = _txtName.Text.Trim();
                if (string.IsNullOrEmpty(name)) { Status("Please enter a design name.", true); return; }
                string category = _txtCategory.Text.Trim();
                if (string.IsNullOrEmpty(category)) category = "General";

                Status("Saving...");
                _manager.SaveSelectedAsDesign(name, category);
                _txtName.Text = "";
                Status("Saved!");
                RefreshCategories();
                RefreshList();
            }
            catch (Exception ex) { Status("Error: " + ex.Message, true); }
        }

        private void BtnInsert_Click(object sender, EventArgs e)
        {
            if (_manager == null) return;
            try
            {
                var item = _lstDesigns.SelectedItem as DisplayItem;
                if (item != null) { Status("Inserting..."); _manager.InsertDesign(item.Design.FileName); Status("Inserted!"); }
                else Status("Select a design first.", true);
            }
            catch (Exception ex) { Status("Error: " + ex.Message, true); }
        }

        private void LstDesigns_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_lstDesigns.SelectedItem == null) { ClearPreview(_picPreview); return; }
            var selected = (DisplayItem)_lstDesigns.SelectedItem;
            string pngPath = Path.Combine(_manager.GetLibraryPath(), selected.Design.FileName.Replace(".pptx", ".png"));
            LoadLocalPreview(_picPreview, pngPath);
        }

        private void BtnRename_Click(object sender, EventArgs e)
        {
            if (_manager == null) return;
            var item = _lstDesigns.SelectedItem as DisplayItem;
            if (item != null)
            {
                string newName = Microsoft.VisualBasic.Interaction.InputBox("New name:", "Rename Design", item.Design.DisplayName);
                if (!string.IsNullOrWhiteSpace(newName)) { _manager.RenameDesign(item.Design.FileName, newName.Trim()); RefreshList(); Status("Renamed."); }
            }
            else Status("Select a design first.", true);
        }

        private void BtnDelete_Click(object sender, EventArgs e)
        {
            if (_manager == null) return;
            var item = _lstDesigns.SelectedItem as DisplayItem;
            if (item != null)
            {
                var result = MessageBox.Show(string.Format("Delete '{0}'?", item.Design.DisplayName), "Confirm Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (result == DialogResult.Yes) { _manager.DeleteDesign(item.Design.FileName); RefreshCategories(); RefreshList(); Status("Deleted."); }
            }
            else Status("Select a design first.", true);
        }

        private void RefreshCategories()
        {
            if (_manager == null) return;
            string current = _cboCategory.SelectedItem != null ? _cboCategory.SelectedItem.ToString() : null;
            _cboCategory.Items.Clear();
            _cboCategory.Items.Add("All Categories");

            var cats = _manager.GetAllDesigns().Select(d => d.Category).Distinct().OrderBy(c => c);
            foreach (var cat in cats) _cboCategory.Items.Add(cat);

            if (current != null && _cboCategory.Items.Contains(current))
                _cboCategory.SelectedItem = current;
            else
                _cboCategory.SelectedIndex = 0;
        }

        private void RefreshList()
        {
            if (_manager == null) return;
            _lstDesigns.Items.Clear();
            string selCat = _cboCategory.SelectedIndex > 0 ? _cboCategory.SelectedItem.ToString() : null;

            foreach (var d in _manager.GetAllDesigns())
            {
                if (selCat != null && d.Category != selCat) continue;
                _lstDesigns.Items.Add(new DisplayItem { Design = d });
            }

            Status(_lstDesigns.Items.Count == 0
                ? "No designs yet. Select shapes and save!"
                : string.Format("{0} design(s).", _lstDesigns.Items.Count));
        }

        private void Status(string msg, bool isError = false)
        {
            _lblStatus.Text      = msg;
            _lblStatus.ForeColor = isError ? Color.DarkRed : Color.Gray;
        }

        // ════════════════════════════════════════════════════════════════════
        // ONLINE / PREMADE EVENT HANDLERS
        // ════════════════════════════════════════════════════════════════════

        private void LoadOnlineDesigns()
        {
            if (_manager == null) return;

            OnlineStatus("Loading from GitHub...");
            _lstOnline.Items.Clear();
            ClearPreview(_picOnlinePreview);
            _onlineDesigns.Clear();

            // Run fetch on background thread so UI doesn't freeze
            ThreadPool.QueueUserWorkItem(_ =>
            {
                try
                {
                    var designs = _manager.FetchOnlineDesigns();
                    this.Invoke((MethodInvoker)delegate
                    {
                        _onlineDesigns = designs;
                        PopulateOnlineCategories();
                        RefreshOnlineList();
                    });
                }
                catch (Exception ex)
                {
                    this.Invoke((MethodInvoker)delegate
                    {
                        OnlineStatus("⚠ Could not load: " + ex.Message, true);
                    });
                }
            });
        }

        private void PopulateOnlineCategories()
        {
            string current = _cboCategoryOnline.SelectedItem != null ? _cboCategoryOnline.SelectedItem.ToString() : null;
            _cboCategoryOnline.Items.Clear();
            _cboCategoryOnline.Items.Add("All Categories");
            foreach (var cat in _onlineDesigns.Select(d => d.Category).Distinct().OrderBy(c => c))
                _cboCategoryOnline.Items.Add(cat);

            if (current != null && _cboCategoryOnline.Items.Contains(current))
                _cboCategoryOnline.SelectedItem = current;
            else
                _cboCategoryOnline.SelectedIndex = 0;
        }

        private void RefreshOnlineList()
        {
            _lstOnline.Items.Clear();
            string selCat = (_cboCategoryOnline.SelectedIndex > 0)
                ? _cboCategoryOnline.SelectedItem.ToString()
                : null;

            foreach (var d in _onlineDesigns)
            {
                if (selCat != null && d.Category != selCat) continue;
                _lstOnline.Items.Add(new OnlineDisplayItem { Design = d });
            }

            OnlineStatus(_lstOnline.Items.Count == 0
                ? "No premade designs found."
                : string.Format("{0} template(s) available.", _lstOnline.Items.Count));
        }

        private void LstOnline_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_lstOnline.SelectedItem == null) { ClearPreview(_picOnlinePreview); return; }
            var item = (OnlineDisplayItem)_lstOnline.SelectedItem;
            if (string.IsNullOrEmpty(item.Design.PreviewUrl)) return;

            OnlineStatus("Loading preview...");

            ThreadPool.QueueUserWorkItem(_ =>
            {
                var stream = _manager.DownloadPreviewImage(item.Design.PreviewUrl);
                this.Invoke((MethodInvoker)delegate
                {
                    ClearPreview(_picOnlinePreview);
                    if (stream != null)
                    {
                        try
                        {
                            var img = Image.FromStream(stream);
                            _picOnlinePreview.Image = new Bitmap(img);
                        }
                        catch { }
                    }
                    OnlineStatus(string.Format("{0} template(s) available.", _lstOnline.Items.Count));
                });
            });
        }

        private void BtnOnlineInsert_Click(object sender, EventArgs e)
        {
            if (_manager == null) return;
            var item = _lstOnline.SelectedItem as OnlineDisplayItem;
            if (item == null) { OnlineStatus("Select a template first.", true); return; }

            OnlineStatus("Downloading and inserting...");
            _btnOnlineInsert.Enabled = false;

            ThreadPool.QueueUserWorkItem(_ =>
            {
                try
                {
                    _manager.InsertOnlineDesign(item.Design.PptxUrl);
                    this.Invoke((MethodInvoker)delegate
                    {
                        OnlineStatus("Inserted!");
                        _btnOnlineInsert.Enabled = true;
                    });
                }
                catch (Exception ex)
                {
                    this.Invoke((MethodInvoker)delegate
                    {
                        OnlineStatus("Error: " + ex.Message, true);
                        _btnOnlineInsert.Enabled = true;
                    });
                }
            });
        }

        private void OnlineStatus(string msg, bool isError = false)
        {
            _lblOnlineStatus.Text      = msg;
            _lblOnlineStatus.ForeColor = isError ? Color.DarkRed : Color.Gray;
        }

        // ════════════════════════════════════════════════════════════════════
        // SHARED HELPERS
        // ════════════════════════════════════════════════════════════════════

        private void ClearPreview(PictureBox pic)
        {
            if (pic.Image != null) { pic.Image.Dispose(); pic.Image = null; }
        }

        private void LoadLocalPreview(PictureBox pic, string pngPath)
        {
            ClearPreview(pic);
            if (!File.Exists(pngPath)) return;
            try
            {
                using (var fs = new FileStream(pngPath, FileMode.Open, FileAccess.Read))
                {
                    var img = Image.FromStream(fs);
                    pic.Image = new Bitmap(img);
                }
            }
            catch { }
        }

        // ── Display helpers ──────────────────────────────────────────────────

        private class DisplayItem
        {
            public DesignItem Design { get; set; }
            public override string ToString()
            {
                return string.Format("{0} [{1}]", Design.DisplayName, Design.Category);
            }
        }

        private class OnlineDisplayItem
        {
            public OnlineDesignItem Design { get; set; }
            public override string ToString()
            {
                return string.Format("{0} [{1}]", Design.Name, Design.Category);
            }
        }
    }
}
