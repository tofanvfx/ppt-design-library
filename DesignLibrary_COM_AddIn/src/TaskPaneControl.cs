using System;
using System.Drawing;
using System.Linq;
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
        private ListBox _lstDesigns;
        private ComboBox _cboCategory;
        private TextBox _txtName;
        private TextBox _txtCategory;
        private Label _lblStatus;
        private Button _btnSave;
        private Button _btnInsert;
        private Button _btnDelete;
        private Button _btnRename;
        private PictureBox _picPreview;

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

        private void InitializeComponent()
        {
            this.Size = new Size(300, 600);
            this.BackColor = ColorTranslator.FromHtml("#FFFFFF");
            this.AutoScroll = true;
            this.Font = new Font("Segoe UI", 9F, FontStyle.Regular);

            int y = 15;
            int margin = 15;
            int width = 270;

            Color primaryBlue = ColorTranslator.FromHtml("#0078D4");
            Color primaryGreen = ColorTranslator.FromHtml("#107C41");
            Color dangerRed = ColorTranslator.FromHtml("#A4262C");
            Color lightGray = ColorTranslator.FromHtml("#F3F2F1");
            Color borderGray = ColorTranslator.FromHtml("#EDEBE9");
            Color textDark = ColorTranslator.FromHtml("#323130");
            Color textLight = ColorTranslator.FromHtml("#605E5C");

            var lblTitle = new Label { Text = "DESIGN LIBRARY", Location = new Point(margin, y), Size = new Size(width, 25), Font = new Font("Segoe UI Semibold", 12F), ForeColor = textDark };
            this.Controls.Add(lblTitle);
            y += 25;

            var lblSub = new Label { Text = "Save and reuse slide objects", Location = new Point(margin, y), Size = new Size(width, 15), ForeColor = textLight };
            this.Controls.Add(lblSub);
            y += 25;

            var divider1 = new Label { BackColor = borderGray, Location = new Point(margin, y), Size = new Size(width, 1) };
            this.Controls.Add(divider1);
            y += 15;

            this.Controls.Add(new Label { Text = "Filter:", Location = new Point(margin, y + 4), AutoSize = true, ForeColor = textDark });
            
            _cboCategory = new ComboBox { Location = new Point(55, y), Size = new Size(140, 25), DropDownStyle = ComboBoxStyle.DropDownList, FlatStyle = FlatStyle.Flat, BackColor = lightGray };
            _cboCategory.SelectedIndexChanged += (s, e) => RefreshList();
            this.Controls.Add(_cboCategory);

            var btnRefresh = new Button { Text = "Refresh", Location = new Point(205, y - 1), Size = new Size(80, 26), FlatStyle = FlatStyle.Flat, BackColor = lightGray, ForeColor = textDark };
            btnRefresh.FlatAppearance.BorderSize = 0;
            btnRefresh.Click += (s, e) => { RefreshCategories(); RefreshList(); };
            this.Controls.Add(btnRefresh);
            y += 35;

            _lstDesigns = new ListBox { Location = new Point(margin, y), Size = new Size(width, 120), Font = new Font("Segoe UI", 9F), BorderStyle = BorderStyle.FixedSingle, BackColor = Color.White, ForeColor = textDark };
            _lstDesigns.SelectedIndexChanged += LstDesigns_SelectedIndexChanged;
            this.Controls.Add(_lstDesigns);
            y += 130;

            _picPreview = new PictureBox { Location = new Point(margin, y), Size = new Size(width, 120), SizeMode = PictureBoxSizeMode.Zoom, BorderStyle = BorderStyle.FixedSingle, BackColor = Color.White };
            this.Controls.Add(_picPreview);
            y += 130;

            _btnInsert = new Button { Text = "Insert Design", Location = new Point(margin, y), Size = new Size(width, 36), FlatStyle = FlatStyle.Flat, BackColor = primaryBlue, ForeColor = Color.White, Font = new Font("Segoe UI Semibold", 9.5F) };
            _btnInsert.FlatAppearance.BorderSize = 0;
            _btnInsert.Click += BtnInsert_Click;
            this.Controls.Add(_btnInsert);
            y += 42;

            _btnRename = new Button { Text = "Rename", Location = new Point(margin, y), Size = new Size(width, 32), FlatStyle = FlatStyle.Flat, BackColor = lightGray, ForeColor = textDark };
            _btnRename.FlatAppearance.BorderSize = 0;
            _btnRename.Click += BtnRename_Click;
            this.Controls.Add(_btnRename);
            y += 38;

            _btnDelete = new Button { Text = "Delete", Location = new Point(margin, y), Size = new Size(width, 32), FlatStyle = FlatStyle.Flat, BackColor = lightGray, ForeColor = dangerRed };
            _btnDelete.FlatAppearance.BorderSize = 0;
            _btnDelete.Click += BtnDelete_Click;
            this.Controls.Add(_btnDelete);
            y += 45;

            var divider2 = new Label { BackColor = borderGray, Location = new Point(margin, y), Size = new Size(width, 1) };
            this.Controls.Add(divider2);
            y += 15;

            this.Controls.Add(new Label { Text = "SAVE SELECTED SHAPES", Location = new Point(margin, y), Size = new Size(width, 20), Font = new Font("Segoe UI Semibold", 9.5F), ForeColor = textDark });
            y += 25;

            this.Controls.Add(new Label { Text = "Name:", Location = new Point(margin, y + 4), AutoSize = true, ForeColor = textDark });
            _txtName = new TextBox { Location = new Point(80, y), Size = new Size(190, 25), BorderStyle = BorderStyle.FixedSingle };
            this.Controls.Add(_txtName);
            y += 32;

            this.Controls.Add(new Label { Text = "Category:", Location = new Point(margin, y + 4), AutoSize = true, ForeColor = textDark });
            _txtCategory = new TextBox { Location = new Point(80, y), Size = new Size(190, 25), Text = "General", BorderStyle = BorderStyle.FixedSingle };
            this.Controls.Add(_txtCategory);
            y += 38;

            _btnSave = new Button { Text = "Save Current Selection", Location = new Point(margin, y), Size = new Size(width, 36), FlatStyle = FlatStyle.Flat, BackColor = primaryGreen, ForeColor = Color.White, Font = new Font("Segoe UI Semibold", 9.5F) };
            _btnSave.FlatAppearance.BorderSize = 0;
            _btnSave.Click += BtnSave_Click;
            this.Controls.Add(_btnSave);
            y += 45;

            _lblStatus = new Label { Location = new Point(margin, y), Size = new Size(width, 40), ForeColor = textLight };
            this.Controls.Add(_lblStatus);
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            if (_manager == null) return;
            try
            {
                string name = _txtName.Text.Trim();
                if (string.IsNullOrEmpty(name))
                {
                    Status("Please enter a design name.", true);
                    return;
                }
                string category = _txtCategory.Text.Trim();
                if (string.IsNullOrEmpty(category)) category = "General";

                Status("Saving...");
                _manager.SaveSelectedAsDesign(name, category);
                _txtName.Text = "";
                Status("Saved!");
                RefreshCategories();
                RefreshList();
            }
            catch (Exception ex)
            {
                Status("Error: " + ex.Message, true);
            }
        }

        private void BtnInsert_Click(object sender, EventArgs e)
        {
            if (_manager == null) return;
            try
            {
                DisplayItem item = _lstDesigns.SelectedItem as DisplayItem;
                if (item != null)
                {
                    Status("Inserting...");
                    _manager.InsertDesign(item.Design.FileName);
                    Status("Inserted!");
                }
                else
                {
                    Status("Select a design first.", true);
                }
            }
            catch (Exception ex)
            {
                Status("Error: " + ex.Message, true);
            }
        }

        private void LstDesigns_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_lstDesigns.SelectedItem == null)
            {
                if (_picPreview.Image != null)
                {
                    _picPreview.Image.Dispose();
                    _picPreview.Image = null;
                }
                return;
            }

            var selected = (DisplayItem)_lstDesigns.SelectedItem;
            string pngPath = System.IO.Path.Combine(_manager.GetLibraryPath(), selected.Design.FileName.Replace(".pptx", ".png"));

            if (_picPreview.Image != null)
            {
                _picPreview.Image.Dispose();
                _picPreview.Image = null;
            }

            if (System.IO.File.Exists(pngPath))
            {
                try
                {
                    using (var fs = new System.IO.FileStream(pngPath, System.IO.FileMode.Open, System.IO.FileAccess.Read))
                    {
                        var img = System.Drawing.Image.FromStream(fs);
                        _picPreview.Image = new System.Drawing.Bitmap(img);
                    }
                }
                catch { }
            }
        }

        private void BtnRename_Click(object sender, EventArgs e)
        {
            if (_manager == null) return;
            DisplayItem item = _lstDesigns.SelectedItem as DisplayItem;
            if (item != null)
            {
                string newName = Microsoft.VisualBasic.Interaction.InputBox("New name:", "Rename Design", item.Design.DisplayName);
                if (!string.IsNullOrWhiteSpace(newName))
                {
                    _manager.RenameDesign(item.Design.FileName, newName.Trim());
                    RefreshList();
                    Status("Renamed.");
                }
            }
            else
            {
                Status("Select a design first.", true);
            }
        }

        private void BtnDelete_Click(object sender, EventArgs e)
        {
            if (_manager == null) return;
            DisplayItem item = _lstDesigns.SelectedItem as DisplayItem;
            if (item != null)
            {
                var result = MessageBox.Show(string.Format("Delete '{0}'?", item.Design.DisplayName), "Confirm Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (result == DialogResult.Yes)
                {
                    _manager.DeleteDesign(item.Design.FileName);
                    RefreshCategories();
                    RefreshList();
                    Status("Deleted.");
                }
            }
            else
            {
                Status("Select a design first.", true);
            }
        }

        private void RefreshCategories()
        {
            if (_manager == null) return;
            string current = _cboCategory.SelectedItem != null ? _cboCategory.SelectedItem.ToString() : null;
            _cboCategory.Items.Clear();
            _cboCategory.Items.Add("All Categories");
            
            var designs = _manager.GetAllDesigns();
            var categories = designs.Select(d => d.Category).Distinct().OrderBy(c => c).ToList();
            
            foreach (var cat in categories)
            {
                _cboCategory.Items.Add(cat);
            }

            if (current != null && _cboCategory.Items.Contains(current))
            {
                _cboCategory.SelectedItem = current;
            }
            else
            {
                _cboCategory.SelectedIndex = 0;
            }
        }

        private void RefreshList()
        {
            if (_manager == null) return;
            _lstDesigns.Items.Clear();
            string selectedCat = _cboCategory.SelectedIndex > 0 ? _cboCategory.SelectedItem.ToString() : null;
            
            var designs = _manager.GetAllDesigns();
            
            foreach (var d in designs)
            {
                if (selectedCat != null && d.Category != selectedCat) continue;
                _lstDesigns.Items.Add(new DisplayItem { Design = d });
            }

            if (_lstDesigns.Items.Count == 0)
                Status("No designs yet. Select shapes and save!");
            else
                Status(string.Format("{0} design(s).", _lstDesigns.Items.Count));
        }

        private void Status(string msg, bool isError = false)
        {
            _lblStatus.Text = msg;
            _lblStatus.ForeColor = isError ? Color.DarkRed : Color.Gray;
        }

        private class DisplayItem
        {
            public DesignItem Design { get; set; }
            public override string ToString()
            {
                return string.Format("{0} [{1}]", Design.DisplayName, Design.Category);
            }
        }
    }
}
