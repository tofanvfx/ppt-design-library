using System;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace DesignLibraryAddIn
{
    public class LibraryForm : Form
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

        public LibraryForm(DesignManager manager)
        {
            _manager = manager;
            InitializeComponent();
            RefreshCategories();
            RefreshList();
        }

        private void InitializeComponent()
        {
            this.Text = "Design Library";
            this.Size = new Size(400, 560);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            int y = 15;

            var lblTitle = new Label { Text = "DESIGN LIBRARY", Location = new Point(15, y), Size = new Size(350, 25), Font = new Font("Segoe UI", 12, FontStyle.Bold), ForeColor = Color.SaddleBrown };
            this.Controls.Add(lblTitle);
            y += 25;

            var lblSub = new Label { Text = "Save and reuse slide design elements", Location = new Point(15, y), Size = new Size(350, 15), ForeColor = Color.Gray };
            this.Controls.Add(lblSub);
            y += 25;

            this.Controls.Add(new Label { BorderStyle = BorderStyle.Fixed3D, Location = new Point(15, y), Size = new Size(350, 2) });
            y += 15;

            this.Controls.Add(new Label { Text = "Filter:", Location = new Point(15, y + 2), Size = new Size(40, 20), Font = new Font(this.Font, FontStyle.Bold) });
            
            _cboCategory = new ComboBox { Location = new Point(60, y), Size = new Size(210, 25), DropDownStyle = ComboBoxStyle.DropDownList };
            _cboCategory.SelectedIndexChanged += (s, e) => RefreshList();
            this.Controls.Add(_cboCategory);

            var btnRefresh = new Button { Text = "Refresh", Location = new Point(280, y - 1), Size = new Size(80, 25) };
            btnRefresh.Click += (s, e) => { RefreshCategories(); RefreshList(); };
            this.Controls.Add(btnRefresh);
            y += 35;

            _lstDesigns = new ListBox { Location = new Point(15, y), Size = new Size(345, 180), Font = new Font("Segoe UI", 9) };
            this.Controls.Add(_lstDesigns);
            y += 190;

            _btnInsert = new Button { Text = "Insert", Location = new Point(15, y), Size = new Size(110, 30), Font = new Font(this.Font, FontStyle.Bold), BackColor = Color.Chocolate, ForeColor = Color.White };
            _btnInsert.Click += BtnInsert_Click;
            this.Controls.Add(_btnInsert);

            _btnRename = new Button { Text = "Rename", Location = new Point(135, y), Size = new Size(100, 30) };
            _btnRename.Click += BtnRename_Click;
            this.Controls.Add(_btnRename);

            _btnDelete = new Button { Text = "Delete", Location = new Point(245, y), Size = new Size(115, 30), ForeColor = Color.DarkRed };
            _btnDelete.Click += BtnDelete_Click;
            this.Controls.Add(_btnDelete);
            y += 45;

            this.Controls.Add(new Label { BorderStyle = BorderStyle.Fixed3D, Location = new Point(15, y), Size = new Size(350, 2) });
            y += 15;

            this.Controls.Add(new Label { Text = "SAVE SELECTED SHAPES", Location = new Point(15, y), Size = new Size(350, 20), Font = new Font("Segoe UI", 9, FontStyle.Bold), ForeColor = Color.SaddleBrown });
            y += 25;

            this.Controls.Add(new Label { Text = "Name:", Location = new Point(15, y + 2), Size = new Size(50, 20) });
            _txtName = new TextBox { Location = new Point(70, y), Size = new Size(290, 25) };
            this.Controls.Add(_txtName);
            y += 30;

            this.Controls.Add(new Label { Text = "Category:", Location = new Point(15, y + 2), Size = new Size(60, 20) });
            _txtCategory = new TextBox { Location = new Point(80, y), Size = new Size(280, 25), Text = "General" };
            this.Controls.Add(_txtCategory);
            y += 35;

            _btnSave = new Button { Text = "Save Current Selection", Location = new Point(15, y), Size = new Size(345, 35), Font = new Font("Segoe UI", 10, FontStyle.Bold), BackColor = Color.MediumSeaGreen, ForeColor = Color.White };
            _btnSave.Click += BtnSave_Click;
            this.Controls.Add(_btnSave);
            y += 45;

            _lblStatus = new Label { Location = new Point(15, y), Size = new Size(345, 20), ForeColor = Color.Gray };
            this.Controls.Add(_lblStatus);
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
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

        private void BtnRename_Click(object sender, EventArgs e)
        {
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
