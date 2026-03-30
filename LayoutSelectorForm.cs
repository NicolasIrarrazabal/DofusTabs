using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace DofusMiniTabber
{
    public partial class LayoutSelectorForm : Form
    {
        private readonly ListBox _layoutsListBox      = new();
        private readonly TextBox _nameTextBox         = new();
        private readonly TextBox _descriptionTextBox  = new();
        private readonly Button  _saveButton          = new();
        private readonly Button  _loadButton          = new();
        private readonly Button  _deleteButton        = new();
        private readonly Button  _newButton           = new();
        private readonly Button  _cancelButton        = new();
        private readonly Button  _setPreferredButton  = new();
        private readonly Label   _titleLabel          = new();
        private readonly Label   _nameLabel           = new();
        private readonly Label   _descriptionLabel    = new();
        private readonly Label   _preferredLabel      = new();

        public string? SelectedLayout          { get; private set; }
        public bool    ShouldLoad              { get; private set; }
        public bool    PreferredLayoutChanged  { get; private set; }
        public string? NewPreferredLayout      { get; private set; }

        private string? _currentPreferred;

        public LayoutSelectorForm() : this(null) { }

        public LayoutSelectorForm(string? currentPreferredLayout)
        {
            _currentPreferred  = currentPreferredLayout;
            NewPreferredLayout = currentPreferredLayout;
            InitializeComponents();
            LoadLayouts();
            UpdatePreferredLabel();
        }

        private void InitializeComponents()
        {
            Text            = "Gestor de Layouts - Wintabber Dofus";
            Size            = new Size(550, 510);
            StartPosition   = FormStartPosition.CenterParent;
            BackColor       = Color.FromArgb(0x0F, 0x19, 0x23);
            Font            = new Font("Segoe UI", 9F);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox     = false;
            MinimizeBox     = false;

            _titleLabel.Text      = "Gestionar Layouts de Ventanas";
            _titleLabel.ForeColor = Color.White;
            _titleLabel.Font      = new Font("Segoe UI", 14F, FontStyle.Bold);
            _titleLabel.Location  = new Point(20, 15);
            _titleLabel.Size      = new Size(400, 30);

            var instructionLabel = new Label
            {
                Text      = "Selecciona un layout de la lista y elige una acción:",
                ForeColor = Color.FromArgb(0x6C, 0x75, 0x7D),
                Location  = new Point(20, 50),
                Size      = new Size(400, 20),
                Font      = new Font("Segoe UI", 9F)
            };

            _layoutsListBox.Location             = new Point(20, 80);
            _layoutsListBox.Size                 = new Size(300, 240);
            _layoutsListBox.BackColor            = Color.FromArgb(0x1E, 0x2A, 0x38);
            _layoutsListBox.ForeColor            = Color.White;
            _layoutsListBox.BorderStyle          = BorderStyle.FixedSingle;
            _layoutsListBox.Font                 = new Font("Segoe UI", 9F);
            _layoutsListBox.SelectedIndexChanged += LayoutsListBox_SelectedIndexChanged;

            var selectionLabel = new Label
            {
                Text      = "Layout seleccionado:",
                ForeColor = Color.White,
                Location  = new Point(340, 80),
                Size      = new Size(120, 20),
                Font      = new Font("Segoe UI", 9F, FontStyle.Bold)
            };

            _nameLabel.Text      = "Nombre:";
            _nameLabel.ForeColor = Color.White;
            _nameLabel.Location  = new Point(340, 110);
            _nameLabel.Size      = new Size(60, 20);

            _nameTextBox.Location    = new Point(340, 135);
            _nameTextBox.Size        = new Size(180, 30);
            _nameTextBox.BackColor   = Color.FromArgb(0x1E, 0x2A, 0x38);
            _nameTextBox.ForeColor   = Color.White;
            _nameTextBox.BorderStyle = BorderStyle.FixedSingle;
            _nameTextBox.Font        = new Font("Segoe UI", 10F);

            _descriptionLabel.Text      = "Descripción:";
            _descriptionLabel.ForeColor = Color.White;
            _descriptionLabel.Location  = new Point(340, 175);
            _descriptionLabel.Size      = new Size(80, 20);

            _descriptionTextBox.Location    = new Point(340, 200);
            _descriptionTextBox.Size        = new Size(180, 60);
            _descriptionTextBox.BackColor   = Color.FromArgb(0x1E, 0x2A, 0x38);
            _descriptionTextBox.ForeColor   = Color.White;
            _descriptionTextBox.BorderStyle = BorderStyle.FixedSingle;
            _descriptionTextBox.Multiline   = true;
            _descriptionTextBox.Font        = new Font("Segoe UI", 9F);

            SetupButton(_newButton,    "NUEVO",    new Point(20,  415), Color.FromArgb(0x17, 0xA2, 0xB8));
            SetupButton(_saveButton,   "GUARDAR",  new Point(120, 415), Color.FromArgb(0x28, 0xA7, 0x45));
            SetupButton(_loadButton,   "CARGAR",   new Point(220, 415), Color.FromArgb(0x00, 0x7A, 0xCC));
            SetupButton(_deleteButton, "ELIMINAR", new Point(320, 415), Color.FromArgb(0xDC, 0x35, 0x45));
            SetupButton(_cancelButton, "CERRAR",   new Point(420, 415), Color.FromArgb(0x6C, 0x75, 0x7D));

            _newButton.Click    += NewButton_Click;
            _saveButton.Click   += SaveButton_Click;
            _loadButton.Click   += LoadButton_Click;
            _deleteButton.Click += DeleteButton_Click;
            _cancelButton.Click += CancelButton_Click;

            _preferredLabel.Text      = "⭐ Preferido: (ninguno)";
            _preferredLabel.ForeColor = Color.FromArgb(0xFF, 0xD7, 0x00);
            _preferredLabel.Location  = new Point(20, 335);
            _preferredLabel.Size      = new Size(300, 20);
            _preferredLabel.Font      = new Font("Segoe UI", 9F, FontStyle.Bold);

            _setPreferredButton.Text      = "⭐ MARCAR PREFERIDO";
            _setPreferredButton.Location  = new Point(20, 358);
            _setPreferredButton.Size      = new Size(200, 40);
            _setPreferredButton.BackColor = Color.FromArgb(0x85, 0x59, 0x00);
            _setPreferredButton.ForeColor = Color.FromArgb(0xFF, 0xD7, 0x00);
            _setPreferredButton.FlatStyle = FlatStyle.Flat;
            _setPreferredButton.Font      = new Font("Segoe UI", 10F, FontStyle.Bold);
            _setPreferredButton.FlatAppearance.BorderSize = 0;
            _setPreferredButton.Click += SetPreferredButton_Click;

            var clearPreferredButton = new Button
            {
                Text      = "✖ QUITAR PREFERIDO",
                Location  = new Point(228, 358),
                Size      = new Size(170, 40),
                BackColor = Color.FromArgb(0x3A, 0x2A, 0x00),
                ForeColor = Color.FromArgb(0xFF, 0xD7, 0x00),
                FlatStyle = FlatStyle.Flat,
                Font      = new Font("Segoe UI", 9F, FontStyle.Bold)
            };
            clearPreferredButton.FlatAppearance.BorderSize = 0;
            clearPreferredButton.Click += ClearPreferredButton_Click;

            Controls.AddRange(new Control[] {
                _titleLabel, instructionLabel, _layoutsListBox, selectionLabel,
                _nameLabel, _nameTextBox, _descriptionLabel, _descriptionTextBox,
                _preferredLabel, _setPreferredButton, clearPreferredButton,
                _newButton, _saveButton, _loadButton, _deleteButton, _cancelButton
            });
        }

        private static void SetupButton(Button b, string text, Point location, Color backColor)
        {
            b.Text      = text;
            b.Location  = location;
            b.Size      = new Size(90, 45);
            b.BackColor = backColor;
            b.ForeColor = Color.White;
            b.FlatStyle = FlatStyle.Flat;
            b.Font      = new Font("Segoe UI", 10F, FontStyle.Bold);
            b.FlatAppearance.BorderSize = 0;
        }

        private void LoadLayouts()
        {
            _layoutsListBox.Items.Clear();
            foreach (var config in WindowPositionManager.GetAllConfigurations().OrderByDescending(c => c.CreatedAt))
            {
                string displayText = string.IsNullOrEmpty(config.Description)
                    ? $"{config.Name} ({config.CreatedAt:dd/MM/yyyy HH:mm})"
                    : $"{config.Name} - {config.Description} ({config.CreatedAt:dd/MM/yyyy HH:mm})";
                _layoutsListBox.Items.Add(new LayoutItem(config.Name, displayText));
            }
        }

        private void LayoutsListBox_SelectedIndexChanged(object? sender, EventArgs e)
        {
            if (_layoutsListBox.SelectedItem is LayoutItem item)
            {
                var config = WindowPositionManager.LoadConfiguration(item.Name);
                if (config != null)
                {
                    _nameTextBox.Text        = config.Name;
                    _descriptionTextBox.Text = config.Description;
                }
            }
        }

        private void NewButton_Click(object? sender, EventArgs e)
        {
            _nameTextBox.Text        = "";
            _descriptionTextBox.Text = "";
            _nameTextBox.Focus();
        }

        private void SaveButton_Click(object? sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(_nameTextBox.Text))
            {
                MessageBox.Show("Por favor ingrese un nombre para el layout.", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            SelectedLayout = _nameTextBox.Text;
            ShouldLoad     = false;
            DialogResult   = DialogResult.OK;
            Close();
        }

        private void LoadButton_Click(object? sender, EventArgs e)
        {
            if (_layoutsListBox.SelectedItem is LayoutItem item)
            {
                SelectedLayout = item.Name;
                ShouldLoad     = true;
                DialogResult   = DialogResult.OK;
                Close();
            }
            else
            {
                MessageBox.Show("Por favor selecciona un layout de la lista para cargar.",
                    "Selección Requerida", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void DeleteButton_Click(object? sender, EventArgs e)
        {
            if (_layoutsListBox.SelectedItem is LayoutItem item)
            {
                var result = MessageBox.Show(
                    $"¿Estás seguro que deseas eliminar el layout '{item.Name}'?\n\nEsta acción no se puede deshacer.",
                    "Confirmar Eliminación", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

                if (result == DialogResult.Yes)
                {
                    try
                    {
                        WindowPositionManager.DeleteConfiguration(item.Name);
                        LoadLayouts();
                        _nameTextBox.Text        = "";
                        _descriptionTextBox.Text = "";
                        MessageBox.Show($"Layout '{item.Name}' eliminado correctamente.",
                            "Eliminación Exitosa", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Error al eliminar layout: {ex.Message}",
                            "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
            else
            {
                MessageBox.Show("Por favor selecciona un layout de la lista para eliminar.",
                    "Selección Requerida", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void CancelButton_Click(object? sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }

        private void SetPreferredButton_Click(object? sender, EventArgs e)
        {
            if (_layoutsListBox.SelectedItem is LayoutItem item)
            {
                _currentPreferred      = item.Name;
                NewPreferredLayout     = item.Name;
                PreferredLayoutChanged = true;
                UpdatePreferredLabel();
                MessageBox.Show(
                    $"Layout '{item.Name}' marcado como preferido.\n\nUsa el botón '⭐ CARGAR PREFERIDO' en el menú principal para cargarlo.",
                    "Layout Preferido", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("Selecciona un layout de la lista primero.",
                    "Selección Requerida", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void ClearPreferredButton_Click(object? sender, EventArgs e)
        {
            _currentPreferred      = null;
            NewPreferredLayout     = null;
            PreferredLayoutChanged = true;
            UpdatePreferredLabel();
        }

        private void UpdatePreferredLabel()
        {
            _preferredLabel.Text = string.IsNullOrWhiteSpace(_currentPreferred)
                ? "⭐ Preferido: (ninguno)"
                : $"⭐ Preferido: {_currentPreferred}";
        }

        private class LayoutItem(string name, string displayText)
        {
            public string Name        { get; } = name;
            public string DisplayText { get; } = displayText;
            public override string ToString() => DisplayText;
        }
    }
}
