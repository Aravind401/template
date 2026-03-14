using ClosedXML.Excel;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;
using System.ComponentModel;
using System.Globalization;
using System.IO;

namespace QuotationTemplateApp;

public class MainForm : Form
{
    private readonly DataGridView _itemsGrid = new();

    private readonly TextBox _txtCompany = new() { Text = "R.R Engineering" };
    private readonly TextBox _txtGstin = new() { Text = "33CZGPR1438E1ZI" };
    private readonly TextBox _txtCompanyAddress = new() { Text = "63/1 Mahaveer Street, Chennai, Tamil Nadu - 600050" };
    private readonly TextBox _txtCompanyPhone = new() { Text = "+91 90924 92393" };
    private readonly TextBox _txtCompanyEmail = new() { Text = "rajadhurai1998@gmail.com" };

    private readonly TextBox _txtCustomer = new();
    private readonly TextBox _txtCustomerAddress = new();
    private readonly TextBox _txtPhone = new();
    private readonly TextBox _txtSupplyPlace = new() { Text = "Tamil Nadu" };

    private readonly TextBox _txtQuotationNo = new() { Text = "EST-15" };
    private readonly DateTimePicker _quoteDate = new() { Value = DateTime.Today, Format = DateTimePickerFormat.Short };
    private readonly DateTimePicker _validityDate = new() { Value = DateTime.Today.AddDays(7), Format = DateTimePickerFormat.Short };

    private readonly NumericUpDown _gstPercent = new() { DecimalPlaces = 2, Minimum = 0, Maximum = 100, Value = 18 };
    private readonly TextBox _txtSubTotal = new() { ReadOnly = true };
    private readonly TextBox _txtGstAmount = new() { ReadOnly = true };
    private readonly TextBox _txtGrandTotal = new() { ReadOnly = true };
    private readonly TextBox _txtAmountWords = new() { ReadOnly = true, Multiline = true, Height = 55 };

    private readonly Button _btnAddRow = new() { Text = "Add" };
    private readonly Button _btnDeleteRow = new() { Text = "Delete" };
    private readonly Button _btnRecalculate = new() { Text = "Recalculate" };
    private readonly Button _btnSaveData = new() { Text = "Save Data" };
    private readonly Button _btnExportExcel = new() { Text = "Export Excel" };
    private readonly Button _btnExportPdf = new() { Text = "Export PDF" };
    private readonly Button _btnOwnerDetails = new() { Text = "Owner Details" };

    private readonly Label _lblOwnerCompany = new() { AutoEllipsis = true, Dock = DockStyle.Fill };
    private readonly Label _lblOwnerGstin = new() { AutoEllipsis = true, Dock = DockStyle.Fill };
    private readonly Label _lblOwnerPhone = new() { AutoEllipsis = true, Dock = DockStyle.Fill };
    private readonly Label _lblOwnerEmail = new() { AutoEllipsis = true, Dock = DockStyle.Fill };

    private readonly PictureBox _logoPreview = new() { SizeMode = PictureBoxSizeMode.Zoom, Dock = DockStyle.Fill };
    private readonly Button _btnUploadLogo = new() { Text = "Upload Logo" };

    private readonly AppDataStore _dataStore;
    private readonly string _dataDirectory;
    private string? _logoPath;
    private bool _isLoadingData;
    private List<string> _materialSuggestions = [];

    public MainForm()
    {
        Text = "Quotation Template (WinForms)";
        Width = 1320;
        Height = 860;
        StartPosition = FormStartPosition.CenterScreen;

        _dataDirectory = Path.Combine(AppContext.BaseDirectory, "app-data");
        _dataStore = new AppDataStore(_dataDirectory);

        BuildLayout();
        ConfigureGrid();
        WireEvents();
        ConfigureButtons();

        LoadSavedState();
        RecalculateTotals();
    }

    private void BuildLayout()
    {
        var root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            RowCount = 4,
            ColumnCount = 1,
            Padding = new Padding(10)
        };

        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 300));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 160));
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 70));

        var topPanel = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3, Padding = new Padding(0, 0, 0, 6) };
        topPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 120));
        topPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 45));
        topPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 130));

        topPanel.Controls.Add(BuildOwnerAndCustomerPanel(), 0, 0);
        topPanel.Controls.Add(BuildQuotationInfoPanel(), 0, 1);
        topPanel.Controls.Add(BuildCustomerPanel(), 0, 2);

        var gridHost = new Panel { Dock = DockStyle.Fill };
        _itemsGrid.Dock = DockStyle.Fill;
        gridHost.Controls.Add(_itemsGrid);

        var totals = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 6,
            RowCount = 3,
            Padding = new Padding(0, 10, 0, 0)
        };

        totals.RowStyles.Add(new RowStyle(SizeType.Absolute, 40));
        totals.RowStyles.Add(new RowStyle(SizeType.Absolute, 40));
        totals.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

        totals.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100));
        totals.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.33F));
        totals.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100));
        totals.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.33F));
        totals.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100));
        totals.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.33F));

        AddField(totals, "GST %", _gstPercent, 0, 0);
        AddField(totals, "Sub Total", _txtSubTotal, 2, 0);
        AddField(totals, "GST Amount", _txtGstAmount, 4, 0);
        AddField(totals, "TOTAL AMOUNT", _txtGrandTotal, 2, 1);
        AddField(totals, "Total Amount (in words)", _txtAmountWords, 0, 2, 6);

        _txtGrandTotal.Font = new Font("Segoe UI", 11, FontStyle.Bold);
        _txtAmountWords.MinimumSize = new Size(0, 60);
        _txtAmountWords.Dock = DockStyle.Fill;

        var actions = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = false,
            Padding = new Padding(0, 5, 0, 0)
        };

        actions.Controls.AddRange(new Control[] { _btnOwnerDetails, _btnAddRow, _btnDeleteRow, _btnRecalculate, _btnSaveData, _btnExportExcel, _btnExportPdf });

        root.Controls.Add(topPanel, 0, 0);
        root.Controls.Add(gridHost, 0, 1);
        root.Controls.Add(totals, 0, 2);
        root.Controls.Add(actions, 0, 3);

        Controls.Add(root);
    }

    private Control BuildOwnerAndCustomerPanel()
    {
        var panel = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1, Margin = new Padding(0, 0, 0, 8) };
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 75));
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25));

        var owner = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 2,
            RowCount = 6,
            Padding = new Padding(8),
            BackColor = Color.WhiteSmoke
        };
        owner.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 170));
        owner.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        owner.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
        owner.RowStyles.Add(new RowStyle(SizeType.Absolute, 26));
        owner.RowStyles.Add(new RowStyle(SizeType.Absolute, 26));
        owner.RowStyles.Add(new RowStyle(SizeType.Absolute, 26));
        owner.RowStyles.Add(new RowStyle(SizeType.Absolute, 26));
        owner.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

        var title = new Label
        {
            Text = "Owner Configuration",
            Dock = DockStyle.Fill,
            Font = new Font("Segoe UI", 10, FontStyle.Bold),
            TextAlign = ContentAlignment.MiddleLeft
        };
        owner.Controls.Add(title, 0, 0);
        owner.SetColumnSpan(title, 2);
        AddSummaryField(owner, "Company", _lblOwnerCompany, 1);
        AddSummaryField(owner, "GSTIN", _lblOwnerGstin, 2);
        AddSummaryField(owner, "Phone", _lblOwnerPhone, 3);
        AddSummaryField(owner, "Email", _lblOwnerEmail, 4);

        var logoPanel = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            BorderStyle = BorderStyle.FixedSingle,
            RowCount = 2,
            ColumnCount = 1,
            Padding = new Padding(6)
        };
        logoPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        logoPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 44));

        _logoPreview.Image = null;
        _logoPreview.BackColor = Color.WhiteSmoke;

        _btnUploadLogo.Dock = DockStyle.Fill;
        _btnUploadLogo.Font = new Font("Segoe UI", 9, FontStyle.Bold);

        logoPanel.Controls.Add(_logoPreview, 0, 0);
        logoPanel.Controls.Add(_btnUploadLogo, 0, 1);

        panel.Controls.Add(owner, 0, 0);
        panel.Controls.Add(logoPanel, 1, 0);

        return panel;
    }

    private static void AddSummaryField(TableLayoutPanel panel, string label, Label value, int row)
    {
        panel.Controls.Add(new Label
        {
            Text = label,
            Dock = DockStyle.Fill,
            Font = new Font("Segoe UI", 9, FontStyle.Bold),
            TextAlign = ContentAlignment.MiddleLeft
        }, 0, row);

        value.TextAlign = ContentAlignment.MiddleLeft;
        value.Font = new Font("Segoe UI", 9);
        panel.Controls.Add(value, 1, row);
    }

    private Control BuildQuotationInfoPanel()
    {
        var panel = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 6, RowCount = 1, Margin = new Padding(0, 0, 0, 8) };
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 120));
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.33F));
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 120));
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.33F));
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 70));
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.33F));

        AddField(panel, "Quotation No", _txtQuotationNo, 0, 0);
        AddField(panel, "Quotation Date", _quoteDate, 2, 0);
        AddField(panel, "Validity", _validityDate, 4, 0);
        return panel;
    }

    private Control BuildCustomerPanel()
    {
        _txtCustomerAddress.Multiline = true;
        _txtCustomerAddress.ScrollBars = ScrollBars.Vertical;
        _txtPhone.MinimumSize = new Size(0, 32);

        var panel = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 5 };
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 130));
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
        panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 34));
        panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 34));
        panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 50));
        panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 34));

        panel.Controls.Add(new Label
        {
            Text = "Customer Details",
            Dock = DockStyle.Fill,
            Font = new Font("Segoe UI", 10, FontStyle.Bold),
            TextAlign = ContentAlignment.MiddleLeft
        }, 0, 0);
        panel.SetColumnSpan(panel.GetControlFromPosition(0, 0), 2);

        AddField(panel, "Name", _txtCustomer, 0, 1);
        AddField(panel, "Phone No", _txtPhone, 0, 2);
        AddField(panel, "Address", _txtCustomerAddress, 0, 3);
        AddField(panel, "Place of Delivery", _txtSupplyPlace, 0, 4);

        return panel;
    }

    private void ConfigureButtons()
    {
        var buttons = new[] { _btnAddRow, _btnDeleteRow, _btnRecalculate, _btnSaveData, _btnExportExcel, _btnExportPdf };
        _btnOwnerDetails.Width = 180;
        _btnOwnerDetails.Height = 38;
        _btnOwnerDetails.Font = new Font("Segoe UI", 10, FontStyle.Bold);
        _btnOwnerDetails.Margin = new Padding(0, 0, 10, 0);

        foreach (var button in buttons)
        {
            button.Width = 160;
            button.Height = 38;
            button.Font = new Font("Segoe UI", 10, FontStyle.Bold);
            button.Margin = new Padding(0, 0, 10, 0);
        }
    }

    private static void AddField(TableLayoutPanel panel, string label, Control editor, int column, int row, int columnSpan = 2)
    {
        var caption = new Label
        {
            Text = label,
            Dock = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleLeft,
            Font = new Font("Segoe UI", 9, FontStyle.Bold)
        };

        if (editor is TextBox textBox && !textBox.Multiline)
        {
            editor.Anchor = AnchorStyles.Left | AnchorStyles.Right;
            textBox.Margin = new Padding(3, 6, 3, 6);
        }
        else
        {
            editor.Dock = DockStyle.Fill;
        }

        panel.Controls.Add(caption, column, row);
        panel.Controls.Add(editor, column + 1, row);

        if (columnSpan > 2)
        {
            panel.SetColumnSpan(editor, columnSpan - 1);
        }
    }

    private void ConfigureGrid()
    {
        _itemsGrid.AutoGenerateColumns = false;
        _itemsGrid.AllowUserToAddRows = false;
        _itemsGrid.AllowUserToDeleteRows = false;
        _itemsGrid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;

        _itemsGrid.Columns.Add(new DataGridViewTextBoxColumn { Name = "No", HeaderText = "NO", Width = 50, ReadOnly = true });
        _itemsGrid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Item", HeaderText = "Material Name", Width = 320 });
        _itemsGrid.Columns.Add(new DataGridViewTextBoxColumn { Name = "W", HeaderText = "W", Width = 80 });
        _itemsGrid.Columns.Add(new DataGridViewTextBoxColumn { Name = "H", HeaderText = "H", Width = 80 });
        _itemsGrid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Qty", HeaderText = "Qty", Width = 80 });
        _itemsGrid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Soft", HeaderText = "Soft", Width = 80 });
        _itemsGrid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Rate", HeaderText = "Rate", Width = 120 });
        _itemsGrid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Amount", HeaderText = "Amount", Width = 120, ReadOnly = true });
    }

    private void WireEvents()
    {
        _btnAddRow.Click += (_, _) => { AddDefaultRow(); SaveState(); };
        _btnDeleteRow.Click += (_, _) => { DeleteSelectedRow(); SaveState(); };
        _btnRecalculate.Click += (_, _) => RecalculateTotals();
        _btnSaveData.Click += (_, _) =>
        {
            SaveState();
            MessageBox.Show("Data saved successfully.", "Save", MessageBoxButtons.OK, MessageBoxIcon.Information);
        };
        _btnUploadLogo.Click += (_, _) => UploadLogo();
        _btnOwnerDetails.Click += (_, _) => ShowOwnerDetailsDialog();
        _btnExportExcel.Click += (_, _) => ExportExcel();
        _btnExportPdf.Click += (_, _) => ExportPdf();
        _gstPercent.ValueChanged += (_, _) =>
        {
            RecalculateTotals();
            SaveState();
        };

        _itemsGrid.EditingControlShowing += OnItemsGridEditingControlShowing;
        _itemsGrid.CellEndEdit += (_, e) =>
        {
            if (e.RowIndex >= 0)
            {
                UpdateRowAmount(e.RowIndex);
                RefreshMaterialSuggestions();
                RecalculateTotals();
                SaveState();
            }
        };

        foreach (var control in new Control[]
        {
            _txtCompany, _txtGstin, _txtCompanyAddress, _txtCompanyPhone, _txtCompanyEmail,
            _txtCustomer, _txtCustomerAddress, _txtPhone, _txtSupplyPlace, _txtQuotationNo
        })
        {
            control.TextChanged += (_, _) => SaveState();
        }

        _quoteDate.ValueChanged += (_, _) => SaveState();
        _validityDate.ValueChanged += (_, _) => SaveState();
    }

    private void LoadSavedState()
    {
        _isLoadingData = true;

        var state = _dataStore.LoadState();

        _txtCompany.Text = state.Owner.Company;
        _txtGstin.Text = state.Owner.Gstin;
        _txtCompanyAddress.Text = state.Owner.Address;
        _txtCompanyPhone.Text = state.Owner.Phone;
        _txtCompanyEmail.Text = state.Owner.Email;

        _txtCustomer.Text = state.Quotation.CustomerName;
        _txtCustomerAddress.Text = state.Quotation.CustomerAddress;
        _txtPhone.Text = state.Quotation.CustomerPhone;
        _txtSupplyPlace.Text = state.Quotation.SupplyPlace;
        _txtQuotationNo.Text = state.Quotation.QuotationNo;
        _quoteDate.Value = state.Quotation.QuoteDate;
        _validityDate.Value = state.Quotation.ValidityDate;
        _gstPercent.Value = Math.Clamp(state.Quotation.GstPercent, _gstPercent.Minimum, _gstPercent.Maximum);

        _logoPath = state.Owner.LogoPath;
        LoadLogoPreview();

        _itemsGrid.Rows.Clear();
        if (state.Items.Count == 0)
        {
            for (var i = 0; i < 18; i++)
            {
                AddDefaultRow();
            }
        }
        else
        {
            foreach (var item in state.Items.OrderBy(i => i.RowNo))
            {
                AddItemRow(item);
            }
        }

        RefreshMaterialSuggestions();
        RefreshOwnerSummary();

        _isLoadingData = false;
    }

    private void AddItemRow(QuoteItemData item)
    {
        var index = _itemsGrid.Rows.Add();
        var row = _itemsGrid.Rows[index];
        row.Cells["No"].Value = (index + 1).ToString(CultureInfo.InvariantCulture);
        row.Cells["Item"].Value = item.MaterialName;
        row.Cells["W"].Value = item.W;
        row.Cells["H"].Value = item.H;
        row.Cells["Qty"].Value = item.Qty;
        row.Cells["Soft"].Value = item.Soft;
        row.Cells["Rate"].Value = item.Rate;
        UpdateRowAmount(index);
    }

    private void SaveState()
    {
        if (_isLoadingData)
        {
            return;
        }

        var owner = new OwnerData(
            _txtCompany.Text.Trim(),
            _txtGstin.Text.Trim(),
            _txtCompanyAddress.Text.Trim(),
            _txtCompanyPhone.Text.Trim(),
            _txtCompanyEmail.Text.Trim(),
            _logoPath);

        var quotation = new QuotationData(
            _txtCustomer.Text.Trim(),
            _txtCustomerAddress.Text.Trim(),
            _txtPhone.Text.Trim(),
            _txtSupplyPlace.Text.Trim(),
            _txtQuotationNo.Text.Trim(),
            _quoteDate.Value.Date,
            _validityDate.Value.Date,
            _gstPercent.Value);

        var items = _itemsGrid.Rows
            .Cast<DataGridViewRow>()
            .Where(r => !r.IsNewRow)
            .Select((r, index) => new QuoteItemData(
                index + 1,
                r.Cells["Item"].Value?.ToString() ?? string.Empty,
                r.Cells["W"].Value?.ToString() ?? string.Empty,
                r.Cells["H"].Value?.ToString() ?? string.Empty,
                r.Cells["Qty"].Value?.ToString() ?? string.Empty,
                r.Cells["Soft"].Value?.ToString() ?? string.Empty,
                r.Cells["Rate"].Value?.ToString() ?? string.Empty))
            .ToList();

        _dataStore.SaveState(new AppState(owner, quotation, items));
        RefreshOwnerSummary();
    }

    private void RefreshOwnerSummary()
    {
        _lblOwnerCompany.Text = _txtCompany.Text;
        _lblOwnerGstin.Text = _txtGstin.Text;
        _lblOwnerPhone.Text = _txtCompanyPhone.Text;
        _lblOwnerEmail.Text = _txtCompanyEmail.Text;
    }

    private void ShowOwnerDetailsDialog()
    {
        using var dialog = new OwnerDetailsForm(new OwnerData(
            _txtCompany.Text,
            _txtGstin.Text,
            _txtCompanyAddress.Text,
            _txtCompanyPhone.Text,
            _txtCompanyEmail.Text,
            _logoPath));

        if (dialog.ShowDialog(this) != DialogResult.OK)
        {
            return;
        }

        var owner = dialog.OwnerDetails;
        _txtCompany.Text = owner.Company;
        _txtGstin.Text = owner.Gstin;
        _txtCompanyAddress.Text = owner.Address;
        _txtCompanyPhone.Text = owner.Phone;
        _txtCompanyEmail.Text = owner.Email;
        _logoPath = owner.LogoPath;
        LoadLogoPreview();
        SaveState();
    }

    private void UploadLogo()
    {
        using var dialog = new OpenFileDialog
        {
            Filter = "Image files|*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.tif;*.tiff;*.webp|All files|*.*",
            Title = "Select owner logo"
        };

        if (dialog.ShowDialog() != DialogResult.OK)
        {
            return;
        }

        var extension = Path.GetExtension(dialog.FileName);
        var savedPath = Path.Combine(_dataDirectory, $"owner-logo{extension}");

        foreach (var file in Directory.GetFiles(_dataDirectory, "owner-logo.*"))
        {
            File.Delete(file);
        }

        File.Copy(dialog.FileName, savedPath, true);
        _logoPath = savedPath;
        LoadLogoPreview();
        SaveState();
    }

    private void LoadLogoPreview()
    {
        _logoPreview.Image?.Dispose();
        _logoPreview.Image = null;

        if (string.IsNullOrWhiteSpace(_logoPath) || !File.Exists(_logoPath))
        {
            return;
        }

        using var stream = new FileStream(_logoPath, FileMode.Open, FileAccess.Read);
        using var source = Image.FromStream(stream);
        _logoPreview.Image = new Bitmap(source);
    }

    private void RefreshMaterialSuggestions()
    {
        var savedSuggestions = _dataStore.LoadMaterialSuggestions();

        var currentItems = _itemsGrid.Rows
            .Cast<DataGridViewRow>()
            .Where(r => !r.IsNewRow)
            .Select(r => r.Cells["Item"].Value?.ToString()?.Trim() ?? string.Empty)
            .Where(name => !string.IsNullOrWhiteSpace(name));

        _materialSuggestions = savedSuggestions
            .Concat(currentItems)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private void OnItemsGridEditingControlShowing(object? sender, DataGridViewEditingControlShowingEventArgs e)
    {
        if (_itemsGrid.CurrentCell?.OwningColumn?.Name != "Item" || e.Control is not TextBox textBox)
        {
            return;
        }

        var source = new AutoCompleteStringCollection();
        source.AddRange(_materialSuggestions.ToArray());
        textBox.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
        textBox.AutoCompleteSource = AutoCompleteSource.CustomSource;
        textBox.AutoCompleteCustomSource = source;
    }

    private void AddDefaultRow()
    {
        var index = _itemsGrid.Rows.Add();
        var row = _itemsGrid.Rows[index];
        row.Cells["No"].Value = (index + 1).ToString(CultureInfo.InvariantCulture);
        UpdateRowAmount(index);
        RecalculateTotals();
    }

    private void DeleteSelectedRow()
    {
        if (_itemsGrid.SelectedRows.Count == 0)
        {
            return;
        }

        foreach (DataGridViewRow row in _itemsGrid.SelectedRows)
        {
            if (!row.IsNewRow)
            {
                _itemsGrid.Rows.Remove(row);
            }
        }

        ResequenceRows();
        RecalculateTotals();
    }

    private void ResequenceRows()
    {
        for (var i = 0; i < _itemsGrid.Rows.Count; i++)
        {
            _itemsGrid.Rows[i].Cells["No"].Value = (i + 1).ToString(CultureInfo.InvariantCulture);
        }
    }

    private void UpdateRowAmount(int rowIndex)
    {
        if (rowIndex < 0 || rowIndex >= _itemsGrid.Rows.Count)
        {
            return;
        }

        var row = _itemsGrid.Rows[rowIndex];
        var qty = ToDouble(row.Cells["Qty"].Value);
        var rate = ToDouble(row.Cells["Rate"].Value);
        var amount = qty * rate;
        row.Cells["Amount"].Value = amount.ToString("0.00", CultureInfo.InvariantCulture);
    }

    private void RecalculateTotals()
    {
        decimal subTotal = 0;

        foreach (DataGridViewRow row in _itemsGrid.Rows)
        {
            subTotal += ToDecimal(row.Cells["Amount"].Value);
        }

        var gstAmount = Math.Round(subTotal * (_gstPercent.Value / 100M), 2);
        var grandTotal = subTotal + gstAmount;

        _txtSubTotal.Text = subTotal.ToString("0.00", CultureInfo.InvariantCulture);
        _txtGstAmount.Text = gstAmount.ToString("0.00", CultureInfo.InvariantCulture);
        _txtGrandTotal.Text = grandTotal.ToString("0.00", CultureInfo.InvariantCulture);
        _txtAmountWords.Text = ToIndianCurrencyWords((long)Math.Round(grandTotal, MidpointRounding.AwayFromZero));
    }

    private void ExportExcel()
    {
        using var dialog = new SaveFileDialog
        {
            Filter = "Excel Workbook|*.xlsx",
            FileName = $"Quotation-{_txtQuotationNo.Text.Trim()}-{DateTime.Now:yyyyMMddHHmmss}.xlsx"
        };

        if (dialog.ShowDialog() != DialogResult.OK)
        {
            return;
        }

        using var workbook = new XLWorkbook();
        var ws = workbook.Worksheets.Add("Quotation");

        ws.Column("A").Width = 5;
        ws.Column("B").Width = 16;
        ws.Column("C").Width = 16;
        ws.Column("D").Width = 12;
        ws.Column("E").Width = 12;
        ws.Column("F").Width = 9;
        ws.Column("G").Width = 9;
        ws.Column("H").Width = 12;
        ws.Column("I").Width = 14;

        ws.Range("A1:C1").Merge().Value = "QUOTATION";
        ws.Cell("A1").Style.Font.SetBold().SetFontColor(XLColor.DarkBlue).SetUnderline(XLFontUnderlineValues.Single);

        ws.Range("A2:F2").Merge().Value = "Quotation Owner Configuration";
        ws.Cell("A2").Style.Font.SetBold();
        ws.Range("A3:F3").Merge().Value = _txtCompany.Text;
        ws.Cell("A3").Style.Font.Bold = true;
        ws.Cell("A3").Style.Font.FontSize = 14;
        ws.Range("A4:F4").Merge().Value = $"GSTIN: {_txtGstin.Text}";
        ws.Range("A5:F5").Merge().Value = $"Address: {_txtCompanyAddress.Text}";
        ws.Range("A6:F6").Merge().Value = $"Phone: {_txtCompanyPhone.Text}";
        ws.Range("A7:F7").Merge().Value = $"Email: {_txtCompanyEmail.Text}";

        TryInsertCompanyLogo(ws, ResolveLogoPath());

        ws.Range("A9:C9").Merge().Value = "Quotation No:";
        ws.Range("D9:D9").Value = _txtQuotationNo.Text;
        ws.Range("E9:F9").Merge().Value = "Quotation Date:";
        ws.Range("G9:G9").Value = _quoteDate.Value.ToString("dd MMM yyyy");
        ws.Range("H9:H9").Value = "Validity:";
        ws.Range("I9:I9").Value = _validityDate.Value.ToString("dd MMM yyyy");

        ws.Range("A11:B11").Merge().Value = "Customer Details:";
        ws.Cell("A11").Style.Font.SetBold().SetUnderline(XLFontUnderlineValues.Single);
        ws.Range("A12:D12").Merge().Value = $"Name: {_txtCustomer.Text}";
        ws.Range("A13:D13").Merge().Value = $"Address: {_txtCustomerAddress.Text}";
        ws.Range("A14:D14").Merge().Value = $"Phone: {_txtPhone.Text}";
        ws.Range("A15:D15").Merge().Value = $"Place of Delivery: {_txtSupplyPlace.Text}";

        ws.Range("A17:A18").Merge().Value = "NO";
        ws.Range("B17:C18").Merge().Value = "Material Name";
        ws.Range("D17:E17").Merge().Value = "Actual Size (MM)";
        ws.Cell("D18").Value = "W";
        ws.Cell("E18").Value = "H";
        ws.Range("F17:F18").Merge().Value = "Qty";
        ws.Range("G17:G18").Merge().Value = "Soft";
        ws.Range("H17:H18").Merge().Value = "Rate";
        ws.Range("I17:I18").Merge().Value = "Amount";

        var tableHeader = ws.Range("A17:I18");
        tableHeader.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        tableHeader.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        tableHeader.Style.Font.Bold = true;

        var itemRows = GetExportItemRows();
        const int firstItemRow = 19;
        var templateItemCount = Math.Max(1, itemRows.Count);

        for (var i = 0; i < templateItemCount; i++)
        {
            var excelRow = firstItemRow + i;
            ws.Cell(excelRow, 1).Value = i + 1;
            ws.Cell(excelRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            if (i < itemRows.Count)
            {
                var row = itemRows[i];
                ws.Range(excelRow, 2, excelRow, 3).Merge().Value = row[1];
                ws.Cell(excelRow, 4).Value = row[2];
                ws.Cell(excelRow, 5).Value = row[3];
                ws.Cell(excelRow, 6).Value = row[4];
                ws.Cell(excelRow, 7).Value = row[5];
                ws.Cell(excelRow, 8).Value = row[6];
                ws.Cell(excelRow, 9).Value = row[7];
            }
            else
            {
                ws.Range(excelRow, 2, excelRow, 3).Merge();
            }
        }

        var summaryStartRow = firstItemRow + templateItemCount + 1;
        ws.Range(summaryStartRow, 5, summaryStartRow + 1, 6).Merge().Value = $"GST {_gstPercent.Value:0.##}%";
        ws.Range(summaryStartRow, 7, summaryStartRow, 8).Merge().Value = "Sub Total";
        ws.Cell(summaryStartRow, 9).Value = _txtSubTotal.Text;
        ws.Range(summaryStartRow + 1, 7, summaryStartRow + 1, 8).Merge().Value = "GST Amount";
        ws.Cell(summaryStartRow + 1, 9).Value = _txtGstAmount.Text;

        ws.Range(summaryStartRow + 3, 6, summaryStartRow + 4, 7).Merge().Value = "TOTAL AMOUNT";
        ws.Range(summaryStartRow + 3, 8, summaryStartRow + 4, 9).Merge().Value = _txtGrandTotal.Text;
        ws.Range(summaryStartRow + 3, 8, summaryStartRow + 4, 9).Style.Font.Bold = true;
        ws.Range(summaryStartRow + 3, 8, summaryStartRow + 4, 9).Style.Font.FontSize = 14;

        ws.Range(summaryStartRow + 6, 2, summaryStartRow + 6, 5).Merge().Value = "Total Amount (in words):";
        ws.Range(summaryStartRow + 6, 6, summaryStartRow + 6, 9).Merge().Value = _txtAmountWords.Text;
        ws.Cell(summaryStartRow + 6, 2).Style.Font.SetBold();

        var bankRow = summaryStartRow + 9;
        ws.Range(bankRow, 1, bankRow, 3).Merge().Value = "Bank Details:";
        ws.Cell(bankRow, 1).Style.Font.SetBold();
        ws.Range(bankRow + 1, 1, bankRow + 1, 4).Merge().Value = "Bank: State Bank of India";
        ws.Range(bankRow + 2, 1, bankRow + 2, 4).Merge().Value = "Branch: Siruthozhil, Ambattur";
        ws.Range(bankRow + 3, 1, bankRow + 3, 4).Merge().Value = "Account No: 44068068544";
        ws.Range(bankRow + 4, 1, bankRow + 4, 4).Merge().Value = "IFSC Code: SBIN0004032";

        ws.Range(bankRow + 1, 7, bankRow + 1, 9).Merge().Value = $"For {_txtCompany.Text}";
        ws.Range(bankRow + 1, 7, bankRow + 1, 9).Style.Font.SetBold();
        ws.Range(bankRow + 5, 7, bankRow + 5, 9).Merge().Value = "Authorized Signatory";
        ws.Range(bankRow + 5, 7, bankRow + 5, 9).Style.Font.SetBold();

        var termsRow = bankRow + 7;
        ws.Range(termsRow, 1, termsRow, 3).Merge().Value = "Terms & Conditions:";
        ws.Cell(termsRow, 1).Style.Font.SetBold();
        ws.Range(termsRow + 1, 1, termsRow + 1, 6).Merge().Value = "1. Payment 100% Advance.";
        ws.Range(termsRow + 2, 1, termsRow + 2, 6).Merge().Value = "2. Any extra items supplied will be charged.";
        ws.Range(termsRow + 3, 1, termsRow + 3, 6).Merge().Value = "3. No cancellation once work has commenced.";

        ws.PageSetup.Footer.Center.AddText("Page &P/&N :This is a computer generated document and requires no signature.");

        ws.Range(17, 1, firstItemRow + templateItemCount - 1, 9).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        ws.Range(17, 1, firstItemRow + templateItemCount - 1, 9).Style.Border.InsideBorder = XLBorderStyleValues.Thin;
        ws.Range(summaryStartRow, 5, summaryStartRow + 1, 9).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        ws.Range(summaryStartRow, 5, summaryStartRow + 1, 9).Style.Border.InsideBorder = XLBorderStyleValues.Thin;
        ws.Range(summaryStartRow + 3, 6, summaryStartRow + 4, 9).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

        workbook.SaveAs(dialog.FileName);
        MessageBox.Show("Excel exported successfully.", "Export", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    private static void TryInsertCompanyLogo(IXLWorksheet worksheet, string? logoPath)
    {
        if (string.IsNullOrWhiteSpace(logoPath) || !File.Exists(logoPath))
        {
            return;
        }

        var picture = worksheet.AddPicture(logoPath).MoveTo(worksheet.Cell("H2"));
        picture.WithSize(120, 90);
    }

    private void ExportPdf()
    {
        using var dialog = new SaveFileDialog
        {
            Filter = "PDF Document|*.pdf",
            FileName = $"Quotation-{_txtQuotationNo.Text.Trim()}-{DateTime.Now:yyyyMMddHHmmss}.pdf"
        };

        if (dialog.ShowDialog() != DialogResult.OK)
        {
            return;
        }

        var logoBytes = LoadLogoBytes(ResolveLogoPath());

        var itemRows = GetExportItemRows();

        Document.Create(container =>
        {
            container.Page(page =>
            {
                page.Size(PageSizes.A4);
                page.Margin(20);
                page.DefaultTextStyle(x => x.FontSize(10));
                page.Footer().AlignCenter().Text(text =>
                {
                    text.Span("Page ").SemiBold();
                    text.CurrentPageNumber().SemiBold();
                    text.Span("/").SemiBold();
                    text.TotalPages().SemiBold();
                    text.Span(" :This is a computer generated document and requires no signature.").SemiBold();
                });

                page.Content().Column(col =>
                {
                    col.Spacing(8);
                    col.Item().Row(row =>
                    {
                        row.RelativeItem().Column(left =>
                        {
                            left.Item().Text("QUOTATION").SemiBold().FontSize(18).FontColor(Colors.Blue.Darken3);
                            left.Item().Text("Quotation Owner Configuration").SemiBold();
                            left.Item().Text(_txtCompany.Text).SemiBold().FontSize(12);
                            left.Item().Text($"GSTIN: {_txtGstin.Text}");
                            left.Item().Text($"Address: {_txtCompanyAddress.Text}");
                            left.Item().Text($"Phone: {_txtCompanyPhone.Text}");
                            left.Item().Text($"Email: {_txtCompanyEmail.Text}");
                        });

                        row.ConstantItem(120).Height(90).Border(1).BorderColor(Colors.Grey.Lighten1).AlignMiddle().AlignCenter().Element(c =>
                        {
                            if (logoBytes is not null)
                            {
                                c.Image(logoBytes).FitArea();
                            }
                            else
                            {
                                c.Text("Logo").FontColor(Colors.Grey.Darken1);
                            }
                        });
                    });

                    col.Item().Row(row =>
                    {
                        row.RelativeItem().Text($"Quotation No: {_txtQuotationNo.Text}");
                        row.RelativeItem().AlignCenter().Text($"Quotation Date: {_quoteDate.Value:dd MMM yyyy}");
                        row.RelativeItem().AlignRight().Text($"Validity: {_validityDate.Value:dd MMM yyyy}");
                    });

                    col.Item().Text("Customer Details").SemiBold().Underline();
                    col.Item().Text($"Name: {_txtCustomer.Text}");
                    col.Item().Text($"Address: {_txtCustomerAddress.Text}");
                    col.Item().Text($"Phone No: {_txtPhone.Text}");
                    col.Item().Text($"Place of Delivery: {_txtSupplyPlace.Text}");

                    col.Item().PaddingTop(6).Table(table =>
                    {
                        table.ColumnsDefinition(columns =>
                        {
                            columns.ConstantColumn(35);
                            columns.RelativeColumn(3);
                            columns.RelativeColumn();
                            columns.RelativeColumn();
                            columns.RelativeColumn();
                            columns.RelativeColumn();
                            columns.RelativeColumn();
                            columns.RelativeColumn();
                        });

                        static void Header(IContainer container, string text) =>
                            container.Background(Colors.Grey.Lighten2).Padding(4).Text(text).SemiBold();

                        Header(table.Cell(), "NO");
                        Header(table.Cell(), "Material Name");
                        Header(table.Cell(), "W");
                        Header(table.Cell(), "H");
                        Header(table.Cell(), "Qty");
                        Header(table.Cell(), "Soft");
                        Header(table.Cell(), "Rate");
                        Header(table.Cell(), "Amount");

                        if (itemRows.Count == 0)
                        {
                            table.Cell().ColumnSpan(8).BorderBottom(1).BorderColor(Colors.Grey.Lighten2).Padding(4).Text("No materials added").Italic();
                        }

                        foreach (var row in itemRows)
                        {
                            foreach (var value in row)
                            {
                                table.Cell().BorderBottom(1).BorderColor(Colors.Grey.Lighten2).Padding(4).Text(value);
                            }
                        }
                    });

                    col.Item().PaddingTop(10).Row(row =>
                    {
                        row.RelativeItem();
                        row.ConstantItem(270).Table(table =>
                        {
                            table.ColumnsDefinition(c =>
                            {
                                c.RelativeColumn(2);
                                c.RelativeColumn(1);
                            });

                            void Summary(string label, string value, bool highlight = false)
                            {
                                table.Cell().Border(1).Padding(4).Text(label).SemiBold();
                                table.Cell().Border(1).Padding(4).AlignRight().Text(value).Style(highlight ? TextStyle.Default.SemiBold().FontSize(14) : TextStyle.Default);
                            }

                            Summary("Sub Total", _txtSubTotal.Text);
                            Summary($"GST {_gstPercent.Value:0.##}%", _txtGstAmount.Text);
                            Summary("TOTAL AMOUNT", _txtGrandTotal.Text, true);
                        });
                    });

                    col.Item().Text($"Total Amount (in words): {_txtAmountWords.Text}").SemiBold();

                    col.Item().PaddingTop(8).Row(row =>
                    {
                        row.RelativeItem().Column(bank =>
                        {
                            bank.Item().Text("Bank Details").SemiBold();
                            bank.Item().Text("Bank: State Bank of India");
                            bank.Item().Text("Branch: Siruthozhil, Ambattur");
                            bank.Item().Text("Account No: 44068068544");
                            bank.Item().Text("IFSC Code: SBIN0004032");
                        });

                        row.RelativeItem().AlignRight().Column(sig =>
                        {
                            sig.Item().Text($"For {_txtCompany.Text}").SemiBold();
                            sig.Item().PaddingTop(35).Text("Authorized Signatory").SemiBold();
                        });
                    });

                    col.Item().PaddingTop(6).Text("Terms & Conditions:").SemiBold();
                    col.Item().Text("1. Payment 100% Advance.");
                    col.Item().Text("2. Any extra items supplied will be charged.");
                    col.Item().Text("3. No cancellation once work has commenced.");

                });
            });
        }).GeneratePdf(dialog.FileName);

        MessageBox.Show("PDF exported successfully.", "Export", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    private string? ResolveLogoPath()
    {
        if (!string.IsNullOrWhiteSpace(_logoPath) && File.Exists(_logoPath))
        {
            return _logoPath;
        }

        var logoCandidates = new[]
        {
            Path.Combine(AppContext.BaseDirectory, "logo.png"),
            Path.Combine(AppContext.BaseDirectory, "rr-logo.png"),
            Path.Combine(AppContext.BaseDirectory, "assets", "logo.png")
        };

        return logoCandidates.FirstOrDefault(File.Exists);
    }

    private static byte[]? LoadLogoBytes(string? logoPath)
    {
        return string.IsNullOrWhiteSpace(logoPath) || !File.Exists(logoPath)
            ? null
            : File.ReadAllBytes(logoPath);
    }

    private List<string[]> GetExportItemRows()
    {
        var rows = _itemsGrid.Rows
            .Cast<DataGridViewRow>()
            .Where(r => !r.IsNewRow)
            .Select(r => new
            {
                MaterialName = r.Cells["Item"].Value?.ToString()?.Trim() ?? string.Empty,
                W = r.Cells["W"].Value?.ToString() ?? string.Empty,
                H = r.Cells["H"].Value?.ToString() ?? string.Empty,
                Qty = r.Cells["Qty"].Value?.ToString() ?? string.Empty,
                Soft = r.Cells["Soft"].Value?.ToString() ?? string.Empty,
                Rate = r.Cells["Rate"].Value?.ToString() ?? string.Empty,
                Amount = r.Cells["Amount"].Value?.ToString() ?? string.Empty
            })
            .Where(r => !string.IsNullOrWhiteSpace(r.MaterialName))
            .ToList();

        var result = new List<string[]>(rows.Count);
        for (var i = 0; i < rows.Count; i++)
        {
            var row = rows[i];
            result.Add([
                (i + 1).ToString(CultureInfo.InvariantCulture),
                row.MaterialName,
                row.W,
                row.H,
                row.Qty,
                row.Soft,
                row.Rate,
                row.Amount
            ]);
        }

        return result;
    }

    private static decimal ToDecimal(object? value)
    {
        if (value is null)
        {
            return 0;
        }

        return decimal.TryParse(value.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out var parsed)
            ? parsed
            : 0;
    }

    private static double ToDouble(object? value)
    {
        if (value is null)
        {
            return 0;
        }

        return double.TryParse(value.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out var parsed)
            ? parsed
            : 0;
    }

    private static string ToIndianCurrencyWords(long number)
    {
        if (number == 0)
        {
            return "INR Zero Only";
        }

        var parts = new List<string>();

        void AddPart(long value, string suffix)
        {
            if (value > 0)
            {
                parts.Add($"{NumberToWords(value)} {suffix}".Trim());
            }
        }

        AddPart(number / 10000000, "Crore");
        number %= 10000000;
        AddPart(number / 100000, "Lakh");
        number %= 100000;
        AddPart(number / 1000, "Thousand");
        number %= 1000;
        AddPart(number / 100, "Hundred");
        number %= 100;

        if (number > 0)
        {
            if (parts.Count > 0)
            {
                parts.Add("and");
            }
            parts.Add(NumberToWords(number));
        }

        return $"INR {string.Join(" ", parts)} Rupees Only";
    }

    private static string NumberToWords(long number)
    {
        string[] units =
        [
            "Zero", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten",
            "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen"
        ];

        string[] tens = ["Zero", "Ten", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"];

        if (number < 20)
        {
            return units[number];
        }

        if (number < 100)
        {
            return tens[number / 10] + (number % 10 > 0 ? " " + units[number % 10] : string.Empty);
        }

        return string.Empty;
    }
}
