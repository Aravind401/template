namespace QuotationTemplateApp;

internal sealed class OwnerDetailsForm : Form
{
    private readonly TextBox _txtCompany;
    private readonly TextBox _txtGstin;
    private readonly TextBox _txtAddress;
    private readonly TextBox _txtPhone;
    private readonly TextBox _txtEmail;
    private readonly TextBox _txtLogoPath;

    public OwnerData OwnerDetails { get; private set; }

    public OwnerDetailsForm(OwnerData currentOwner)
    {
        OwnerDetails = currentOwner;

        Text = "Owner Details";
        Width = 720;
        Height = 430;
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MinimizeBox = false;
        MaximizeBox = false;

        _txtCompany = new TextBox { Text = currentOwner.Company, Dock = DockStyle.Fill };
        _txtGstin = new TextBox { Text = currentOwner.Gstin, Dock = DockStyle.Fill };
        _txtAddress = new TextBox { Text = currentOwner.Address, Multiline = true, ScrollBars = ScrollBars.Vertical, Dock = DockStyle.Fill };
        _txtPhone = new TextBox { Text = currentOwner.Phone, Dock = DockStyle.Fill };
        _txtEmail = new TextBox { Text = currentOwner.Email, Dock = DockStyle.Fill };
        _txtLogoPath = new TextBox { Text = currentOwner.LogoPath ?? string.Empty, Dock = DockStyle.Fill, ReadOnly = true };

        BuildLayout();
    }

    private void BuildLayout()
    {
        var root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 2,
            Padding = new Padding(12)
        };
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 56));

        var form = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 2,
            RowCount = 7
        };
        form.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 140));
        form.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        form.RowStyles.Add(new RowStyle(SizeType.Absolute, 34));
        form.RowStyles.Add(new RowStyle(SizeType.Absolute, 34));
        form.RowStyles.Add(new RowStyle(SizeType.Absolute, 84));
        form.RowStyles.Add(new RowStyle(SizeType.Absolute, 34));
        form.RowStyles.Add(new RowStyle(SizeType.Absolute, 34));
        form.RowStyles.Add(new RowStyle(SizeType.Absolute, 34));
        form.RowStyles.Add(new RowStyle(SizeType.Absolute, 34));

        AddField(form, "Company", _txtCompany, 0);
        AddField(form, "GSTIN", _txtGstin, 1);
        AddField(form, "Address", _txtAddress, 2);
        AddField(form, "Phone", _txtPhone, 3);
        AddField(form, "Email", _txtEmail, 4);

        var logoHost = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1 };
        logoHost.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        logoHost.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 120));
        logoHost.Controls.Add(_txtLogoPath, 0, 0);

        var browse = new Button { Text = "Browse...", Dock = DockStyle.Fill };
        browse.Click += (_, _) => BrowseLogoPath();
        logoHost.Controls.Add(browse, 1, 0);
        AddField(form, "Logo", logoHost, 5);

        root.Controls.Add(form, 0, 0);

        var actions = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            FlowDirection = FlowDirection.RightToLeft,
            WrapContents = false
        };

        var save = new Button { Text = "Save", Width = 110, Height = 36 };
        save.Click += (_, _) => SaveAndClose();

        var cancel = new Button { Text = "Cancel", Width = 110, Height = 36 };
        cancel.Click += (_, _) => DialogResult = DialogResult.Cancel;

        actions.Controls.Add(save);
        actions.Controls.Add(cancel);
        root.Controls.Add(actions, 0, 1);

        Controls.Add(root);
    }

    private static void AddField(TableLayoutPanel panel, string label, Control input, int row)
    {
        var caption = new Label
        {
            Text = label,
            Font = new Font("Segoe UI", 9, FontStyle.Bold),
            Dock = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleLeft
        };

        panel.Controls.Add(caption, 0, row);
        panel.Controls.Add(input, 1, row);
    }

    private void BrowseLogoPath()
    {
        using var dialog = new OpenFileDialog
        {
            Filter = "Image files|*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.tif;*.tiff;*.webp|All files|*.*",
            Title = "Select owner logo"
        };

        if (dialog.ShowDialog(this) == DialogResult.OK)
        {
            _txtLogoPath.Text = dialog.FileName;
        }
    }

    private void SaveAndClose()
    {
        OwnerDetails = OwnerDetails with
        {
            Company = _txtCompany.Text.Trim(),
            Gstin = _txtGstin.Text.Trim(),
            Address = _txtAddress.Text.Trim(),
            Phone = _txtPhone.Text.Trim(),
            Email = _txtEmail.Text.Trim(),
            LogoPath = string.IsNullOrWhiteSpace(_txtLogoPath.Text) ? null : _txtLogoPath.Text.Trim()
        };

        DialogResult = DialogResult.OK;
    }
}
