namespace QuotationTemplateApp;

internal sealed class CustomerDetailsForm : Form
{
    private readonly TextBox _txtName;
    private readonly TextBox _txtAddress;
    private readonly TextBox _txtPhone;
    private readonly TextBox _txtSupplyPlace;

    public QuotationData CustomerDetails { get; private set; }

    public CustomerDetailsForm(QuotationData currentDetails)
    {
        CustomerDetails = currentDetails;

        Text = "Customer Details";
        Width = 720;
        Height = 380;
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MinimizeBox = false;
        MaximizeBox = false;

        _txtName = new TextBox { Text = currentDetails.CustomerName, Dock = DockStyle.Fill };
        _txtAddress = new TextBox { Text = currentDetails.CustomerAddress, Multiline = true, ScrollBars = ScrollBars.Vertical, Dock = DockStyle.Fill };
        _txtPhone = new TextBox { Text = currentDetails.CustomerPhone, Dock = DockStyle.Fill };
        _txtSupplyPlace = new TextBox { Text = currentDetails.SupplyPlace, Dock = DockStyle.Fill };

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
            RowCount = 4
        };

        form.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 140));
        form.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        form.RowStyles.Add(new RowStyle(SizeType.Absolute, 36));
        form.RowStyles.Add(new RowStyle(SizeType.Absolute, 96));
        form.RowStyles.Add(new RowStyle(SizeType.Absolute, 36));
        form.RowStyles.Add(new RowStyle(SizeType.Absolute, 36));

        AddField(form, "Name", _txtName, 0);
        AddField(form, "Address", _txtAddress, 1);
        AddField(form, "Phone", _txtPhone, 2);
        AddField(form, "Place of Delivery", _txtSupplyPlace, 3);

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

    private void SaveAndClose()
    {
        CustomerDetails = CustomerDetails with
        {
            CustomerName = _txtName.Text.Trim(),
            CustomerAddress = _txtAddress.Text.Trim(),
            CustomerPhone = _txtPhone.Text.Trim(),
            SupplyPlace = _txtSupplyPlace.Text.Trim()
        };

        DialogResult = DialogResult.OK;
    }
}
