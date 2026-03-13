using ClosedXML.Excel;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;
using System.ComponentModel;
using System.Globalization;

namespace QuotationTemplateApp;

public class MainForm : Form
{
    private readonly DataGridView _itemsGrid = new();
    private readonly TextBox _txtCompany = new() { Text = "R.R Engineering" };
    private readonly TextBox _txtCustomer = new();
    private readonly TextBox _txtPhone = new();
    private readonly TextBox _txtSupplyPlace = new() { Text = "Tamil Nadu" };
    private readonly TextBox _txtQuotationNo = new() { Text = "EST-15" };
    private readonly DateTimePicker _quoteDate = new() { Value = DateTime.Today, Format = DateTimePickerFormat.Short };
    private readonly DateTimePicker _validityDate = new() { Value = DateTime.Today, Format = DateTimePickerFormat.Short };

    private readonly NumericUpDown _gstPercent = new() { DecimalPlaces = 2, Minimum = 0, Maximum = 100, Value = 18 };
    private readonly TextBox _txtSubTotal = new() { ReadOnly = true };
    private readonly TextBox _txtGstAmount = new() { ReadOnly = true };
    private readonly TextBox _txtGrandTotal = new() { ReadOnly = true };
    private readonly TextBox _txtAmountWords = new() { ReadOnly = true, Multiline = true, Height = 55 };

    private readonly Button _btnAddRow = new() { Text = "Add Row" };
    private readonly Button _btnDeleteRow = new() { Text = "Delete Selected Row" };
    private readonly Button _btnRecalculate = new() { Text = "Recalculate" };
    private readonly Button _btnExportExcel = new() { Text = "Export to Excel" };
    private readonly Button _btnExportPdf = new() { Text = "Export to PDF" };

    public MainForm()
    {
        Text = "Quotation Template (WinForms)";
        Width = 1200;
        Height = 780;
        StartPosition = FormStartPosition.CenterScreen;

        BuildLayout();
        ConfigureGrid();
        WireEvents();

        for (var i = 0; i < 18; i++)
        {
            AddDefaultRow();
        }

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

        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 140));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 120));
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 45));

        var header = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 6, RowCount = 3 };
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 16));
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 17));
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 16));
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 17));
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 16));
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 18));

        AddField(header, "Company", _txtCompany, 0, 0);
        AddField(header, "Customer", _txtCustomer, 2, 0);
        AddField(header, "Phone", _txtPhone, 4, 0);
        AddField(header, "Place of Supply", _txtSupplyPlace, 0, 1);
        AddField(header, "Quotation No", _txtQuotationNo, 2, 1);
        AddField(header, "Quote Date", _quoteDate, 4, 1);
        AddField(header, "Validity Date", _validityDate, 0, 2);

        var gridHost = new Panel { Dock = DockStyle.Fill };
        _itemsGrid.Dock = DockStyle.Fill;
        gridHost.Controls.Add(_itemsGrid);

        var totals = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 6,
            RowCount = 4,
            Padding = new Padding(0, 10, 0, 0)
        };

        for (var i = 0; i < 6; i++)
        {
            totals.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 16.66F));
        }

        AddField(totals, "GST %", _gstPercent, 0, 0);
        AddField(totals, "Sub Total", _txtSubTotal, 2, 0);
        AddField(totals, "GST Amount", _txtGstAmount, 4, 0);
        AddField(totals, "Grand Total", _txtGrandTotal, 2, 1);
        AddField(totals, "Total Amount (in words)", _txtAmountWords, 0, 2, 6);

        var actions = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            FlowDirection = FlowDirection.LeftToRight
        };

        actions.Controls.AddRange(new Control[] { _btnAddRow, _btnDeleteRow, _btnRecalculate, _btnExportExcel, _btnExportPdf });

        root.Controls.Add(header, 0, 0);
        root.Controls.Add(gridHost, 0, 1);
        root.Controls.Add(totals, 0, 2);
        root.Controls.Add(actions, 0, 3);

        Controls.Add(root);
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

        editor.Dock = DockStyle.Fill;

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
        _itemsGrid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Item", HeaderText = "Item", Width = 320 });
        _itemsGrid.Columns.Add(new DataGridViewTextBoxColumn { Name = "W", HeaderText = "W", Width = 80 });
        _itemsGrid.Columns.Add(new DataGridViewTextBoxColumn { Name = "H", HeaderText = "H", Width = 80 });
        _itemsGrid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Qty", HeaderText = "Qty", Width = 80 });
        _itemsGrid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Soft", HeaderText = "Soft", Width = 80 });
        _itemsGrid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Rate", HeaderText = "Rate", Width = 120 });
        _itemsGrid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Amount", HeaderText = "Amount", Width = 120, ReadOnly = true });
    }

    private void WireEvents()
    {
        _btnAddRow.Click += (_, _) => AddDefaultRow();
        _btnDeleteRow.Click += (_, _) => DeleteSelectedRow();
        _btnRecalculate.Click += (_, _) => RecalculateTotals();
        _btnExportExcel.Click += (_, _) => ExportExcel();
        _btnExportPdf.Click += (_, _) => ExportPdf();
        _gstPercent.ValueChanged += (_, _) => RecalculateTotals();

        _itemsGrid.CellEndEdit += (_, e) =>
        {
            if (e.RowIndex >= 0)
            {
                UpdateRowAmount(e.RowIndex);
                RecalculateTotals();
            }
        };
    }

    private void AddDefaultRow()
    {
        var index = _itemsGrid.Rows.Add();
        var row = _itemsGrid.Rows[index];
        row.Cells["No"].Value = (index + 1).ToString(CultureInfo.InvariantCulture);
        row.Cells["Qty"].Value = "1";
        row.Cells["Rate"].Value = "100";
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
        var qty = ToDecimal(row.Cells["Qty"].Value);
        var rate = ToDecimal(row.Cells["Rate"].Value);
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

        ws.Cell("A1").Value = "QUOTATION";
        ws.Cell("A2").Value = _txtCompany.Text;
        ws.Cell("A3").Value = $"Customer: {_txtCustomer.Text}";
        ws.Cell("D3").Value = $"Phone: {_txtPhone.Text}";
        ws.Cell("A4").Value = $"Place of Supply: {_txtSupplyPlace.Text}";
        ws.Cell("D4").Value = $"Quotation No: {_txtQuotationNo.Text}";
        ws.Cell("F4").Value = $"Quotation Date: {_quoteDate.Value:dd MMM yyyy}";
        ws.Cell("H4").Value = $"Validity: {_validityDate.Value:dd MMM yyyy}";

        var headers = new[] { "NO", "Item", "W", "H", "Qty", "Soft", "Rate", "Amount" };
        for (var i = 0; i < headers.Length; i++)
        {
            ws.Cell(6, i + 1).Value = headers[i];
            ws.Cell(6, i + 1).Style.Font.Bold = true;
            ws.Cell(6, i + 1).Style.Fill.BackgroundColor = XLColor.LightGray;
        }

        var rowPointer = 7;
        foreach (DataGridViewRow row in _itemsGrid.Rows)
        {
            ws.Cell(rowPointer, 1).Value = row.Cells["No"].Value?.ToString();
            ws.Cell(rowPointer, 2).Value = row.Cells["Item"].Value?.ToString();
            ws.Cell(rowPointer, 3).Value = row.Cells["W"].Value?.ToString();
            ws.Cell(rowPointer, 4).Value = row.Cells["H"].Value?.ToString();
            ws.Cell(rowPointer, 5).Value = row.Cells["Qty"].Value?.ToString();
            ws.Cell(rowPointer, 6).Value = row.Cells["Soft"].Value?.ToString();
            ws.Cell(rowPointer, 7).Value = row.Cells["Rate"].Value?.ToString();
            ws.Cell(rowPointer, 8).Value = row.Cells["Amount"].Value?.ToString();
            rowPointer++;
        }

        rowPointer += 1;
        ws.Cell(rowPointer, 6).Value = "Sub Total";
        ws.Cell(rowPointer, 8).Value = _txtSubTotal.Text;
        rowPointer++;
        ws.Cell(rowPointer, 6).Value = $"GST {_gstPercent.Value:0.##}%";
        ws.Cell(rowPointer, 8).Value = _txtGstAmount.Text;
        rowPointer++;
        ws.Cell(rowPointer, 6).Value = "Grand Total";
        ws.Cell(rowPointer, 8).Value = _txtGrandTotal.Text;
        rowPointer++;
        ws.Cell(rowPointer, 1).Value = "Amount in words:";
        ws.Cell(rowPointer, 2).Value = _txtAmountWords.Text;
        ws.Range(rowPointer, 2, rowPointer, 8).Merge();

        ws.Columns().AdjustToContents();
        ws.RangeUsed().Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        ws.RangeUsed().Style.Border.InsideBorder = XLBorderStyleValues.Thin;

        workbook.SaveAs(dialog.FileName);
        MessageBox.Show("Excel exported successfully.", "Export", MessageBoxButtons.OK, MessageBoxIcon.Information);
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

        var itemRows = _itemsGrid.Rows
            .Cast<DataGridViewRow>()
            .Where(r => !r.IsNewRow)
            .Select(r => new[]
            {
                r.Cells["No"].Value?.ToString() ?? string.Empty,
                r.Cells["Item"].Value?.ToString() ?? string.Empty,
                r.Cells["W"].Value?.ToString() ?? string.Empty,
                r.Cells["H"].Value?.ToString() ?? string.Empty,
                r.Cells["Qty"].Value?.ToString() ?? string.Empty,
                r.Cells["Soft"].Value?.ToString() ?? string.Empty,
                r.Cells["Rate"].Value?.ToString() ?? string.Empty,
                r.Cells["Amount"].Value?.ToString() ?? string.Empty
            })
            .ToList();

        Document.Create(container =>
        {
            container.Page(page =>
            {
                page.Size(PageSizes.A4);
                page.Margin(20);
                page.DefaultTextStyle(x => x.FontSize(10));

                page.Content().Column(col =>
                {
                    col.Spacing(8);
                    col.Item().Text("QUOTATION").SemiBold().FontSize(16);
                    col.Item().Text(_txtCompany.Text).SemiBold().FontSize(12);
                    col.Item().Text($"Customer: {_txtCustomer.Text}");
                    col.Item().Text($"Phone: {_txtPhone.Text}");
                    col.Item().Text($"Place of Supply: {_txtSupplyPlace.Text}");
                    col.Item().Text($"Quotation No: {_txtQuotationNo.Text}");
                    col.Item().Text($"Quotation Date: {_quoteDate.Value:dd MMM yyyy}");
                    col.Item().Text($"Validity: {_validityDate.Value:dd MMM yyyy}");

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
                        Header(table.Cell(), "Item");
                        Header(table.Cell(), "W");
                        Header(table.Cell(), "H");
                        Header(table.Cell(), "Qty");
                        Header(table.Cell(), "Soft");
                        Header(table.Cell(), "Rate");
                        Header(table.Cell(), "Amount");

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
                        row.ConstantItem(250).Column(summary =>
                        {
                            summary.Item().Text($"Sub Total: {_txtSubTotal.Text}");
                            summary.Item().Text($"GST {_gstPercent.Value:0.##}%: {_txtGstAmount.Text}");
                            summary.Item().Text($"Grand Total: {_txtGrandTotal.Text}").SemiBold();
                        });
                    });

                    col.Item().PaddingTop(8).Text($"Amount in words: {_txtAmountWords.Text}");
                });
            });
        }).GeneratePdf(dialog.FileName);

        MessageBox.Show("PDF exported successfully.", "Export", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
