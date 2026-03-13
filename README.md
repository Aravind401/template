# Quotation Template Windows App

This repository contains a **C# WinForms (.NET 8)** desktop application for creating quotation sheets.

## Features
- Manual row/column entry for quotation items.
- Add row and delete selected row support.
- Auto amount calculation (`Qty x Rate`) per row.
- GST percentage input and automatic subtotal/GST/grand total calculation.
- Total amount conversion to words (Indian currency format).
- Export full quotation data to an Excel (`.xlsx`) file.
- Export quotation data to a PDF (`.pdf`) file.

## Build & Run (Windows)
1. Install .NET 8 SDK on Windows.
2. Open `QuotationApp.sln` in Visual Studio 2022+.
3. Restore NuGet packages.
4. Run the `QuotationTemplateApp` project.

## Excel Export
Use the **Export to Excel** button and choose a save location. The generated workbook includes:
- Header details (company, customer, dates, quotation number)
- Full item table
- Subtotal, GST, grand total
- Amount in words


## PDF Export
Use the **Export to PDF** button and choose a save location. The generated PDF includes:
- Header details (company, customer, dates, quotation number)
- Full item table
- Subtotal, GST, grand total
- Amount in words
