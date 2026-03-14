using Microsoft.Data.Sqlite;
using System.Globalization;

namespace QuotationTemplateApp;

internal sealed class AppDataStore
{
    private readonly string _connectionString;

    public AppDataStore(string dataDirectory)
    {
        Directory.CreateDirectory(dataDirectory);
        var databasePath = Path.Combine(dataDirectory, "quotation-data.db");
        _connectionString = $"Data Source={databasePath}";
        EnsureSchema();
    }

    public AppState LoadState()
    {
        using var connection = new SqliteConnection(_connectionString);
        connection.Open();

        var owner = ReadOwner(connection);
        var quote = ReadQuotation(connection);
        var items = ReadItems(connection);

        return new AppState(owner, quote, items);
    }

    public void SaveState(AppState state)
    {
        using var connection = new SqliteConnection(_connectionString);
        connection.Open();
        using var transaction = connection.BeginTransaction();

        UpsertOwner(connection, transaction, state.Owner);
        UpsertQuotation(connection, transaction, state.Quotation);
        ReplaceItems(connection, transaction, state.Items);
        SaveMaterialSuggestions(connection, transaction, state.Items);

        transaction.Commit();
    }

    public List<string> LoadMaterialSuggestions()
    {
        using var connection = new SqliteConnection(_connectionString);
        connection.Open();

        using var command = connection.CreateCommand();
        command.CommandText = "SELECT Name FROM MaterialSuggestions ORDER BY Name COLLATE NOCASE";

        var names = new List<string>();
        using var reader = command.ExecuteReader();
        while (reader.Read())
        {
            names.Add(reader.GetString(0));
        }

        return names;
    }

    private void EnsureSchema()
    {
        using var connection = new SqliteConnection(_connectionString);
        connection.Open();

        using var command = connection.CreateCommand();
        command.CommandText = """
            CREATE TABLE IF NOT EXISTS Owner (
                Id INTEGER PRIMARY KEY CHECK (Id = 1),
                Company TEXT NOT NULL,
                Gstin TEXT NOT NULL,
                Address TEXT NOT NULL,
                Phone TEXT NOT NULL,
                Email TEXT NOT NULL,
                LogoPath TEXT
            );

            CREATE TABLE IF NOT EXISTS Quotation (
                Id INTEGER PRIMARY KEY CHECK (Id = 1),
                CustomerName TEXT NOT NULL,
                CustomerAddress TEXT NOT NULL,
                CustomerPhone TEXT NOT NULL,
                SupplyPlace TEXT NOT NULL,
                QuotationNo TEXT NOT NULL,
                QuoteDate TEXT NOT NULL,
                ValidityDate TEXT NOT NULL,
                GstPercent REAL NOT NULL
            );

            CREATE TABLE IF NOT EXISTS QuoteItem (
                Id INTEGER PRIMARY KEY AUTOINCREMENT,
                RowNo INTEGER NOT NULL,
                MaterialName TEXT,
                W TEXT,
                H TEXT,
                Qty TEXT,
                Soft TEXT,
                Rate TEXT
            );

            CREATE TABLE IF NOT EXISTS MaterialSuggestions (
                Name TEXT PRIMARY KEY
            );
        """;
        command.ExecuteNonQuery();
    }

    private static OwnerData ReadOwner(SqliteConnection connection)
    {
        using var command = connection.CreateCommand();
        command.CommandText = "SELECT Company, Gstin, Address, Phone, Email, LogoPath FROM Owner WHERE Id = 1";

        using var reader = command.ExecuteReader();
        if (!reader.Read())
        {
            return OwnerData.Default;
        }

        return new OwnerData(
            reader.GetString(0),
            reader.GetString(1),
            reader.GetString(2),
            reader.GetString(3),
            reader.GetString(4),
            reader.IsDBNull(5) ? null : reader.GetString(5));
    }

    private static QuotationData ReadQuotation(SqliteConnection connection)
    {
        using var command = connection.CreateCommand();
        command.CommandText = "SELECT CustomerName, CustomerAddress, CustomerPhone, SupplyPlace, QuotationNo, QuoteDate, ValidityDate, GstPercent FROM Quotation WHERE Id = 1";

        using var reader = command.ExecuteReader();
        if (!reader.Read())
        {
            return QuotationData.Default;
        }

        return new QuotationData(
            reader.GetString(0),
            reader.GetString(1),
            reader.GetString(2),
            reader.GetString(3),
            reader.GetString(4),
            DateTime.Parse(reader.GetString(5), CultureInfo.InvariantCulture),
            DateTime.Parse(reader.GetString(6), CultureInfo.InvariantCulture),
            Convert.ToDecimal(reader.GetDouble(7), CultureInfo.InvariantCulture));
    }

    private static List<QuoteItemData> ReadItems(SqliteConnection connection)
    {
        using var command = connection.CreateCommand();
        command.CommandText = "SELECT RowNo, MaterialName, W, H, Qty, Soft, Rate FROM QuoteItem ORDER BY RowNo";

        using var reader = command.ExecuteReader();
        var items = new List<QuoteItemData>();
        while (reader.Read())
        {
            items.Add(new QuoteItemData(
                reader.GetInt32(0),
                reader.IsDBNull(1) ? string.Empty : reader.GetString(1),
                reader.IsDBNull(2) ? string.Empty : reader.GetString(2),
                reader.IsDBNull(3) ? string.Empty : reader.GetString(3),
                reader.IsDBNull(4) ? string.Empty : reader.GetString(4),
                reader.IsDBNull(5) ? string.Empty : reader.GetString(5),
                reader.IsDBNull(6) ? string.Empty : reader.GetString(6)));
        }

        return items;
    }

    private static void UpsertOwner(SqliteConnection connection, SqliteTransaction transaction, OwnerData owner)
    {
        using var command = connection.CreateCommand();
        command.Transaction = transaction;
        command.CommandText = """
            INSERT INTO Owner (Id, Company, Gstin, Address, Phone, Email, LogoPath)
            VALUES (1, $company, $gstin, $address, $phone, $email, $logoPath)
            ON CONFLICT(Id) DO UPDATE SET
                Company = excluded.Company,
                Gstin = excluded.Gstin,
                Address = excluded.Address,
                Phone = excluded.Phone,
                Email = excluded.Email,
                LogoPath = excluded.LogoPath;
        """;
        command.Parameters.AddWithValue("$company", owner.Company);
        command.Parameters.AddWithValue("$gstin", owner.Gstin);
        command.Parameters.AddWithValue("$address", owner.Address);
        command.Parameters.AddWithValue("$phone", owner.Phone);
        command.Parameters.AddWithValue("$email", owner.Email);
        command.Parameters.AddWithValue("$logoPath", (object?)owner.LogoPath ?? DBNull.Value);
        command.ExecuteNonQuery();
    }

    private static void UpsertQuotation(SqliteConnection connection, SqliteTransaction transaction, QuotationData quotation)
    {
        using var command = connection.CreateCommand();
        command.Transaction = transaction;
        command.CommandText = """
            INSERT INTO Quotation (Id, CustomerName, CustomerAddress, CustomerPhone, SupplyPlace, QuotationNo, QuoteDate, ValidityDate, GstPercent)
            VALUES (1, $customerName, $customerAddress, $customerPhone, $supplyPlace, $quotationNo, $quoteDate, $validityDate, $gstPercent)
            ON CONFLICT(Id) DO UPDATE SET
                CustomerName = excluded.CustomerName,
                CustomerAddress = excluded.CustomerAddress,
                CustomerPhone = excluded.CustomerPhone,
                SupplyPlace = excluded.SupplyPlace,
                QuotationNo = excluded.QuotationNo,
                QuoteDate = excluded.QuoteDate,
                ValidityDate = excluded.ValidityDate,
                GstPercent = excluded.GstPercent;
        """;
        command.Parameters.AddWithValue("$customerName", quotation.CustomerName);
        command.Parameters.AddWithValue("$customerAddress", quotation.CustomerAddress);
        command.Parameters.AddWithValue("$customerPhone", quotation.CustomerPhone);
        command.Parameters.AddWithValue("$supplyPlace", quotation.SupplyPlace);
        command.Parameters.AddWithValue("$quotationNo", quotation.QuotationNo);
        command.Parameters.AddWithValue("$quoteDate", quotation.QuoteDate.ToString("O", CultureInfo.InvariantCulture));
        command.Parameters.AddWithValue("$validityDate", quotation.ValidityDate.ToString("O", CultureInfo.InvariantCulture));
        command.Parameters.AddWithValue("$gstPercent", Convert.ToDouble(quotation.GstPercent, CultureInfo.InvariantCulture));
        command.ExecuteNonQuery();
    }

    private static void ReplaceItems(SqliteConnection connection, SqliteTransaction transaction, IReadOnlyCollection<QuoteItemData> items)
    {
        using (var deleteCommand = connection.CreateCommand())
        {
            deleteCommand.Transaction = transaction;
            deleteCommand.CommandText = "DELETE FROM QuoteItem";
            deleteCommand.ExecuteNonQuery();
        }

        foreach (var item in items)
        {
            using var insertCommand = connection.CreateCommand();
            insertCommand.Transaction = transaction;
            insertCommand.CommandText = """
                INSERT INTO QuoteItem (RowNo, MaterialName, W, H, Qty, Soft, Rate)
                VALUES ($rowNo, $materialName, $w, $h, $qty, $soft, $rate)
            """;
            insertCommand.Parameters.AddWithValue("$rowNo", item.RowNo);
            insertCommand.Parameters.AddWithValue("$materialName", item.MaterialName);
            insertCommand.Parameters.AddWithValue("$w", item.W);
            insertCommand.Parameters.AddWithValue("$h", item.H);
            insertCommand.Parameters.AddWithValue("$qty", item.Qty);
            insertCommand.Parameters.AddWithValue("$soft", item.Soft);
            insertCommand.Parameters.AddWithValue("$rate", item.Rate);
            insertCommand.ExecuteNonQuery();
        }
    }

    private static void SaveMaterialSuggestions(SqliteConnection connection, SqliteTransaction transaction, IReadOnlyCollection<QuoteItemData> items)
    {
        var names = items
            .Select(i => i.MaterialName.Trim())
            .Where(name => !string.IsNullOrWhiteSpace(name))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
            .ToList();

        using (var deleteCommand = connection.CreateCommand())
        {
            deleteCommand.Transaction = transaction;
            deleteCommand.CommandText = "DELETE FROM MaterialSuggestions";
            deleteCommand.ExecuteNonQuery();
        }

        foreach (var name in names)
        {
            using var insertCommand = connection.CreateCommand();
            insertCommand.Transaction = transaction;
            insertCommand.CommandText = "INSERT INTO MaterialSuggestions (Name) VALUES ($name)";
            insertCommand.Parameters.AddWithValue("$name", name);
            insertCommand.ExecuteNonQuery();
        }
    }
}

internal sealed record AppState(OwnerData Owner, QuotationData Quotation, List<QuoteItemData> Items);

internal sealed record OwnerData(string Company, string Gstin, string Address, string Phone, string Email, string? LogoPath)
{
    public static OwnerData Default { get; } = new(
        "R.R Engineering",
        "33CZGPR1438E1ZI",
        "63/1 Mahaveer Street, Chennai, Tamil Nadu - 600050",
        "+91 90924 92393",
        "rajadhurai1998@gmail.com",
        null);
}

internal sealed record QuotationData(
    string CustomerName,
    string CustomerAddress,
    string CustomerPhone,
    string SupplyPlace,
    string QuotationNo,
    DateTime QuoteDate,
    DateTime ValidityDate,
    decimal GstPercent)
{
    public static QuotationData Default { get; } = new(
        string.Empty,
        string.Empty,
        string.Empty,
        "Tamil Nadu",
        "EST-15",
        DateTime.Today,
        DateTime.Today.AddDays(7),
        18);
}

internal sealed record QuoteItemData(int RowNo, string MaterialName, string W, string H, string Qty, string Soft, string Rate);
