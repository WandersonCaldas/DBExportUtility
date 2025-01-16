using System.Data;
using ClosedXML.Excel;
using MySqlConnector;

class Program
{
    static async Task Main(string[] args)
    {        
        string connectionString = "Server=localhost;Port=3306;Database=db_contrato;User=root;Password=;";

        using var connection = new MySqlConnection(connectionString);
        await connection.OpenAsync();
        
        string tableQuery = @"
            SELECT TABLE_NAME 
            FROM information_schema.tables 
            WHERE TABLE_SCHEMA = 'db_contrato' 
            AND TABLE_NAME LIKE 'tbl_%'";

        using var tableCommand = new MySqlCommand(tableQuery, connection);
        using var reader = await tableCommand.ExecuteReaderAsync();

        var tableNames = new List<string>();
        while (await reader.ReadAsync())
        {
            tableNames.Add(reader.GetString("TABLE_NAME"));
        }
        reader.Close();

        foreach (var tableName in tableNames)
        {            
            string query = $"SELECT * FROM {tableName}";
            using var command = new MySqlCommand(query, connection);
            using var adapter = new MySqlDataAdapter(command);
            var dataTable = new DataTable();
            adapter.Fill(dataTable);
            
            ExportToExcel(dataTable, tableName);
        }
    }

    static void ExportToExcel(DataTable dataTable, string tableName)
    {
        string truncatedTableName = tableName.Length > 31 ? tableName.Substring(0, 31) : tableName;

        using var workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add(dataTable, truncatedTableName);
                
        string tempFolderPath = @"D:\inetpub\wwwroot\GitHub\DBExportUtility\src\temp";
        
        if (!Directory.Exists(tempFolderPath))
            Directory.CreateDirectory(tempFolderPath);
        
        string fileName = $"{truncatedTableName}.xlsx";
        string filePath = Path.Combine(tempFolderPath, fileName);
        
        workbook.SaveAs(filePath);
        Console.WriteLine($"Tabela '{tableName}' exportada para '{filePath}' com sucesso.");
    }

}
