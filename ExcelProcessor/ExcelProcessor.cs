using System;
using System.IO;
using System.Text;
using System.Data;
using ExcelDataReader;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Linq;
using System.Globalization;

namespace ExcelProcessor
{
    public class ExcelFileProcessor
    {
        private readonly string projectPath;
        private readonly string dbName;
        private readonly CultureInfo cultureInfo;
        private readonly Dictionary<string, string> transliterationMap;
        private readonly string currentUser;
        private readonly DateTime currentDateTime;

        public ExcelFileProcessor(string projectPath)
        {
            if (string.IsNullOrWhiteSpace(projectPath))
                throw new ArgumentNullException(nameof(projectPath));

            this.projectPath = projectPath;
            this.dbName = $"DB_{DateTime.Now:yyyyMMdd_HHmmss}";
            this.cultureInfo = CultureInfo.InvariantCulture;
            this.transliterationMap = InitTransliterationMap();
            this.currentUser = Environment.UserName;
            this.currentDateTime = DateTime.UtcNow;
        }

        private Dictionary<string, string> InitTransliterationMap()
        {
            return new Dictionary<string, string>
            {
                {"а", "a"}, {"б", "b"}, {"в", "v"}, {"г", "g"}, {"д", "d"},
                {"е", "e"}, {"ё", "yo"}, {"ж", "zh"}, {"з", "z"}, {"и", "i"},
                {"й", "y"}, {"к", "k"}, {"л", "l"}, {"м", "m"}, {"н", "n"},
                {"о", "o"}, {"п", "p"}, {"р", "r"}, {"с", "s"}, {"т", "t"},
                {"у", "u"}, {"ф", "f"}, {"х", "h"}, {"ц", "ts"}, {"ч", "ch"},
                {"ш", "sh"}, {"щ", "sch"}, {"ъ", ""}, {"ы", "y"}, {"ь", ""},
                {"э", "e"}, {"ю", "yu"}, {"я", "ya"},
                {"А", "A"}, {"Б", "B"}, {"В", "V"}, {"Г", "G"}, {"Д", "D"},
                {"Е", "E"}, {"Ё", "Yo"}, {"Ж", "Zh"}, {"З", "Z"}, {"И", "I"},
                {"Й", "Y"}, {"К", "K"}, {"Л", "L"}, {"М", "M"}, {"Н", "N"},
                {"О", "O"}, {"П", "P"}, {"Р", "R"}, {"С", "S"}, {"Т", "T"},
                {"У", "U"}, {"Ф", "F"}, {"Х", "H"}, {"Ц", "Ts"}, {"Ч", "Ch"},
                {"Ш", "Sh"}, {"Щ", "Sch"}, {"Э", "E"}, {"Ю", "Yu"}, {"Я", "Ya"}
            };
        }

        public string ConvertToSql()
        {
            var sqlBuilder = new StringBuilder();
            var excelFiles = Directory.GetFiles(projectPath, "*.xls*");

            if (excelFiles.Length == 0)
            {
                throw new FileNotFoundException("Excel файлы не найдены в рабочей папке.");
            }

            WriteHeader(sqlBuilder);

            foreach (var excelFile in excelFiles)
            {
                ProcessExcelFile(excelFile, sqlBuilder);
            }

            string sqlPath = Path.Combine(projectPath, "database_creation.sql");
            File.WriteAllText(sqlPath, sqlBuilder.ToString(), Encoding.UTF8);

            return sqlPath;
        }

        private void WriteHeader(StringBuilder sb)
        {
            sb.AppendLine("/* =============================================");
            sb.AppendLine($"   Database: {dbName}");
            sb.AppendLine($"   Created: {currentDateTime:yyyy-MM-dd HH:mm:ss}");
            sb.AppendLine($"   User: {currentUser}");
            sb.AppendLine("   ============================================= */");
            sb.AppendLine();

            sb.AppendLine("USE [master]");
            sb.AppendLine("GO");
            sb.AppendLine();

            sb.AppendLine($"IF EXISTS (SELECT name FROM sys.databases WHERE name = N'{dbName}')");
            sb.AppendLine("BEGIN");
            sb.AppendLine($"    ALTER DATABASE [{dbName}] SET SINGLE_USER WITH ROLLBACK IMMEDIATE;");
            sb.AppendLine($"    DROP DATABASE [{dbName}];");
            sb.AppendLine("END");
            sb.AppendLine("GO");
            sb.AppendLine();

            sb.AppendLine($"CREATE DATABASE [{dbName}]");
            sb.AppendLine("    COLLATE Cyrillic_General_CI_AS");
            sb.AppendLine("GO");
            sb.AppendLine();

            sb.AppendLine($"USE [{dbName}]");
            sb.AppendLine("GO");
            sb.AppendLine();
        }

        private void ProcessExcelFile(string filePath, StringBuilder sb)
        {
            try
            {
                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                    {
                        ConfigureDataTable = _ => new ExcelDataTableConfiguration()
                        {
                            UseHeaderRow = true
                        }
                    });

                    foreach (DataTable table in result.Tables)
                    {
                        if (table.Rows.Count == 0) continue;

                        string tableName = Path.GetFileNameWithoutExtension(filePath);
                        tableName = TransliterateToLatin(tableName);

                        CreateTable(table, tableName, sb);
                        InsertData(table, tableName, sb);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Ошибка при обработке файла {Path.GetFileName(filePath)}: {ex.Message}", ex);
            }
        }

        private void CreateTable(DataTable table, string tableName, StringBuilder sb)
        {
            sb.AppendLine($"-- Создание таблицы {tableName}");
            sb.AppendLine($"IF OBJECT_ID(N'[dbo].[{tableName}]', N'U') IS NOT NULL");
            sb.AppendLine($"    DROP TABLE [dbo].[{tableName}]");
            sb.AppendLine("GO");
            sb.AppendLine();

            sb.AppendLine($"CREATE TABLE [dbo].[{tableName}] (");
            sb.AppendLine("    [Id] INT IDENTITY(1,1) PRIMARY KEY");

            foreach (DataColumn col in table.Columns)
            {
                string columnName = TransliterateToLatin(col.ColumnName);
                if (string.IsNullOrWhiteSpace(columnName) || columnName == "Undefined")
                {
                    columnName = $"Column_{table.Columns.IndexOf(col)}";
                }
                string columnType = GetSqlType(col);
                sb.AppendLine($"    ,[{columnName}] {columnType} NULL -- {col.ColumnName}");
            }

            sb.AppendLine(")");
            sb.AppendLine("GO");
            sb.AppendLine();
        }

        private string GetSqlType(DataColumn column)
        {
            string columnName = column.ColumnName.ToLower();

            if (columnName.Contains("телефон") || columnName.Contains("phone"))
                return "NVARCHAR(20)";
            if (columnName.Contains("адрес"))
                return "NVARCHAR(500)";
            if (columnName.Contains("инн"))
                return "NVARCHAR(12)";
            if (columnName.Contains("рейтинг"))
                return "INT";
            if (columnName.Contains("тип") ||
                columnName.Contains("наименование") ||
                columnName.Contains("директор"))
                return "NVARCHAR(255)";

            return "NVARCHAR(255)";
        }

        private void InsertData(DataTable table, string tableName, StringBuilder sb)
        {
            var columns = new List<string>();

            // Формируем список столбцов
            foreach (DataColumn col in table.Columns)
            {
                string columnName = TransliterateToLatin(col.ColumnName);
                if (string.IsNullOrWhiteSpace(columnName) || columnName == "Undefined")
                {
                    columnName = $"Column_{table.Columns.IndexOf(col)}";
                }
                columns.Add(columnName);
            }

            // Добавляем данные построчно
            foreach (DataRow row in table.Rows)
            {
                var values = new List<string>();

                foreach (DataColumn col in table.Columns)
                {
                    var value = row[col];
                    if (value == DBNull.Value || value == null)
                    {
                        values.Add("NULL");
                        continue;
                    }

                    string strValue = value.ToString();
                    string columnName = col.ColumnName.ToLower();

                    // Обработка телефонов
                    if (columnName.Contains("телефон") || columnName.Contains("phone"))
                    {
                        values.Add($"N'{strValue}'");
                        continue;
                    }

                    // Обработка адресов
                    if (columnName.Contains("адрес"))
                    {
                        values.Add($"N'{strValue.Replace("'", "''")}'");
                        continue;
                    }

                    // Обработка ИНН
                    if (columnName.Contains("инн"))
                    {
                        values.Add($"N'{strValue}'");
                        continue;
                    }

                    // Обработка рейтинга
                    if (columnName.Contains("рейтинг"))
                    {
                        values.Add(strValue);
                        continue;
                    }

                    // Все остальные поля
                    values.Add($"N'{strValue.Replace("'", "''")}'");
                }

                // Формируем INSERT запрос
                sb.AppendLine($"INSERT INTO [dbo].[{tableName}]");
                sb.AppendLine($"([{string.Join("], [", columns)}])");
                sb.AppendLine($"VALUES ({string.Join(", ", values)})");
                sb.AppendLine("GO");
            }
            sb.AppendLine();
        }

        private string TransliterateToLatin(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return "Undefined";

            text = text.Trim();
            string result = text;

            foreach (var pair in transliterationMap)
            {
                result = result.Replace(pair.Key, pair.Value);
            }

            // Заменяем недопустимые символы
            result = Regex.Replace(result, "[^a-zA-Z0-9]", "_");

            // Убираем множественные подчеркивания
            result = Regex.Replace(result, "_+", "_");

            // Убираем подчеркивания в начале и конце
            result = result.Trim('_');

            if (string.IsNullOrWhiteSpace(result))
                return "Undefined";

            if (char.IsDigit(result[0]))
                result = "T_" + result;

            return result;
        }
    }
}