using System;
using System.IO;
using Microsoft.Extensions.Configuration;
using OfficeOpenXml;

namespace SQLCreate
{
    class Program
    {
        private static string _tableName = string.Empty;
        private static string _filePath = string.Empty;

        static void Main(string[] args)
        {
            var configuration = new ConfigurationBuilder()
                .AddJsonFile($"appsettings.json");

            var config = configuration.Build();
            Settings settings = (Settings)config.GetSection("Settings");
            _filePath = settings.FilePath;
            _tableName = settings.TableName;
            ReadExcel();
        }

        static void ReadExcel()
        {
            var dir = AppDomain.CurrentDomain.BaseDirectory;
            using (var package = new ExcelPackage(new FileInfo(_filePath)))
            {
                var firstSheet = package.Workbook.Worksheets["Sheet1"];
                var rowCount = firstSheet.Dimension.End.Row;
                var colCount = firstSheet.Dimension.End.Column;
                
                for(int row = 2; row <= rowCount; row++) // start at two to account for the header
                {
                    Console.WriteLine($"Update public.\"{_tableName}\"");
                    Console.WriteLine($"set \"DateOfBirth\"='{firstSheet.Cells[row, 5].Text}'");
                    Console.WriteLine($"where \"FirstName\"='{firstSheet.Cells[row, 1].Text}'");
                    Console.WriteLine($"and \"LastName\"='{firstSheet.Cells[row, 2].Text}';");
                }
                Console.WriteLine("");
                Console.ReadLine();
            }
        }
    }
}
