using System;
using System.IO;
using OfficeOpenXml;

namespace SQLCreate
{
    class Program
    {
        static void Main(string[] args)
        {
            ReadExcel();
        }

        static void ReadExcel()
        {
            ExcelPackage.LicenseContext = LicenseContext.Commercial;
            var dir = AppDomain.CurrentDomain.BaseDirectory;
            using (var package = new ExcelPackage(new FileInfo(@"C:\Users\aberman\source\repos\SQLCreate\Data.xlsx")))
            {
                var firstSheet = package.Workbook.Worksheets["Sheet1"];
                var rowCount = firstSheet.Dimension.End.Row;
                var colCount = firstSheet.Dimension.End.Column;
                Console.WriteLine("Sheet 1 Data");
                
                for(int row = 2; row <= rowCount; row++) // start at two for the header
                {
                    Console.WriteLine("Update public.\"Associate\"");
                    Console.WriteLine($"set \"DateOfBirth\"='{firstSheet.Cells[row, 15].Text}'");
                    Console.WriteLine($"where \"FirstName\"='{firstSheet.Cells[row, 6].Text}'");
                    Console.WriteLine($"and \"LastName\"='{firstSheet.Cells[row, 7].Text}';");
                }
                Console.WriteLine("");
                Console.ReadLine();
            }
        }
    }
}
