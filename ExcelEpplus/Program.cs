using OfficeOpenXml;

namespace ExcelEpplus
{
    internal class Program
    {
        static async Task Main(string[] args)
        {
            var path = Path.Combine(Directory.GetCurrentDirectory(), "ecel.xlsx");
            var file = new FileInfo(path);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var people = GetMockData();

            await SaveExcelFile(people, file);
        }

        private static async Task SaveExcelFile(List<Person> people, FileInfo file)
        {
            if (file.Exists)
            {
                file.Delete();
            }

            //simple table from list of data
            using var package = new ExcelPackage(file);
            
            var ws = package.Workbook.Worksheets.Add("main");

            var range = ws.Cells["A2"].LoadFromCollection(people, true);
            range.AutoFitColumns();

            // style header
            ws.Cells["A1"].Value = "Modules";
            ws.Cells["A1:C1"].Merge = true;
            ws.Column(1).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            ws.Row(1).Style.Font.Size = 24;
            ws.Row(1).Style.Font.Color.SetColor(OfficeOpenXml.Drawing.eThemeSchemeColor.Accent1);

            ws.Row(2).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            ws.Row(2).Style.Font.Bold = true;

            await package.SaveAsync();
        }

        static List<Person> GetMockData()
        {
            var output = new List<Person>()
            {
                new(){Id = 1, Name = "asdf" , Email = "asdf@sd.cad"},
                new(){Id = 2, Name = "asdf" , Email = "asdf@sd.cad"},
                new(){Id = 3, Name = "asdf" , Email = "asdf@sd.cad"},
                new(){Id = 4, Name = "asdf" , Email = "asdf@sd.cad"},
                new(){Id = 5, Name = "asdf" , Email = "asdf@sd.cad"},
                new(){Id = 6, Name = "asdf" , Email = "asdf@sd.cad"},
                new(){Id = 7, Name = "asdf" , Email = "asdf@sd.cad"},
            };
            return output;
        }
    }

    public class Person
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string Email { get; set; }
    }
}
