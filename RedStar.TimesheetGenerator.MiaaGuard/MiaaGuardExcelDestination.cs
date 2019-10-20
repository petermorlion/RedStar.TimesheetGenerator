using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Text;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using RedStar.TimesheetGenerator.Core;

namespace RedStar.TimesheetGenerator.MiaaGuard
{
    public class MiaaGuardExcelDestination : ITimesheetDestination
    {
        private readonly FileInfo _fileDestination;
        private readonly int _month;
        private readonly int _year;

        public MiaaGuardExcelDestination(FileInfo fileDestination, int month, int year)
        {
            _fileDestination = fileDestination;
            _month = month;
            _year = year;
        }

        private void SetCell(ExcelWorksheet worksheet, string cellIdentifier, string text, int size, Color color, Color? backgroundColor = null)
        {
            var cell = worksheet.Cells[cellIdentifier];

            cell.Value = text;
            cell.Style.Font.Size = size;
            cell.Style.Font.Color.SetColor(color);
            cell.Style.Font.Name = "Helvetica Neue Light";
            if (backgroundColor.HasValue)
            {
                cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                cell.Style.Fill.BackgroundColor.SetColor(backgroundColor.Value);
            }
        }

        public void CreateTimesheet(IList<TimeTrackingEntry> entries)
        {
            using (var excelPackage = new ExcelPackage())
            {
                var worksheet = excelPackage.Workbook.Worksheets.Add("Timesheet");
                var red = Color.FromArgb(0, 128, 0, 0);
                var black = Color.Black;
                var white = Color.White;
                SetCell(worksheet, "A7", "Timesheet", 36, red);
                
                SetCell(worksheet, "A9", "Project:", 11, black);
                SetCell(worksheet, "B9", "", 11, black);
                SetCell(worksheet, "A10", "Period:", 11, black);
                SetCell(worksheet, "B10", $"{_month} {_year}", 11, black);
                SetCell(worksheet, "A11", "Consultant:", 11, black);
                SetCell(worksheet, "B11", "Peter Morlion", 11, black);
                SetCell(worksheet, "A12", "Reporting:", 11, black);
                SetCell(worksheet, "B12", $"{_month} {_year}", 11, black);
                
                SetCell(worksheet, "A15", "Week", 11, white, red);
                SetCell(worksheet, "B15", "Date", 11, white, red);
                SetCell(worksheet, "C15", "Consultant", 11, white, red);
                SetCell(worksheet, "D15", "# manhours", 11, white, red);
                SetCell(worksheet, "E15", "Task", 11, white, red);

                for (var i = 0; i < entries.Count; i++)
                {
                    var rowNumber = 16 + i;
                    var entry = entries[i];
                    var date = entry.Date;

                    SetCell(worksheet, $"B{rowNumber}", date.ToString("dd/MM/yyyy"), 11, black);
                    SetCell(worksheet, $"C{rowNumber}", "Peter Morlion", 11, black);
                    SetCell(worksheet, $"D{rowNumber}", entry.Hours.ToString(), 11, black);
                    SetCell(worksheet, $"E{rowNumber}", "Development", 11, black);
                }

                worksheet.Row(7).Height = 44.8;
                worksheet.Row(8).Height = 44.8;
                worksheet.Row(15).Height = 30;
                worksheet.Row(15).Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Column(1).Width = 10;
                worksheet.Column(2).Width = 10;
                worksheet.Column(3).Width = 10;
                worksheet.Column(4).Width = 10;
                worksheet.Column(5).Width = 52;

                var assembly = Assembly.GetExecutingAssembly();
                var resourceName = "RedStar.TimesheetGenerator.MiaaGuard.logo.png";

                using (var stream = assembly.GetManifestResourceStream(resourceName))
                {
                    var image = new Bitmap(stream);
                    var logo = worksheet.Drawings.AddPicture("logo", image);
                    logo.SetSize(65);
                    logo.SetPosition(0, 350);
                }

                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                excelPackage.SaveAs(_fileDestination);
            }
        }
    }
}
