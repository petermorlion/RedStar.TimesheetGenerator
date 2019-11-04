using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
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

        private void SetCell(ExcelWorksheet worksheet, string cellIdentifier, object text, int size, Color color, Color? backgroundColor = null)
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
                SetCell(worksheet, "B9", "reference to Annex", 11, black);
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

                var orderedEntries = entries.OrderBy(x => x.Date).ToList();
                for (var i = 0; i < orderedEntries.Count; i++)
                {
                    var rowNumber = 16 + i;
                    var entry = orderedEntries[i];
                    var date = entry.Date;

                    SetCell(worksheet, $"B{rowNumber}", date, 11, black);
                    SetCell(worksheet, $"C{rowNumber}", "Peter Morlion", 11, black);
                    SetCell(worksheet, $"D{rowNumber}", entry.Hours, 11, black);
                    SetCell(worksheet, $"E{rowNumber}", "Development", 11, black);
                }

                SetCell(worksheet, "D38", "Total # manhours:", 11, black);
                SetCell(worksheet, "D39", "Total # mandays:", 11, black);
                SetCell(worksheet, "D40", "Total # days invoiced:", 11, black);
                worksheet.Cells["D38"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                worksheet.Cells["D39"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                worksheet.Cells["D40"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                worksheet.Cells["E38"].Formula = "=SUM(D16:D35)";
                worksheet.Cells["E38"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells["E39"].Formula = "=E38/8";
                worksheet.Cells["E39"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells["E40"].Formula = "=E39";
                worksheet.Cells["E40"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells["F39"].Formula = "=E39*1100";

                SetRowHeights(worksheet);
                SetDateFormats(worksheet);
                SetColumnWidths(worksheet);
                SetTimeEntryBorders(worksheet);
                AddLogo(worksheet);

                worksheet.Row(15).Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                excelPackage.SaveAs(_fileDestination);
            }
        }

        private void SetDateFormats(ExcelWorksheet worksheet)
        {
            for (var i = 16; i <= 35; i++)
            {
                worksheet.Cells[$"B{i}"].Style.Numberformat.Format = "d-mmm-yy";
            }
        }

        private static void AddLogo(ExcelWorksheet worksheet)
        {
            var assembly = Assembly.GetExecutingAssembly();
            var resourceName = "RedStar.TimesheetGenerator.MiaaGuard.logo.png";

            using (var stream = assembly.GetManifestResourceStream(resourceName))
            {
                var image = new Bitmap(stream);
                var logo = worksheet.Drawings.AddPicture("logo", image);
                logo.SetSize(65);
                logo.SetPosition(0, 350);
            }
        }

        private static void SetRowHeights(ExcelWorksheet worksheet)
        {
            worksheet.Row(7).Height = 44.25;
            worksheet.Row(8).Height = 44.25;
            worksheet.Row(9).Height = 17;
            worksheet.Row(10).Height = 17;
            worksheet.Row(11).Height = 17;
            worksheet.Row(12).Height = 17;
            worksheet.Row(15).Height = 29;

            for (var i = 16; i <= 35; i++)
            {
                worksheet.Row(i).Height = 19;
            }
        }

        private static void SetColumnWidths(ExcelWorksheet worksheet)
        {
            worksheet.Column(1).Width = 11;
            worksheet.Column(2).Width = 11;
            worksheet.Column(3).Width = 11;
            worksheet.Column(4).Width = 11;
            worksheet.Column(5).Width = 52;
            worksheet.Column(6).Width = 11;
        }

        private void SetTimeEntryBorders(ExcelWorksheet worksheet)
        {
            var borderColor = Color.Black;
            for (var i = 15; i <= 34; i++)
            {
                worksheet.Cells[$"A{i}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                worksheet.Cells[$"A{i}"].Style.Border.Bottom.Color.SetColor(borderColor);
                worksheet.Cells[$"B{i}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                worksheet.Cells[$"B{i}"].Style.Border.Bottom.Color.SetColor(borderColor);
                worksheet.Cells[$"C{i}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                worksheet.Cells[$"C{i}"].Style.Border.Bottom.Color.SetColor(borderColor);
                worksheet.Cells[$"D{i}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                worksheet.Cells[$"D{i}"].Style.Border.Bottom.Color.SetColor(borderColor);
                worksheet.Cells[$"E{i}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                worksheet.Cells[$"E{i}"].Style.Border.Bottom.Color.SetColor(borderColor);
            }

            var red = Color.FromArgb(0, 128, 0, 0);
            worksheet.Cells["A35"].Style.Border.Bottom.Style = ExcelBorderStyle.Thick;
            worksheet.Cells["A35"].Style.Border.Bottom.Color.SetColor(red);
            worksheet.Cells["B35"].Style.Border.Bottom.Style = ExcelBorderStyle.Thick;
            worksheet.Cells["B35"].Style.Border.Bottom.Color.SetColor(red);
            worksheet.Cells["C35"].Style.Border.Bottom.Style = ExcelBorderStyle.Thick;
            worksheet.Cells["C35"].Style.Border.Bottom.Color.SetColor(red);
            worksheet.Cells["D35"].Style.Border.Bottom.Style = ExcelBorderStyle.Thick;
            worksheet.Cells["D35"].Style.Border.Bottom.Color.SetColor(red);
            worksheet.Cells["E35"].Style.Border.Bottom.Style = ExcelBorderStyle.Thick;
            worksheet.Cells["E35"].Style.Border.Bottom.Color.SetColor(red);
        }
    }
}
