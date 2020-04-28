using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
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

        public MiaaGuardExcelDestination(Options options)
        {
            _fileDestination = options.FileDestination;
            _month = options.Month;
            _year = options.Year;
        }

        private void SetCell(ExcelWorksheet worksheet, string cellIdentifier, object text, int size, Color color, Color? backgroundColor = null)
        {
            var cell = worksheet.Cells[cellIdentifier];

            cell.Value = text;
            cell.Style.Font.Size = size;
            cell.Style.Font.Color.SetColor(color);
            cell.Style.Font.Name = "Montserrat Regular";
            if (backgroundColor.HasValue)
            {
                cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                cell.Style.Fill.BackgroundColor.SetColor(backgroundColor.Value);
            }
        }

        public string Name => "miaa";

        public void CreateTimesheet(IList<TimeTrackingEntry> entries)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var excelPackage = new ExcelPackage())
            {
                var worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");
                var black = Color.Black;
                var white = Color.White;
                SetCell(worksheet, "B7", "Timesheet", 36, black);
                
                SetCell(worksheet, "B9", "Project:", 11, black);
                SetCell(worksheet, "C9", "reference to Annex", 11, black);
                SetCell(worksheet, "B10", "Period:", 11, black);
                SetCell(worksheet, "C10", new DateTime(_year, _month, 1), 11, black);
                SetCell(worksheet, "B11", "Consultant:", 11, black);
                SetCell(worksheet, "C11", "Peter Morlion", 11, black);
                SetCell(worksheet, "B12", "Reporting:", 11, black);
                SetCell(worksheet, "C12", new DateTime(_year, _month, 1), 11, black);
                
                SetCell(worksheet, "B15", "Week", 11, white, black);
                SetCell(worksheet, "C15", "Date", 11, white, black);
                SetCell(worksheet, "D15", "Consultant", 11, white, black);
                SetCell(worksheet, "E15", "# manhours", 11, white, black);
                SetCell(worksheet, "F15", "Task", 11, white, black);
                SetCell(worksheet, "G15", "Details", 11, white, black);

                var orderedEntries = entries.OrderBy(x => x.Date).ToList();
                var rowNumber = 16;
                for (var i = 0; i < orderedEntries.Count; i++)
                {
                    rowNumber += 1;
                    var entry = orderedEntries[i];
                    var date = entry.Date;
                    var task = entry.Task;
                    var details = entry.Details;

                    SetCell(worksheet, $"C{rowNumber}", date, 11, black);
                    SetCell(worksheet, $"D{rowNumber}", "Peter Morlion", 11, black);
                    SetCell(worksheet, $"E{rowNumber}", Math.Round(entry.Hours, 4), 11, black);
                    SetCell(worksheet, $"F{rowNumber}", task, 11, black);
                    SetCell(worksheet, $"G{rowNumber}", details, 11, black);
                }

                rowNumber += 3;

                SetCell(worksheet, $"E{rowNumber}", "Total # manhours:", 11, black);
                worksheet.Cells[$"E{rowNumber}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                worksheet.Cells[$"F{rowNumber}"].Formula = $"=SUM(E16:E{rowNumber - 2})";
                worksheet.Cells[$"F{rowNumber}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells[$"F{rowNumber}"].Style.Font.Size = 11;
                worksheet.Cells[$"F{rowNumber}"].Style.Font.Name = "Montserrat Regular";

                SetCell(worksheet, $"E{rowNumber + 1}", "Total # mandays:", 11, black);
                worksheet.Cells[$"E{rowNumber + 1}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                worksheet.Cells[$"F{rowNumber + 1}"].Formula = $"=SUM(F{rowNumber}/8)";
                worksheet.Cells[$"F{rowNumber + 1}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells[$"F{rowNumber + 1}"].Style.Font.Size = 11;
                worksheet.Cells[$"F{rowNumber + 1}"].Style.Font.Name = "Montserrat Regular";

                SetCell(worksheet, $"E{rowNumber + 2}", "Total # days invoiced:", 11, black);
                worksheet.Cells[$"E{rowNumber + 2}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                worksheet.Cells[$"F{rowNumber + 2}"].Formula = $"=F{rowNumber + 1}";
                worksheet.Cells[$"F{rowNumber + 2}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells[$"F{rowNumber + 2}"].Style.Font.Size = 11;
                worksheet.Cells[$"F{rowNumber + 2}"].Style.Font.Name = "Montserrat Regular";

                SetRowHeights(worksheet);
                SetDateFormats(worksheet, rowNumber);
                SetColumnWidths(worksheet);
                SetTimeEntryBorders(worksheet, rowNumber);

                worksheet.Row(15).Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                excelPackage.SaveAs(_fileDestination);
            }
        }

        private void SetDateFormats(ExcelWorksheet worksheet, int numberOfRows)
        {
            worksheet.Cells["C10"].Style.Numberformat.Format = "dd/mm/yyyy";
            worksheet.Cells["C12"].Style.Numberformat.Format = "mmm-yy";

            for (var i = 16; i <= numberOfRows; i++)
            {
                worksheet.Cells[$"C{i}"].Style.Numberformat.Format = "dd/mm/yyyy";
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
            worksheet.Column(1).Width = 6;
            worksheet.Column(2).Width = 34;
            worksheet.Column(3).Width = 18;
            worksheet.Column(4).Width = 15;
            worksheet.Column(5).Width = 19;
            worksheet.Column(6).Width = 37;
            worksheet.Column(7).Width = 33;
        }

        private void SetTimeEntryBorders(ExcelWorksheet worksheet, int numberOfRows)
        {
            var borderColor = Color.Black;
            for (var i = 15; i <= numberOfRows - 3; i++)
            {
                worksheet.Cells[$"B{i}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                worksheet.Cells[$"B{i}"].Style.Border.Bottom.Color.SetColor(borderColor);
                worksheet.Cells[$"C{i}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                worksheet.Cells[$"C{i}"].Style.Border.Bottom.Color.SetColor(borderColor);
                worksheet.Cells[$"D{i}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                worksheet.Cells[$"D{i}"].Style.Border.Bottom.Color.SetColor(borderColor);
                worksheet.Cells[$"E{i}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                worksheet.Cells[$"E{i}"].Style.Border.Bottom.Color.SetColor(borderColor);
                worksheet.Cells[$"F{i}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                worksheet.Cells[$"F{i}"].Style.Border.Bottom.Color.SetColor(borderColor);
                worksheet.Cells[$"G{i}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                worksheet.Cells[$"G{i}"].Style.Border.Bottom.Color.SetColor(borderColor);
            }

            worksheet.Cells[$"B{numberOfRows - 3}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thick;
            worksheet.Cells[$"B{numberOfRows - 3}"].Style.Border.Bottom.Color.SetColor(borderColor);
            worksheet.Cells[$"C{numberOfRows - 3}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thick;
            worksheet.Cells[$"C{numberOfRows - 3}"].Style.Border.Bottom.Color.SetColor(borderColor);
            worksheet.Cells[$"D{numberOfRows - 3}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thick;
            worksheet.Cells[$"D{numberOfRows - 3}"].Style.Border.Bottom.Color.SetColor(borderColor);
            worksheet.Cells[$"E{numberOfRows - 3}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thick;
            worksheet.Cells[$"E{numberOfRows - 3}"].Style.Border.Bottom.Color.SetColor(borderColor);
            worksheet.Cells[$"F{numberOfRows - 3}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thick;
            worksheet.Cells[$"F{numberOfRows - 3}"].Style.Border.Bottom.Color.SetColor(borderColor);
            worksheet.Cells[$"G{numberOfRows - 3}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thick;
            worksheet.Cells[$"G{numberOfRows - 3}"].Style.Border.Bottom.Color.SetColor(borderColor);
        }
    }
}
