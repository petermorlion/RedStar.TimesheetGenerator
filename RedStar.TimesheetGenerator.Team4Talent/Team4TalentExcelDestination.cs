using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using RedStar.TimesheetGenerator.Core;

namespace RedStar.TimesheetGenerator.Team4Talent
{
    public class Team4TalentExcelDestination : ITimesheetDestination
    {
        private readonly FileInfo _fileDestination;
        private readonly int _month;
        private readonly int _year;

        private int _daysInMonth;
        private int _totalsRowIndex;
        private int _footerRowIndex;
        private int _grandTotalRowIndex;

        public Team4TalentExcelDestination(FileInfo fileDestination, int month, int year)
        {
            _fileDestination = fileDestination;
            _month = month;
            _year = year;
        }

        public void CreateTimesheet(IList<TimeTrackingEntry> entries)
        {
            using (var excelPackage = new ExcelPackage())
            {
                var worksheet = excelPackage.Workbook.Worksheets.Add("Timesheet");

                CalculateRowIndexes();
                SetRowHeights(worksheet);
                SetPageBorders(worksheet);
                MergeProjectCells(worksheet);

                var darkBlue = Color.FromArgb(0, 0, 0, 128);
                worksheet.Cells["A1:M60"].Style.Font.Color.SetColor(darkBlue);

                worksheet.Cells["A8"].Value = "MONTHLY ACTIVITY REPORT";
                worksheet.Cells["A8"].Style.Font.Size = 14;
                worksheet.Cells["A8"].Style.Font.Bold = true;

                worksheet.Cells["I9"].Value = "Mail:";
                worksheet.Cells["I9"].Style.Font.Bold = true;
                worksheet.Cells["J9"].Value = "invoice@team4talent.be";
            
                worksheet.Cells["A10"].Value = "NAME:";
                worksheet.Cells["A10"].Style.Font.Bold = true;
                worksheet.Cells["A14"].Value = "MONTH/YEAR:";
                worksheet.Cells["C14"].Value = CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(_month);
                worksheet.Cells["D14"].Value = _year;

                worksheet.Cells["A17"].Value = "DAY";
                worksheet.Cells["A17"].CenterHorizontallAndVertically();
                worksheet.Cells["A17:A19"].Merge = true;

                worksheet.Cells["B17"].Value = "PROJECT";
                worksheet.Cells["B17"].CenterHorizontallAndVertically();
                worksheet.Cells["B17:E19"].Merge = true;

                worksheet.Cells["F17"].Value = "Hours";
                worksheet.Cells["F17"].CenterHorizontallAndVertically();
                worksheet.Cells["F17:F18"].Merge = true;
                worksheet.Cells["F19"].Value = "Week";

                worksheet.Cells["G17"].Value = "Overtime";
                worksheet.Cells["G17"].CenterHorizontallAndVertically();
                worksheet.Cells["G17:H18"].Merge = true;
                worksheet.Cells["G19"].Value = "Sat";
                worksheet.Cells["H19"].Value = "Sun";

                worksheet.Cells["I17"].Value = "Holiday";
                worksheet.Cells["I17"].CenterHorizontallAndVertically();
                worksheet.Cells["I17:I18"].Merge = true;
                worksheet.Cells["I17"].Style.TextRotation = 90;

                worksheet.Cells["J17"].Value = "Legal Holiday";
                worksheet.Cells["J17"].CenterHorizontallAndVertically();
                worksheet.Cells["J17:J18"].Merge = true;
                worksheet.Cells["J17"].Style.TextRotation = 90;

                worksheet.Cells["K17"].Value = "Training";
                worksheet.Cells["K17"].CenterHorizontallAndVertically();
                worksheet.Cells["K17:K18"].Merge = true;
                worksheet.Cells["K17"].Style.TextRotation = 90;

                worksheet.Cells["L17"].Value = "Sickness";
                worksheet.Cells["L17"].CenterHorizontallAndVertically();
                worksheet.Cells["L17:L18"].Merge = true;
                worksheet.Cells["L17"].Style.TextRotation = 90;

                worksheet.Cells["M17"].Value = "Others";
                worksheet.Cells["M17"].CenterHorizontallAndVertically();
                worksheet.Cells["M17:M18"].Merge = true;
                worksheet.Cells["M17"].Style.TextRotation = 90;

                worksheet.Cells["A17:M17"].Style.Font.Bold = true;
                
                for (var i = 0; i < _daysInMonth; i++)
                {
                    var rowNumber = 20 + i;
                    var day = 1 + i;
                    var date = new DateTime(_year, _month, day);
                    worksheet.Cells[$"A{rowNumber}"].Value = $"{date.DayOfWeek.ToString().Substring(0, 2)} {day:##}";
                    worksheet.Cells[$"A{rowNumber}"].Style.Font.Size = 8;
                    worksheet.Cells[$"A{rowNumber}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    if (date.DayOfWeek == DayOfWeek.Saturday || date.DayOfWeek == DayOfWeek.Sunday)
                    {
                        worksheet.Cells[$"A{rowNumber}:M{rowNumber}"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        worksheet.Cells[$"A{rowNumber}:M{rowNumber}"].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                    }

                    var entry = entries.SingleOrDefault(x => x.Date == date);
                    if (entry != null)
                    {
                        worksheet.Cells[$"F{rowNumber}"].Value = Math.Round(entry.Hours, 2);
                        worksheet.Cells[$"F{rowNumber}"].Style.Font.Size = 8;
                    }
                }

                worksheet.Cells[$"A17:M{20 + _daysInMonth - 1}"].AddBlackBorder();

                worksheet.Cells[$"D{_totalsRowIndex}"].Value = "TOTAL";
                worksheet.Cells[$"D{_totalsRowIndex}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells[$"D{_totalsRowIndex}:E{_totalsRowIndex}"].Merge = true;
                worksheet.Cells[$"F{_totalsRowIndex}"].Value = entries.Sum(x => Math.Round(x.Hours, 2));

                worksheet.Cells[$"F{_totalsRowIndex}:M{_totalsRowIndex}"].AddThickBlackBorder();

                worksheet.Cells[$"K{_grandTotalRowIndex}:L{_grandTotalRowIndex}"].Merge = true;
                worksheet.Cells[$"K{_grandTotalRowIndex}:L{_grandTotalRowIndex}"].Value = "TOTAL";
                worksheet.Cells[$"K{_grandTotalRowIndex}:L{_grandTotalRowIndex}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells[$"M{_grandTotalRowIndex}"].Value = entries.Sum(x => Math.Round(x.Hours, 2));
                worksheet.Cells[$"M{_grandTotalRowIndex}"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[$"M{_grandTotalRowIndex}"].Style.Fill.BackgroundColor.SetColor(darkBlue);
                worksheet.Cells[$"M{_grandTotalRowIndex}"].Style.Font.Color.SetColor(Color.White);

                worksheet.Cells[$"B{_footerRowIndex}"].Value = "CONSULTANT";
                worksheet.Cells[$"B{_footerRowIndex}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells[$"A{_footerRowIndex + 1}"].Value = "NAME";
                worksheet.Cells[$"A{_footerRowIndex + 1}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells[$"A{_footerRowIndex + 1}"].AddBlackBorder();
                worksheet.Cells[$"B{_footerRowIndex + 1}:E{_footerRowIndex + 1}"].Merge = true;
                worksheet.Cells[$"B{_footerRowIndex + 1}:E{_footerRowIndex + 1}"].AddBlackBorder();

                worksheet.Cells[$"A{_footerRowIndex + 2}"].Value = "SIGNATURE";
                worksheet.Cells[$"A{_footerRowIndex + 2}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells[$"A{_footerRowIndex + 2}"].AddBlackBorder();
                worksheet.Cells[$"B{_footerRowIndex + 2}:E{_footerRowIndex + 2}"].Merge = true;
                worksheet.Cells[$"B{_footerRowIndex + 2}:E{_footerRowIndex + 2}"].AddBlackBorder();

                worksheet.Cells[$"I{_footerRowIndex}"].Value = "CUSTOMER";
                worksheet.Cells[$"I{_footerRowIndex}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells[$"G{_footerRowIndex + 1}"].Value = "NAME";
                worksheet.Cells[$"G{_footerRowIndex + 1}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells[$"G{_footerRowIndex + 1}:H{_footerRowIndex + 1}"].AddBlackBorder();
                worksheet.Cells[$"G{_footerRowIndex + 1}:H{_footerRowIndex + 1}"].Merge = true;
                worksheet.Cells[$"I{_footerRowIndex + 1}:M{_footerRowIndex + 1}"].Merge = true;
                worksheet.Cells[$"I{_footerRowIndex + 1}:M{_footerRowIndex + 1}"].AddBlackBorder();

                worksheet.Cells[$"G{_footerRowIndex + 2}"].Value = "SIGNATURE";
                worksheet.Cells[$"G{_footerRowIndex + 2}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells[$"G{_footerRowIndex + 2}:H{_footerRowIndex + 2}"].AddBlackBorder();
                worksheet.Cells[$"G{_footerRowIndex + 2}:H{_footerRowIndex + 2}"].Merge = true;
                worksheet.Cells[$"I{_footerRowIndex + 2}:M{_footerRowIndex + 2}"].Merge = true;
                worksheet.Cells[$"I{_footerRowIndex + 2}:M{_footerRowIndex + 2}"].AddBlackBorder();

                worksheet.Cells[$"A{_totalsRowIndex}:M{_footerRowIndex + 2}"].Style.Font.Bold = true;
                
                SetColumnWidths(worksheet);

                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                excelPackage.SaveAs(_fileDestination);
            }
        }

        private static void SetColumnWidths(ExcelWorksheet worksheet)
        {
            worksheet.Column(1).Width = 10;
            worksheet.Column(2).Width = 8.12;
            worksheet.Column(3).Width = 11.88;
            worksheet.Column(4).Width = 7.41;
            worksheet.Column(5).Width = 1.24;
            worksheet.Column(6).Width = 5.82;
            worksheet.Column(7).Width = 6;
            worksheet.Column(8).Width = 6;
            worksheet.Column(9).Width = 5.5;
            worksheet.Column(10).Width = 5.5;
            worksheet.Column(11).Width = 5.5;
            worksheet.Column(12).Width = 5.5;
            worksheet.Column(13).Width = 5.5;
        }

        private void MergeProjectCells(ExcelWorksheet worksheet)
        {
            for (var i = 20; i <= _totalsRowIndex - 1; i++)
            {
                worksheet.Cells[$"B{i}:E{i}"].Merge = true;
            }
        }

        private static void SetPageBorders(ExcelWorksheet worksheet)
        {
            var topAndBottomPageMargin = (decimal) (1.5 / 2.54);
            var leftAndRightPageMargin = (decimal) (1 / 2.54);
            worksheet.PrinterSettings.BottomMargin = topAndBottomPageMargin;
            worksheet.PrinterSettings.TopMargin = topAndBottomPageMargin;
            worksheet.PrinterSettings.LeftMargin = leftAndRightPageMargin;
            worksheet.PrinterSettings.RightMargin = leftAndRightPageMargin;
        }

        private void CalculateRowIndexes()
        {
            _daysInMonth = DateTime.DaysInMonth(_year, _month);
            _totalsRowIndex = 20 + _daysInMonth;
            _grandTotalRowIndex = _totalsRowIndex + 1;
            _footerRowIndex = _grandTotalRowIndex + 2;
        }

        private void SetRowHeights(ExcelWorksheet worksheet)
        {
            for (var i = 1; i <= 7; i++)
            {
                worksheet.Row(i).Height = 5;
            }

            worksheet.Row(8).Height = 17.7;

            for (var i = 9; i <= 16; i++)
            {
                worksheet.Row(i).Height = 12.7;
            }

            worksheet.Row(17).Height = 10.4;
            worksheet.Row(18).Height = 53;

            for (var i = 19; i <= _footerRowIndex; i++)
            {
                worksheet.Row(i).Height = 11.3;
            }

            worksheet.Row(_footerRowIndex + 1).Height = 30;
            worksheet.Row(_footerRowIndex + 2).Height = 30;
        }
    }
}
