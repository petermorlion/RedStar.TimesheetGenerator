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

namespace RedStar.TimesheetGenerator.Excel
{
    public class ExcelDestination : ITimesheetDestination
    {
        private readonly FileInfo _fileDestination;
        private readonly int _month;
        private readonly int _year;

        public ExcelDestination(FileInfo fileDestination, int month, int year)
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

                var daysInMonth = DateTime.DaysInMonth(_year, _month);
                for (var i = 0; i < daysInMonth; i++)
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
                        worksheet.Cells[$"F{rowNumber}"].Value = entry.Hours;
                    }
                }

                worksheet.Cells[$"A17:M{20 + daysInMonth - 1}"].AddBlackBorder();

                var totalsRowIndex = 20 + daysInMonth;
                worksheet.Cells[$"D{totalsRowIndex}"].Value = "TOTAL";
                worksheet.Cells[$"D{totalsRowIndex}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells[$"D{totalsRowIndex}:E{totalsRowIndex}"].Merge = true;
                worksheet.Cells[$"F{totalsRowIndex}"].Value = entries.Sum(x => x.Hours);

                worksheet.Cells[$"F{totalsRowIndex}:M{totalsRowIndex}"].AddThickBlackBorder();

                var grandTotalRowIndex = totalsRowIndex + 1;
                worksheet.Cells[$"K{grandTotalRowIndex}:L{grandTotalRowIndex}"].Merge = true;
                worksheet.Cells[$"K{grandTotalRowIndex}:L{grandTotalRowIndex}"].Value = "TOTAL";
                worksheet.Cells[$"K{grandTotalRowIndex}:L{grandTotalRowIndex}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells[$"M{grandTotalRowIndex}"].Value = entries.Sum(x => x.Hours);
                worksheet.Cells[$"M{grandTotalRowIndex}"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[$"M{grandTotalRowIndex}"].Style.Fill.BackgroundColor.SetColor(darkBlue);
                worksheet.Cells[$"M{grandTotalRowIndex}"].Style.Font.Color.SetColor(Color.White);

                var footerRowIndex = grandTotalRowIndex + 5;
                worksheet.Cells[$"B{footerRowIndex}"].Value = "CONSULTANT";
                worksheet.Cells[$"B{footerRowIndex}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells[$"A{footerRowIndex + 1}"].Value = "NAME";
                worksheet.Cells[$"A{footerRowIndex + 1}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells[$"A{footerRowIndex + 1}"].AddBlackBorder();
                worksheet.Cells[$"B{footerRowIndex + 1}:E{footerRowIndex + 1}"].Merge = true;
                worksheet.Cells[$"B{footerRowIndex + 1}:E{footerRowIndex + 1}"].AddBlackBorder();

                worksheet.Cells[$"A{footerRowIndex + 2}"].Value = "SIGNATURE";
                worksheet.Cells[$"A{footerRowIndex + 2}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells[$"A{footerRowIndex + 2}"].AddBlackBorder();
                worksheet.Cells[$"B{footerRowIndex + 2}:E{footerRowIndex + 2}"].Merge = true;
                worksheet.Cells[$"B{footerRowIndex + 2}:E{footerRowIndex + 2}"].AddBlackBorder();

                worksheet.Cells[$"I{footerRowIndex}"].Value = "CUSTOMER";
                worksheet.Cells[$"I{footerRowIndex}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells[$"G{footerRowIndex + 1}"].Value = "NAME";
                worksheet.Cells[$"G{footerRowIndex + 1}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells[$"G{footerRowIndex + 1}:H{footerRowIndex + 1}"].AddBlackBorder();
                worksheet.Cells[$"G{footerRowIndex + 1}:H{footerRowIndex + 1}"].Merge = true;
                worksheet.Cells[$"I{footerRowIndex + 1}:M{footerRowIndex + 1}"].Merge = true;
                worksheet.Cells[$"I{footerRowIndex + 1}:M{footerRowIndex + 1}"].AddBlackBorder();

                worksheet.Cells[$"G{footerRowIndex + 2}"].Value = "SIGNATURE";
                worksheet.Cells[$"G{footerRowIndex + 2}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells[$"G{footerRowIndex + 2}:H{footerRowIndex + 2}"].AddBlackBorder();
                worksheet.Cells[$"G{footerRowIndex + 2}:H{footerRowIndex + 2}"].Merge = true;
                worksheet.Cells[$"I{footerRowIndex + 2}:M{footerRowIndex + 2}"].Merge = true;
                worksheet.Cells[$"I{footerRowIndex + 2}:M{footerRowIndex + 2}"].AddBlackBorder();

                worksheet.Cells[$"A{totalsRowIndex}:M{footerRowIndex + 2}"].Style.Font.Bold = true;

                worksheet.Row(18).Height = 53;
                worksheet.Row(footerRowIndex + 1).Height = 30;
                worksheet.Row(footerRowIndex + 2).Height = 30;

                worksheet.Column(1).Width = 10;
                worksheet.Column(2).Width = 8;
                worksheet.Column(3).Width = 11.86;
                worksheet.Column(4).Width = 7.43;
                worksheet.Column(5).Width = 1.14;
                worksheet.Column(6).Width = 5.71;
                worksheet.Column(7).Width = 4.71;
                worksheet.Column(8).Width = 4.29;
                worksheet.Column(9).Width = 5;
                worksheet.Column(10).Width = 5;
                worksheet.Column(11).Width = 5;
                worksheet.Column(12).Width = 5;

                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                excelPackage.SaveAs(_fileDestination);
            }
        }
    }
}
