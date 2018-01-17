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

                worksheet.Cells["A1:I60"].Style.Font.Color.SetColor(Color.FromArgb(0, 0, 0, 128));

                worksheet.Cells["A8"].Value = "MONTHLY ACTIVITY REPORT";
                worksheet.Cells["A10"].Value = "NAME:";
                worksheet.Cells["A10"].Style.Font.Bold = true;
                worksheet.Cells["A14"].Value = "MONTH/YEAR:";
                worksheet.Cells["C14"].Value = CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(_month);
                worksheet.Cells["D14"].Value = _year;

                worksheet.Cells["A17"].Value = "DAY";
                worksheet.Cells["A17:A19"].Merge = true;

                worksheet.Cells["B17"].Value = "PROJECT";
                worksheet.Cells["B17:E19"].Merge = true;

                worksheet.Cells["F17"].Value = "Hours";
                worksheet.Cells["F17:F18"].Merge = true;
                worksheet.Cells["F19"].Value = "Week";

                worksheet.Cells["A17:L17"].Style.Font.Bold = true;

                var daysInMonth = DateTime.DaysInMonth(_year, _month);
                for (var i = 0; i < daysInMonth; i++)
                {
                    var rowNumber = 20 + i;
                    var day = 1 + i;
                    var date = new DateTime(_year, _month, day);
                    worksheet.Cells[$"A{rowNumber}"].Value = $"{date.DayOfWeek.ToString().Substring(0, 2)} {day:##}";

                    var entry = entries.SingleOrDefault(x => x.Date == date);
                    if (entry != null)
                    {
                        worksheet.Cells[$"F{rowNumber}"].Value = entry.Hours;
                    }
                }

                var totalsRowIndex = 20 + daysInMonth;
                worksheet.Cells[$"E{totalsRowIndex}"].Value = "TOTAL";
                worksheet.Cells[$"F{totalsRowIndex}"].Value = entries.Sum(x => x.Hours);

                var grandTotalRowIndex = totalsRowIndex + 1;
                worksheet.Cells[$"M{grandTotalRowIndex}"].Value = entries.Sum(x => x.Hours);

                var footerRowIndex = grandTotalRowIndex + 5;
                worksheet.Cells[$"B{footerRowIndex}"].Value = "CONSULTANT";
                worksheet.Cells[$"A{footerRowIndex + 1}"].Value = "NAME";
                worksheet.Cells[$"A{footerRowIndex + 2}"].Value = "SIGNATURE";

                worksheet.Cells[$"I{footerRowIndex}"].Value = "CUSTOMER";
                worksheet.Cells[$"G{footerRowIndex + 1}"].Value = "NAME";
                worksheet.Cells[$"G{footerRowIndex + 2}"].Value = "SIGNATURE";

                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                excelPackage.SaveAs(_fileDestination);
            }
        }
    }
}
