using System;
using RedStar.TimesheetGenerator.Core;
using RedStar.TimesheetGenerator.Excel;
using RedStar.TimesheetGenerator.Freshbooks;

namespace RedStar.TimesheetGenerator.ConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            ITimeTrackingSource source = new FreshbooksSource(args[0], args[1], args[2]);

            var dateArg = args[3];
            // TODO: validate dateArg

            var year = int.Parse(dateArg.Substring(0, 4));
            var month = int.Parse(dateArg.Substring(4, 2));
            var dateFrom = new DateTime(year, month, 1);
            var dateTo = new DateTime(year, month, DateTime.DaysInMonth(year, month));

            var entries = source.GetEntries(dateFrom, dateTo);

            ITimesheetDestination destination = new ExcelDestination();
            destination.CreateTimesheet(entries);
        }
    }
}
