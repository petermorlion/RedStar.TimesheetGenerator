using RedStar.TimesheetGenerator.Core;
using RedStar.TimesheetGenerator.Excel;
using RedStar.TimesheetGenerator.Freshbooks;

namespace RedStar.TimesheetGenerator.ConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            ITimeTrackingSource source = new FreshbooksSource();
            var entries = source.GetEntries();

            ITimesheetDestination destination = new ExcelDestination();
            destination.CreateTimesheet(entries);
        }
    }
}
