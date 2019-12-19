using System.Collections.Generic;

namespace RedStar.TimesheetGenerator.Core
{
    public interface ITimesheetDestination : IPlugin
    {
        void CreateTimesheet(IList<TimeTrackingEntry> entries);
    }
}