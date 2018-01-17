using System.Collections;
using System.Collections.Generic;

namespace RedStar.TimesheetGenerator.Core
{
    public interface ITimesheetDestination
    {
        void CreateTimesheet(IList<TimeTrackingEntry> entries);
    }
}