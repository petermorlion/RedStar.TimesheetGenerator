using System;
using System.Collections.Generic;

namespace RedStar.TimesheetGenerator.Core
{
    public interface ITimeTrackingSource
    {
        IList<TimeTrackingEntry> GetEntries(DateTime dateFrom, DateTime dateTo);
    }
}