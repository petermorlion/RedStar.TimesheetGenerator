using System;
using System.Collections.Generic;
using RedStar.TimesheetGenerator.Core;

namespace RedStar.TimesheetGenerator.Freshbooks
{
    public class FreshbooksSource : ITimeTrackingSource
    {
        public IList<TimeTrackingEntry> GetEntries()
        {
            throw new NotImplementedException();
        }
    }
}
