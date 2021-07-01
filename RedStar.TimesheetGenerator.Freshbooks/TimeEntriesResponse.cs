using System;
using System.Collections.Generic;
using System.Text;

namespace RedStar.TimesheetGenerator.Freshbooks
{
    internal class TimeEntriesResponse
    {
        public IList<TimeEntry> time_entries { get; set; }
    }

    internal class TimeEntry
    {
        public DateTime started_at { get; set; }
        public int duration { get; set; }
        public int project_id { get; set; }
        public int service_id { get; set; }
    }
}
