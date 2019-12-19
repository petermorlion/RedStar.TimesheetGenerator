using System.IO;

namespace RedStar.TimesheetGenerator.Core
{
    public class Options
    {
        public int Year { get; set; }
        public int Month { get; set; }
        public FileInfo FileDestination { get; set; }
    }
}