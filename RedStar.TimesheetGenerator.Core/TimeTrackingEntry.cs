﻿using System;

namespace RedStar.TimesheetGenerator.Core
{
    public class TimeTrackingEntry
    {
        public DateTime Date { get; set; }
        public double Hours { get; set; }
        public string Task { get; set; }
        public string Details { get; set; }
    }
}