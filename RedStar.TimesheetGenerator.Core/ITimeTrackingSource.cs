﻿using System;
using System.Collections.Generic;

namespace RedStar.TimesheetGenerator.Core
{
    public interface ITimeTrackingSource : IPlugin
    {

        IList<TimeTrackingEntry> GetEntries();
    }
}