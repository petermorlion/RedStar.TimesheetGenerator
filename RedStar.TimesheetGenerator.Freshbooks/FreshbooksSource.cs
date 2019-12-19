using System;
using System.Collections.Generic;
using System.Linq;
using RedStar.TimesheetGenerator.Core;
using Vooban.FreshBooks;
using Vooban.FreshBooks.TimeEntry.Models;

namespace RedStar.TimesheetGenerator.Freshbooks
{
    public class FreshbooksSource : ITimeTrackingSource
    {
        private readonly string _username;
        private readonly string _token;
        private readonly string _projectId;
        private FreshBooksApi _freshbooks;
        private DateTime _dateFrom;
        private DateTime _dateTo;

        public FreshbooksSource(Options options)
        {
            _username = Environment.GetEnvironmentVariable("freshbooks_username");
            _token = Environment.GetEnvironmentVariable("freshbooks_token");
            _projectId = Environment.GetEnvironmentVariable("freshbooks_projectid");
            _dateFrom = new DateTime(options.Year, options.Month, 1);
            _dateTo = new DateTime(options.Year, options.Month, DateTime.DaysInMonth(options.Year, options.Month));
        }

        public string Name => "Freshbooks";

        public IList<TimeTrackingEntry> GetEntries()
        {
            var freshbooksEntries = Freshbooks.TimeEntries.SearchAll(new TimeEntryFilter
            {
                ProjectId = _projectId,
                DateFrom = _dateFrom,
                DateTo = _dateTo
            });

            return freshbooksEntries
                .Where(x => x.Date.HasValue && x.Hours.HasValue)
                .GroupBy(x => x.Date)
                .Select(x =>
                {
                    return new TimeTrackingEntry
                    {
                        Date = x.Key.Value,
                        Hours = x.Sum(y => y.Hours.Value)
                    };
                })
                .ToList();
        }

        private FreshBooksApi Freshbooks => _freshbooks ?? (_freshbooks = FreshBooksApi.Build(_username, _token));
    }
}
