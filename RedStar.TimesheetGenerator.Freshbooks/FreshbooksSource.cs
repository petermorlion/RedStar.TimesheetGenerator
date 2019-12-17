using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
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

        public FreshbooksSource(string username, string token, string projectId)
        {
            _username = username;
            _token = token;
            _projectId = projectId;
        }

        public IList<TimeTrackingEntry> GetEntries(DateTime dateFrom, DateTime dateTo)
        {
            var freshbooksEntries = Freshbooks.TimeEntries.SearchAll(new TimeEntryFilter
            {
                ProjectId = _projectId,
                DateFrom = dateFrom,
                DateTo = dateTo
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
