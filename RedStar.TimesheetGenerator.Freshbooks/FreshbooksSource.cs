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
                .Select(x =>
                {
                    if (!x.Date.HasValue || !x.Hours.HasValue)
                    {
                        return null;
                    }

                    return new TimeTrackingEntry
                    {
                        Date = x.Date.Value,
                        Hours = x.Hours.Value
                    };
                })
                .Where(x => x != null)
                .ToList();
        }

        private FreshBooksApi Freshbooks => _freshbooks ?? (_freshbooks = FreshBooksApi.Build(_username, _token));
    }
}
