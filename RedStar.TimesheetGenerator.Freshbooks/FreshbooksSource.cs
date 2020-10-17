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
        private readonly string _projectIds;
        private FreshBooksApi _freshbooks;
        private DateTime _dateFrom;
        private DateTime _dateTo;

        public FreshbooksSource(Options options)
        {
            _username = Environment.GetEnvironmentVariable("freshbooks_username");
            _token = Environment.GetEnvironmentVariable("freshbooks_token");
            _projectIds = Environment.GetEnvironmentVariable("freshbooks_projectids");
            _dateFrom = new DateTime(options.Year, options.Month, 1);
            _dateTo = new DateTime(options.Year, options.Month, DateTime.DaysInMonth(options.Year, options.Month));
        }

        public string Name => "Freshbooks";

        public IList<TimeTrackingEntry> GetEntries()
        {
            var result = new List<TimeTrackingEntry>();

            var projects = Freshbooks.Projects.GetAllPages().ToDictionary(x => x.Id, x => x.Name);
            var tasks = Freshbooks.Tasks.GetAllPages().ToDictionary(x => x.Id, x => x.Name);

            var projectIdList = _projectIds.Split(",").Select(x => x.Trim()).ToList();
            foreach (var projectId in projectIdList)
            {
                try
                {
                    var freshbooksEntries = Freshbooks.TimeEntries.SearchAll(new TimeEntryFilter
                    {
                        ProjectId = projectId,
                        DateFrom = _dateFrom,
                        DateTo = _dateTo
                    });

                    var projectResult = freshbooksEntries
                        .GroupBy(x => x.Id)
                        .Select(x => x.First())
                        .Where(x => x.Date.HasValue && x.Hours.HasValue)
                        .Select(x =>
                        {
                            return new TimeTrackingEntry
                            {
                                Date = x.Date.Value,
                                Hours = x.Hours.Value,
                                Task = projects[x.ProjectId].Replace("SA | ", ""),
                                Details = tasks[x.TaskId]
                            };
                        })
                        .ToList();

                    result.AddRange(projectResult);
                }
                catch (NullReferenceException e)
                {
                    continue;
                }
                catch (Exception e)
                {
                    continue;
                }
            }

            return result.OrderBy(x => x.Date).ToList();
        }

        private FreshBooksApi Freshbooks => _freshbooks ?? (_freshbooks = FreshBooksApi.Build(_username, _token));
    }
}
