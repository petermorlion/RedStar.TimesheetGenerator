using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net.Http;
using System.Runtime.InteropServices;
using IdentityModel.Client;
using RedStar.TimesheetGenerator.Core;
using RestSharp;

namespace RedStar.TimesheetGenerator.Freshbooks
{
    public class FreshbooksSource : ITimeTrackingSource
    {
        private DateTime _dateFrom;
        private DateTime _dateTo;
        private string _clientId;
        private string _businessId;

        public FreshbooksSource(Options options)
        {
            _dateFrom = new DateTime(options.Year, options.Month, 1);
            _dateTo = new DateTime(options.Year, options.Month, DateTime.DaysInMonth(options.Year, options.Month));
            _clientId = Environment.GetEnvironmentVariable("freshbooks_client_id");
            _businessId = Environment.GetEnvironmentVariable("freshbooks_business_id");
        }

        public string Name => "Freshbooks";

        public IList<TimeTrackingEntry> GetEntries()
        {
            var result = new List<TimeTrackingEntry>();

            // authorize
            var clientId = Environment.GetEnvironmentVariable("freshbooks_api_client_id");
            var clientSecret = Environment.GetEnvironmentVariable("freshbooks_api_client_secret");
            var redirectUri = "https://www.redstar.be";
            var authorizationUrl = $"https://auth.freshbooks.com/service/auth/oauth/authorize?client_id={clientId}^&response_type=code^&redirect_uri={redirectUri}";
            OpenBrowser(authorizationUrl);
            Console.WriteLine("Please enter the code you'll find in your browser's address bar: ");
            var code = Console.ReadLine();

            var httpClient = new HttpClient();
            var response = httpClient.RequestAuthorizationCodeTokenAsync(new AuthorizationCodeTokenRequest
            {
                Address = "https://api.freshbooks.com/auth/oauth/token",

                ClientId = clientId,
                ClientSecret = clientSecret,

                Code = code,
                RedirectUri = redirectUri
            }).Result;


            var client = new RestClient("https://api.freshbooks.com");
            var accessToken = response.AccessToken;

            var projectsRequest = new RestRequest($"/projects/business/{_businessId}/projects");
            projectsRequest.Method = Method.GET;
            projectsRequest.AddHeader("Authorization", $"Bearer {accessToken}");
            var projectsResponse = client.Execute<ProjectsResponse>(projectsRequest);
            var projects = projectsResponse.Data.projects.ToDictionary(x => x.id, x => x.title);
            var services1 = projectsResponse.Data.projects.SelectMany(x => x.services);
            var services2 = services1.Distinct(new ServiceEqualityComparer());
            var services = services2.ToDictionary(x => x.id, x => x.name);

            var timeEntriesRequest = new RestRequest($"/timetracking/business/{_businessId}/time_entries");
            timeEntriesRequest.AddParameter("started_from", _dateFrom.ToString("u"));
            timeEntriesRequest.AddParameter("started_to", _dateTo.ToString("u"));
            timeEntriesRequest.AddParameter("client_id", _clientId);
            timeEntriesRequest.Method = Method.GET;
            timeEntriesRequest.AddHeader("Authorization", $"Bearer {accessToken}");

            var timeEntriesResponse = client.Execute<TimeEntriesResponse>(timeEntriesRequest);
            foreach (var entry in timeEntriesResponse.Data.time_entries)
            {
                result.Add(new TimeTrackingEntry
                {
                    Date = entry.started_at.Date,
                    Hours = (double)entry.duration / 60 / 60,
                    Task = projects[entry.project_id].Replace("SA | ", ""),
                    Details = services[entry.service_id]
                });
            }

            return result.OrderBy(x => x.Date).ToList();
        }

        private static void OpenBrowser(string url)
        {
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            {
                Process.Start(new ProcessStartInfo("cmd", $"/c start {url}"));
            }
            else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
            {
                Process.Start("xdg-open", url);
            }
            else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
            {
                Process.Start("open", url);
            }
            else
            {
                throw new NotSupportedException("Your platform is not supported");
            }
        }
    }
}
