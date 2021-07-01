using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Text;

namespace RedStar.TimesheetGenerator.Freshbooks
{
    internal class ProjectsResponse
    {
        public IList<Project> projects { get; set; }
    }

    internal class Project
    {
        public int id { get; set; }
        public string title { get; set; }
        public IList<Service> services { get; set; }
    }

    internal class Service
    {
        public int id { get; set; }
        public string name { get; set; }
    }

    internal class ServiceEqualityComparer : IEqualityComparer<Service>
    {
        public bool Equals([AllowNull] Service x, [AllowNull] Service y)
        {
            return x != null && y != null && x.id == y.id;
        }

        public int GetHashCode([DisallowNull] Service obj)
        {
            return obj.id.GetHashCode();
        }
    }
}
