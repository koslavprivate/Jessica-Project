using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UDS.Activities.Services
{
    internal class ActivityService
    {
        private readonly IOrganizationService service;

        public ActivityService(IOrganizationService service)
        {
            this.service = service;
        }

        internal void UpdateLastAndUpcomingActivityDates(Entity target)
        {            
            var targetFull = service.Retrieve(target.LogicalName, target.Id, new ColumnSet(new string[] { "scheduledstart", "scheduledend", "regardingobjectid" }));

            var opportunityTarget = targetFull.GetAttributeValue<EntityReference>("regardingobjectid");
            if (opportunityTarget.LogicalName != "opportunity")
                return;

            var opportunity = service.Retrieve("opportunity", opportunityTarget.Id,
                        new ColumnSet(new string[] { "new_lastactivitydate", "new_upcomingactivitydate" }));

            var queryLastFetch = $@"<fetch version=""1.0"" output-format=""xml-platform"" mapping=""logical"" distinct=""false"" count=""1"">
  <entity name=""activitypointer"">
    <attribute name=""activityid"" />
    <attribute name=""scheduledend"" />
    <order attribute=""scheduledend"" descending=""true"" />
    <filter type=""and"">
      <condition attribute=""regardingobjectid"" operator=""eq"" value=""{opportunity.GetAttributeValue<Guid>("opportunityid")}"" />
      <condition attribute=""scheduledend"" operator=""on-or-before"" value=""{DateTime.UtcNow.AddDays(-1).ToString("yyyy-MM-dd")}"" />
    </filter>
  </entity>
</fetch>";
            var queryUpcomingFetch = $@"<fetch version=""1.0"" output-format=""xml-platform"" mapping=""logical"" distinct=""false"" count=""1"">
  <entity name=""activitypointer"">
    <attribute name=""activityid"" />
    <attribute name=""scheduledstart"" />
    <order attribute=""scheduledstart"" descending=""false"" />
    <filter type=""and"">
      <condition attribute=""regardingobjectid"" operator=""eq"" value=""{opportunity.GetAttributeValue<Guid>("opportunityid")}"" />
      <condition attribute=""scheduledstart"" operator=""on-or-after"" value=""{DateTime.UtcNow.AddDays(1).ToString("yyyy-MM-dd")}"" />
    </filter>
  </entity>
</fetch>";
            var lastActivityDate = service.RetrieveMultiple(new FetchExpression(queryLastFetch))
                .Entities
                .FirstOrDefault()?
                .GetAttributeValue<DateTime?>("scheduledend");

            var upcomingActivityDate = service.RetrieveMultiple(new FetchExpression(queryUpcomingFetch))
                .Entities
                .FirstOrDefault()?
                .GetAttributeValue<DateTime?>("scheduledstart");

            var opportunityNew = new Entity("opportunity", opportunity.Id);

            if (lastActivityDate != null || upcomingActivityDate != null)
            {
                if (lastActivityDate != null)
                {
                    opportunityNew["new_lastactivitydate"] = lastActivityDate;
                }

                if (upcomingActivityDate != null)
                {
                    opportunityNew["new_upcomingactivitydate"] = upcomingActivityDate;
                }
            }

            service.Update(opportunityNew);
        }
    }
}
