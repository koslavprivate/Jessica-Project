using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;

namespace UDS.Opportunities.Actions
{
    public class UpdatingLastAndUpcomingActivityDate : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            ITracingService tracingService =
            (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            IPluginExecutionContext context = (IPluginExecutionContext)
                serviceProvider.GetService(typeof(IPluginExecutionContext));

            //Entity currentOpportunity = (Entity)context.InputParameters["Target"];

            IOrganizationServiceFactory serviceFactory =
                    (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            try
            {
                var query = new QueryExpression("opportunity")
                { ColumnSet = new ColumnSet(new string[] { "opportunityid" }) };
                var opportunities = service.RetrieveMultiple(query);

                foreach (var opportunity in opportunities.Entities)
                {
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

                        service.Update(opportunityNew);
                    }
                }
            }

            catch (FaultException<OrganizationServiceFault> ex)
            {
                throw new InvalidPluginExecutionException("An error occurred in FollowUpPlugin.", ex);
            }

            catch (Exception ex)
            {
                tracingService.Trace("FollowUpPlugin: {0}", ex.ToString());
                throw;
            }

        }
    }
}
