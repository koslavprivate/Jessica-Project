using BIT.AccountMarketingActivity.Plugins;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;
using UDS.Activities.Services;

namespace UDS.Activities.Plugins
{
    public class CreateUpdateActivity : BasePlugin
    {
        public CreateUpdateActivity()
        {
            StepManager
                .NewStep()
                    .EntityName("email")
                    .Message("Create")
                    .Stage(PluginStage.PostOperation)
                    .PluginAction(ExecuteAction)
                .Register();
            StepManager
                .NewStep()
                    .EntityName("email")
                    .Message("Update")
                    .Stage(PluginStage.PostOperation)
                    .PluginAction(ExecuteAction)
                .Register();

            StepManager
                .NewStep()
                    .EntityName("appointment")
                    .Message("Create")
                    .Stage(PluginStage.PostOperation)
                    .PluginAction(ExecuteAction)
                .Register();
            StepManager
                .NewStep()
                    .EntityName("appointment")
                    .Message("Update")
                    .Stage(PluginStage.PostOperation)
                    .PluginAction(ExecuteAction)
                .Register();

            StepManager
                .NewStep()
                    .EntityName("msdyn_copilottranscript")
                    .Message("Create")
                    .Stage(PluginStage.PostOperation)
                    .PluginAction(ExecuteAction)
                .Register();
            StepManager
                .NewStep()
                    .EntityName("msdyn_copilottranscript")
                    .Message("Update")
                    .Stage(PluginStage.PostOperation)
                    .PluginAction(ExecuteAction)
                .Register();

            StepManager
                .NewStep()
                    .EntityName("msdyn_ocsession")
                    .Message("Create")
                    .Stage(PluginStage.PostOperation)
                    .PluginAction(ExecuteAction)
                .Register();
            StepManager
                .NewStep()
                    .EntityName("msdyn_ocsession")
                    .Message("Update")
                    .Stage(PluginStage.PostOperation)
                    .PluginAction(ExecuteAction)
                .Register();

            StepManager
                .NewStep()
                    .EntityName("msfp_alert")
                    .Message("Create")
                    .Stage(PluginStage.PostOperation)
                    .PluginAction(ExecuteAction)
                .Register();
            StepManager
                .NewStep()
                    .EntityName("msfp_alert")
                    .Message("Update")
                    .Stage(PluginStage.PostOperation)
                    .PluginAction(ExecuteAction)
                .Register();

            StepManager
                .NewStep()
                    .EntityName("phonecall")
                    .Message("Create")
                    .Stage(PluginStage.PostOperation)
                    .PluginAction(ExecuteAction)
                .Register();
            StepManager
                .NewStep()
                    .EntityName("phonecall")
                    .Message("Update")
                    .Stage(PluginStage.PostOperation)
                    .PluginAction(ExecuteAction)
                .Register();

            StepManager
                .NewStep()
                    .EntityName("task")
                    .Message("Create")
                    .Stage(PluginStage.PostOperation)
                    .PluginAction(ExecuteAction)
                .Register();
            StepManager
                .NewStep()
                    .EntityName("task")
                    .Message("Update")
                    .Stage(PluginStage.PostOperation)
                    .PluginAction(ExecuteAction)
                .Register();
        }

        protected void ExecuteAction(LocalPluginContext context)
        {
            if (!context.PluginExecutionContext.InputParameters.Contains("Target") || !(context.PluginExecutionContext.InputParameters["Target"] is Entity))
                return;
            var target = (Entity)context.PluginExecutionContext.InputParameters["Target"];
            try
            {
                new ActivityService(context.InitUserOrganizationService).UpdateLastAndUpcomingActivityDates(target);
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException($"An error occurred in ActivityService.UpdateLastAndUpcomingActivityDates.\n{ex.Message}", ex);
            }
        }
    }
}
