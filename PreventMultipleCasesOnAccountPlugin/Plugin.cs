using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace PreventMultipleCasesOnAccountPlugin
{
    public class Plugin : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            // Get the execution context
            IPluginExecutionContext context = (IPluginExecutionContext) serviceProvider.GetService(typeof(IPluginExecutionContext));

            // Get the organisation service
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory) serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            // Get the tracing service
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            try
            {
                // Get the case entity (also referred to as incident in the Plugin Registration Tool)
                Entity caseEntity = (Entity) context.InputParameters["Target"];

                // Get the customer id (account or contact - it doesn't matter)
                EntityReference customerRef = (EntityReference) caseEntity ["customerid"];

                // Get the customer id
                Guid customerId = customerRef.Id;

                // Check whether the case entity has any open cases
                bool hasOpenCases = ContainsOpenCases(service, customerId);

                if (hasOpenCases)
                {
                    throw new InvalidPluginExecutionException("Only one open case is permitted per account or contact.");
                }
            } catch (Exception ex)
            {
                tracingService.Trace("PreventMultipleCasesOnAccountPlugin: {0}", ex.ToString());
                throw;
            }

        }

        // Checks whether there are open cases
        private bool ContainsOpenCases(IOrganizationService service, Guid customerId)
        {
            // Query to check for open cases
            QueryExpression query = new QueryExpression("incident")
            {
                ColumnSet = new ColumnSet("statecode", "customerid"),
                Criteria = new FilterExpression()
                {
                    Conditions =
                    {
                        new ConditionExpression("statecode", ConditionOperator.Equal, 0),
                        new ConditionExpression("customerid", ConditionOperator.Equal, customerId)
                    }
                }
            };

            // Use query to get collection containing all open cases
            EntityCollection openCases = service.RetrieveMultiple(query);
            
            // Check if collection has any cases
            if (openCases.Entities.Count > 0)
            {
                // Contains open cases
                return true;
            } else
            {
                // Doesn't contain open cases
                return false;
            }
        }

    }
}