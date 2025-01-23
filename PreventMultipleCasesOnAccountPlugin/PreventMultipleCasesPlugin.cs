using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace Plugins
{
    public class PreventMultipleCasesPlugin : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            // Get the execution context
            IPluginExecutionContext context = serviceProvider.GetService(typeof(IPluginExecutionContext)) as IPluginExecutionContext;

            // Check that context exists and it contains the Target entity
            if (context == null || context.InputParameters.Contains("Target = false") || !(context.InputParameters["Target"] is Entity))
            {
                return;
            }

            // Get the organisation service
            IOrganizationServiceFactory serviceFactory = serviceProvider.GetService(typeof(IOrganizationServiceFactory)) as IOrganizationServiceFactory;
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

                
            // Get the tracing service
            ITracingService tracingService = serviceProvider.GetService(typeof(ITracingService)) as ITracingService;

            try
            {
                // Get the case entity (also referred to as incident in the Plugin Registration Tool)
                Entity caseEntity = (Entity) context.InputParameters["Target"];

                // Check that customerid attribute is present and of type EntityReference
                if (!caseEntity.Attributes.Contains("customerid") || !(caseEntity["customerid"] is EntityReference))
                {
                    return;
                }

                // Get the customer id (account or contact - it doesn't matter)
                EntityReference customerRef = caseEntity ["customerid"] as EntityReference;

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
                // Rethrow to ensure exception is logged with error details traced 
                tracingService.Trace("PreventMultipleCasesOnAccountPlugin: {0}", ex.ToString());
                throw;
            }

        }

        // Checks whether there are open cases
        private bool ContainsOpenCases(IOrganizationService service, Guid customerId)
        {
            // Query to check for one open case
            QueryExpression query = new QueryExpression("incident")
            {
                // ColumnSet does not retrieve attribute data (as only the existence of a record is needed)
                ColumnSet = new ColumnSet(false),
                Criteria = new FilterExpression()
                {
                    Conditions =
                    {
                        new ConditionExpression("statecode", ConditionOperator.Equal, 0),
                        new ConditionExpression("customerid", ConditionOperator.Equal, customerId)
                    }
                },
                // Limit to retrieve just 1 record for efficiency
                TopCount = 1
            };

            // Use query to get a collection either containing a single case or none
            EntityCollection openCases = service.RetrieveMultiple(query);
            
            // Check if collection has a case
            if (openCases.Entities.Count > 0)
            {
                // Contains an open case
                return true;
            }

            // Doesn't contain an open case
            return false;
        }

    }
}
