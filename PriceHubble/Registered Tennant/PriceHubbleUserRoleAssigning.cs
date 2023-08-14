using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;

namespace RegisteredTennant
{
    public class PriceHubbleUserRoleAssigning : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            ITracingService tracingService =
            (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            IPluginExecutionContext context = (IPluginExecutionContext)
                serviceProvider.GetService(typeof(IPluginExecutionContext));

            if (context.InputParameters.Contains("Target") &&
                context.InputParameters["Target"] is Entity)
            {
                Entity registrationTennant = (Entity)context.InputParameters["Target"];

                IOrganizationServiceFactory serviceFactory =
                    (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

                try
                {
                    var queryRole = new QueryExpression
                    {
                        EntityName = "role",
                        ColumnSet = new ColumnSet("roleid"),
                        Criteria = new FilterExpression
                        {
                            Conditions =
                            {

                                new ConditionExpression
                                {
                                    AttributeName = "name",
                                    Operator = ConditionOperator.Equal,
                                    Values = { "PriceHubble Standard User" }
                                }
                            }
                        }
                    };

                    var role = service.RetrieveMultiple(queryRole);
                    if (role.Entities.Count > 0)
                    {
                        var priceHubbleStandardUser = service.RetrieveMultiple(queryRole).Entities[0];
                        var currentUser = new Entity("systemuser", new Guid("da8011f4-4be6-ed11-a7c6-000d3a8bcaaa"));

                        if (priceHubbleStandardUser.Id != Guid.Empty)
                        {
                            service.Associate(
                                    "systemuser",
                                    currentUser.Id,
                                    new Relationship("systemuserroles_association"),
                                    new EntityReferenceCollection() { new EntityReference("role", priceHubbleStandardUser.Id) });
                        }


                        /*var priceHubbleStandardUser = service.RetrieveMultiple(queryRole).Entities[0];
                        var queryUsers = new QueryExpression("systemuser");
                        queryUsers.ColumnSet = new ColumnSet(true);
                        var users = service.RetrieveMultiple(queryUsers);

                        if (priceHubbleStandardUser.Id != Guid.Empty)
                        {
                            foreach (var user in users.Entities)
                            {
                                service.Associate(
                                    "systemuser",
                                    user.Id,
                                    new Relationship("systemuserroles_association"),
                                    new EntityReferenceCollection() { new EntityReference("role", priceHubbleStandardUser.Id) });
                            }
                        }*/
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
}
