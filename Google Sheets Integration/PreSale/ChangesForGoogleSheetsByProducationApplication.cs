using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using PreSale.Properties;
using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;

namespace PreSale
{
    public class ChangesForGoogleSheetsByProducationApplication : IPlugin
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
                Entity entity = (Entity)context.InputParameters["Target"];

                IOrganizationServiceFactory serviceFactory =
                    (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

                try
                {
                    var helper = new GoogleSheetsHelper();

                    if (context.MessageName == "Delete")
                    {
                        helper.Delete(entity.Id);
                        return;
                    }
                    
                    var producationApplication = new cr254_ProducationApplication()
                    {
                        cr254_Name = entity.GetAttributeValue<string>("cr254_name"),
                        cr254_Field1 = entity.GetAttributeValue<string>("cr254_field1"),
                        cr254_Field2 = entity.GetAttributeValue<Money>("cr254_field2"),
                        cr254_Field3 = entity.GetAttributeValue<string>("cr254_field3"),
                        Id = entity.Id
                    };
                    
                    if (context.MessageName == "Update")
                    {
                        helper.Update(entity.Id, producationApplication);
                    }
                    else if (context.MessageName == "Create")
                    {
                        helper.Create(producationApplication);
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
