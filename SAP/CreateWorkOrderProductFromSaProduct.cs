
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Text;
using System.ServiceModel;
using System.Threading.Tasks;
using System.Activities;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using Microsoft.Crm.Sdk.Messages;
using System.Runtime.Serialization;
using Rimrock.Fieldboss.Common;

namespace Rimrock.Fieldboss.Workflows
{
    /// </summary>    
    public class CreateWorkOrderProductFromSaProduct : WorkFlowActivityBase
    {
        #region Properties 
        //Property for Entity new_project
        [RequiredArgument]
        [Input("Service Activity Product")]
        [ReferenceTarget("fsip_serviceactivityproduct")]
        public InArgument<EntityReference> SaProductParam { get; set; }

        [RequiredArgument]
        [Input("Work Order")]
        [ReferenceTarget("salesorder")]
        public InArgument<EntityReference> WorkOrderParam { get; set; }             
        #endregion

        /// <summary>
        /// Executes the WorkFlow.
        /// </summary>
        public override void ExecuteCRMWorkFlowActivity(CodeActivityContext executionContext, LocalWorkflowContext crmWorkflowContext)
        {
            if (crmWorkflowContext == null)
                throw new ArgumentNullException("crmWorkflowContext");

            ContextWrapper contextWrapper = new ContextWrapper
            (
                crmWorkflowContext.TracingService,
                crmWorkflowContext.OrganizationService
            );

            try
            {
                // get the input parameters
                EntityReference saProdRef = SaProductParam.Get(executionContext);
                EntityReference woRef = WorkOrderParam.Get(executionContext);
               // crmWorkflowContext.Trace("instantiated workflow parameters woRef, saProdRef");

                if (saProdRef != null)
                {
                    // get the related SA product records
                    //crmWorkflowContext.Trace("saRef and woRef instantiated, now run fetch/retrievemultiple");

                    //Get all of the SA products
                    var fetchXml = @"
                    <fetch version='1.0' no-lock='true'>
                      <entity name='fsip_serviceactivityproduct'>
                        <attribute name='fsip_quantity' />
                        <attribute name='fsip_unitofmeasure' />
                        <attribute name='fsip_product' />
                        <attribute name='fsip_productname' />
                        <attribute name='fsip_serviceactivity' />
                        <attribute name='fsip_comment' />
                        <attribute name='clima_pocostperunit' />
                        <attribute name='clima_vendor' />
                        <filter>
                          <condition attribute='fsip_serviceactivityproductid' operator='eq' value='" + saProdRef.Id + @"' />                          
                          <condition attribute='statecode' operator='eq' value='0' />
                        </filter>
                      </entity>
                    </fetch>";

                    EntityCollection saProducts = contextWrapper.service.RetrieveMultiple(new FetchExpression(fetchXml));
                    //crmWorkflowContext.Trace(String.Format("retrievemultiple complete, with a count of {0}", saProducts.Entities.Count.ToString()));

                    if (saProducts.Entities.Count > 0)
                    {
                        //crmWorkflowContext.Trace("records found, about to enter for each");

                        foreach (Entity saProduct in saProducts.Entities)
                        {
                            Entity workOrderDetail = new Entity("salesorderdetail");
                            workOrderDetail["salesorderid"] = woRef;

                            FBCommon.mapValuesToEntity(workOrderDetail, "quantity", saProduct, "fsip_quantity");
                            FBCommon.mapValuesToEntity(workOrderDetail, "productid", saProduct, "fsip_product");
                            FBCommon.mapValuesToEntity(workOrderDetail, "uomid", saProduct, "fsip_unitofmeasure");
                            FBCommon.mapValuesToEntity(workOrderDetail, "fsip_committedcostperunit", saProduct, "clima_pocostperunit");
                            if (saProduct.Contains("fsip_comment"))
                            {
                                FBCommon.mapValuesToEntity(workOrderDetail, "fsip_podescription", saProduct, "fsip_comment");
                            }
                            FBCommon.mapValuesToEntity(workOrderDetail, "fsip_vendorcode", saProduct, "clima_vendor");

                            workOrderDetail["fsip_autocreateponumber"] = true;

                            //Set local time for the PO Date          
                            int? userTimeZoneCode = FBCommon.getUserTimeZoneCode(contextWrapper);
                            DateTime dateTimeLocal = FBCommon.UTCToLocalTime(DateTime.Now, userTimeZoneCode, contextWrapper);
                            workOrderDetail["fsip_purchaseorderdate"] = dateTimeLocal;
                            contextWrapper.service.Create(workOrderDetail);

                            //Submit the PO
                            Entity wo = new Entity("salesorder", "salesorderid", woRef.Id);
                            wo["fsip_submittoerp"] = true;
                            contextWrapper.service.Update(wo);
                        }
                    }
                    //crmWorkflowContext.Trace("complete, leaving CreateWorkOrderProductFromSaProduct class");
                }
            }
            catch (FaultException<OrganizationServiceFault> e)
            {
                // Handle the exception.
                throw e;
            }
        }
    }
}
