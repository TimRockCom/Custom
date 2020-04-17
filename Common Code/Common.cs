using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System.Net;
using System.IO;
using System.Xml;
using Microsoft.Crm.Sdk.Messages;


namespace Rimrock.Fieldboss.Common
{
    class FBCommon
    {
        public static readonly int BASED_ON_CUSTOMER = 683440000;
        public static readonly int NON_TAXABLE = 683440001;
        public static readonly int SPECIFIC_SCHEDULE = 683440002;

        public const int EXTERNALBUILDINGLINK_NONE = 683440000;
        public const int EXTERNALBUILDINGLINK_NYC = 683440001;

        public const int COST_CATEGORY_LABOR = 683440000;
        public const int COST_CATEGORY_FE = 683440001;
        public const int COST_CATEGORY_MATERIAL = 683440002;
        public const int COST_CATEGORY_SUBCONTRACTOR = 683440003;
        public const int COST_CATEGORY_EXPENSES = 683440004;
        public const int COST_CATEGORY_NA = 683440005;

        public const int SA_STATUS_DONE = 683440000;
        public const int SA_STATUS_COMPLETED = 8;

        public const int TIME_REG = 683440000;
        public const int TIME_OT = 683440001;
        public const int TIME_DT = 683440002;

        public const int ERP_INTEGRATION_MANUAL = 683440000;
        public const int ERP_INTEGRATION_GP = 683440001;

        public const String NYC_DOB_BUILDING_TEMPLATE = "http://a810-bisweb.nyc.gov/bisweb/PropertyProfileOverviewServlet?boro={0}&houseno={1}&street={2}";
        public const String NYC_DOB_DEVICE_TEMPLATE = "http://a810-bisweb.nyc.gov/bisweb/ElevatorDetailsForDeviceServlet?passdevicenumber={0}";

        public const String RECURRENCE_SEQ_NUM_KEY = "recurrence_series";

        public static void setEntityStatus(EntityReference entityRef, int state, int status, ContextWrapper contextWrapper)
        {
            contextWrapper.service.Execute(new SetStateRequest
            {
                EntityMoniker = entityRef,
                State = new OptionSetValue(state),
                Status = new OptionSetValue(status)
            });
        }

        public static void mapValuesToEntity(Entity toEntity, String toAttr, Entity fromEntity, String fromAttr)
        {
            if (fromEntity.Contains(fromAttr))
                toEntity[toAttr] = fromEntity[fromAttr];
        }

        public static EntityReference getRateTable(EntityReference resRef, ContextWrapper contextWrapper)
        {
            EntityReference rateTableRef = null;
            ColumnSet userCols = new ColumnSet("fsip_ratetable");
            Entity res = contextWrapper.service.Retrieve(resRef.LogicalName, resRef.Id, userCols);
            if (res != null)
                rateTableRef = RRUtils.getEntityReference(res, "fsip_ratetable");
            return rateTableRef;
        }

        //public static ServiceActivity.WageInfo getLaborRateWageInfo(EntityReference rateTableRef, ContextWrapper contextWrapper)
        //{
        //    ColumnSet cols = new ColumnSet
        //    (
        //        "fsip_overtime1laborcostrate",
        //        "fsip_overtime2laborcostrate",
        //        "fsip_paycode",
        //        "fsip_positionname",
        //        "fsip_regularlaborcostrate"
        //    );

        //    Entity laborRate = contextWrapper.service.Retrieve(rateTableRef.LogicalName, rateTableRef.Id, cols);
        //    return ServiceActivity.WageInfo.fromLaborRate(laborRate);
        //}

        //public static ServiceActivity.WageInfo getWageCategoryWageInfo(EntityReference rateTableRef, EntityReference wageCategoryRef, ContextWrapper contextWrapper)
        //{
        //    ServiceActivity.WageInfo wageInfo = null;

        //    if (wageCategoryRef != null)
        //    {

        //        String fetchXml = String.Format(@"
        //        <fetch mapping='logical'>
        //          <entity name='fsip_laborratedetail'>
        //            <filter>
        //              <condition attribute='statecode' operator='eq' value='0' />
        //              <condition attribute='fsip_laborrate' operator='eq' value='{0}' />
        //              <condition attribute='fsip_wagecategory' operator='eq' value='{1}' />
        //            </filter>
        //          </entity>
        //        </fetch>", rateTableRef.Id.ToString(), wageCategoryRef.Id.ToString());

        //        EntityCollection laborRateDetails = contextWrapper.service.RetrieveMultiple(new FetchExpression(fetchXml));
        //        if (laborRateDetails.Entities.Count == 1)
        //            wageInfo = ServiceActivity.WageInfo.fromLaborRateDetail(laborRateDetails.Entities[0]);
        //    }

        //    return wageInfo;
        //}

        //public static ServiceActivity.WageInfo getWageInfoToUse(ServiceActivity.WageInfo baseWageInfo, ServiceActivity.WageInfo wageCategoryWageInfo, String which)
        //{
        //    ServiceActivity.WageInfo wageInfo = baseWageInfo;

        //    if(wageCategoryWageInfo != null)
        //    {
        //        // compare the base wage info to the wage category wage info
        //        // in case of a tie, pick the wage category one
        //        switch(which)
        //        {
        //            case "reg":
        //                wageInfo = baseWageInfo.rateRegular > wageCategoryWageInfo.rateRegular ? baseWageInfo : wageCategoryWageInfo;
        //                break;
        //            case "OT":
        //                wageInfo = baseWageInfo.rateOT > wageCategoryWageInfo.rateOT ? baseWageInfo : wageCategoryWageInfo;
        //                break;
        //            case "DT":
        //                wageInfo = baseWageInfo.rateDT > wageCategoryWageInfo.rateDT ? baseWageInfo : wageCategoryWageInfo;
        //                break;
        //        }
        //    }

        //    return wageInfo;
        //}

        public static Entity getPriceListItem( Guid productId, Guid priceListId, Guid uomId, ContextWrapper contextWrapper)
        {
            Entity priceListItem = null;

            QueryExpression qe = new QueryExpression
            {
                EntityName = "productpricelevel",
                ColumnSet = new ColumnSet("amount", "percentage", "pricingmethodcode"),
                Criteria = new FilterExpression
                {
                    Conditions = 
                    {
                        new ConditionExpression 
                        {
                            AttributeName = "productid", Operator = ConditionOperator.Equal, Values = { productId }
                        },
                        new ConditionExpression 
                        {
                            AttributeName = "pricelevelid", Operator = ConditionOperator.Equal, Values = { priceListId }
                        },
                        new ConditionExpression 
                        {
                            AttributeName = "uomid", Operator = ConditionOperator.Equal, Values = { uomId }
                        }
                    }
                }
            };

            EntityCollection priceListItems = contextWrapper.service.RetrieveMultiple(qe);
            if (priceListItems.Entities.Count == 1)
                priceListItem = priceListItems.Entities[0];

            return priceListItem;
        }

        public static EntityCollection cloneEntityCollection(EntityCollection source)
        {
            EntityCollection ret = new EntityCollection();
            foreach (Entity sourceEntity in source.Entities)
            {
                EntityReference sourceParty = RRUtils.getEntityReference(sourceEntity, "partyid");
                Entity newParty = new Entity("activityparty");
                newParty["partyid"] = new EntityReference(sourceParty.LogicalName, sourceParty.Id);
                ret.Entities.Add(newParty);
            }

            return ret;
        }

        public static DateTime combineDateTime(DateTime dayDateTime, DateTime timeDateTime)
        {
            return new DateTime
            (
                dayDateTime.Year,
                dayDateTime.Month,
                dayDateTime.Day,
                timeDateTime.Hour,
                timeDateTime.Minute,
                timeDateTime.Second
            );
        }

        public static bool relationshipExists(String entity, String key1, Guid val1, String key2, Guid val2, IOrganizationService service)
        {
            string fetchXml = String.Format(@"
            <fetch mapping='logical' >
                <entity name='{0}'>
                <attribute name='{1}'  />
                <filter>
                    <condition attribute='{1}' operator='eq' value='{2}' />
                    <condition attribute='{3}' operator='eq' value='{4}' />
                </filter>
                </entity>
            </fetch>", entity, key1, val1.ToString(), key2, val2.ToString());

            EntityCollection collection = service.RetrieveMultiple(new FetchExpression(fetchXml));
            return collection.Entities.Count != 0;
        }

        public static int? getUserTimeZoneCode(ContextWrapper contextWrapper)
        {
            int? timeZoneCode = null;

            EntityCollection users = contextWrapper.service.RetrieveMultiple
            (
                new QueryExpression("usersettings")
                {
                    ColumnSet = new ColumnSet("timezonecode"),
                    Criteria = new FilterExpression
                    {
                        Conditions =
                        {
                            new ConditionExpression("systemuserid", ConditionOperator.EqualUserId)
                        }
                    }
                }
            );

            if( users.Entities.Count ==  1 )
            {
                Entity user = users.Entities[0];
                timeZoneCode = (int?)user.Attributes["timezonecode"];
            }

            return timeZoneCode;

        }

        // call on CRM to convert a UTC time to the local time
        public static DateTime UTCToLocalTime(DateTime utcTime, int? timeZoneCode, ContextWrapper contextWrapper)
        {
            if (!timeZoneCode.HasValue)
                return utcTime;

            var request = new LocalTimeFromUtcTimeRequest
            {
                TimeZoneCode = timeZoneCode.Value,
                UtcTime = DateTime.SpecifyKind(utcTime, DateTimeKind.Utc)
            };

            var response = (LocalTimeFromUtcTimeResponse)contextWrapper.service.Execute(request);
            return response.LocalTime;
        }

        public static DateTime LocalTimeToUTC(DateTime localTime, int? timeZoneCode, ContextWrapper contextWrapper)
        {
            if (!timeZoneCode.HasValue)
                return localTime;

            var request = new UtcTimeFromLocalTimeRequest
            {
                TimeZoneCode = timeZoneCode.Value,
                LocalTime = DateTime.SpecifyKind(localTime, DateTimeKind.Local)
            };

            var response = (UtcTimeFromLocalTimeResponse)contextWrapper.service.Execute(request);
            return response.UtcTime;
        }

        public static String buildNYCDobUrl(int borough, String address)
        {
            String str = String.Empty;

            if (!String.IsNullOrEmpty(address))
            {
                int blankPos = address.IndexOf(" ");
                String houseno = address.Substring(0, blankPos);
                String street = Uri.EscapeUriString(address.Substring(blankPos + 1));
                str = String.Format(NYC_DOB_BUILDING_TEMPLATE, borough, houseno, street);
            }

            return str;
        }

        public static String buildNYCDeviceUrl(String deviceId)
        {
            return String.Format(NYC_DOB_DEVICE_TEMPLATE, deviceId);
        }

        public static void changeEntityStatus
        (
            EntityReference entityRef,
            int stateCode,
            int statusCode,
            ContextWrapper contextWrapper
        )
        {
            changeEntityStatus(entityRef.LogicalName,entityRef.Id,stateCode,statusCode, contextWrapper);
        }
        public static void changeEntityStatus
        (
            Entity entity,
            int stateCode,
            int statusCode,
            ContextWrapper contextWrapper
        )
        {
            changeEntityStatus(entity.LogicalName, entity.Id, stateCode, statusCode, contextWrapper);
        }

        public static void changeEntityStatus
        (
            String logicalName,
            Guid id,
            int stateCode,
            int statusCode,
            ContextWrapper contextWrapper
        )
        {
            contextWrapper.service.Execute(new SetStateRequest
            {
                EntityMoniker = new EntityReference(logicalName, id),
                State = new OptionSetValue(stateCode),
                Status = new OptionSetValue(statusCode)
            });
        }

        public static void setPluginMessage(Entity target, String message, ContextWrapper contextWrapper)
        {
            // ensure the message is 200 chars or less
            if (message.Length > 200)
                message = message.Substring(0, 200);
            
            // add the message to the plugin message entity
            Entity pluginMessage = new Entity("fsip_pluginmessage");
            pluginMessage["fsip_name"] = message;
            Guid guid = contextWrapper.service.Create(pluginMessage);

            // add the reference to the target entity
            target["fsip_pluginmessage"] = new EntityReference("fsip_pluginmessage", guid);
        }

        public static void ExecuteAssociateDisassociate(ContextWrapper contextWrapper)
        {
            // see which relationship changed
            if (contextWrapper.context.InputParameters.Contains("Relationship")
                && contextWrapper.context.InputParameters.Contains("Target")
                && contextWrapper.context.InputParameters.Contains("RelatedEntities"))
            {
                Relationship relationship = (Relationship)contextWrapper.context.InputParameters["Relationship"];
                EntityReference targetEntityRef = (EntityReference)contextWrapper.context.InputParameters["Target"];
                EntityReferenceCollection relatedEntityRefs = contextWrapper.context.InputParameters["RelatedEntities"] as EntityReferenceCollection;
                EntityReference relatedEntityRef = relatedEntityRefs[0];

                String schemaName = relationship.SchemaName;

                //// Maintenance Contract and Device
                //if (relationship.SchemaName == "fsip_contract_new_servloc")
                //{
                //    // pushmepullyou - the Device can be the target or related depending on if this is "associate" or "disassociate"
                //    EntityReference deviceRef = targetEntityRef.LogicalName == "new_servloc" ? targetEntityRef : relatedEntityRef;
                //    MaintenanceContract.updateActiveContractFields(deviceRef, contextWrapper);
                //}
            }

        }
    }

    public class LatLong
    {
        public Double lattitude = 0;
        public Double longitude = 0;
        public String Response = String.Empty;
        public String googleApiKey = String.Empty;

        public String url = String.Empty;

        public bool succeeded = false;

        public LatLong( String street, String city, String state, String zip, ContextWrapper cw )
        {
            String address = Uri.EscapeDataString(String.Format("{0},{1},{2},{3}", street, city, state, zip));

            try
            {
                Config config = new Config(cw);
                googleApiKey = config.googleApiKey;

                // hardcode the Google geocode URL
                String googleGeocodeUrl = "https://maps.googleapis.com/maps/api/geocode/xml";
                url = String.Format("{0}?address={1}&key={2}",
                    googleGeocodeUrl,
                    address,
                    googleApiKey);

                HttpWebRequest webrequest = (HttpWebRequest)WebRequest.Create(url);
                webrequest.KeepAlive = false;
                webrequest.Method = "GET";

                HttpWebResponse webresponse = (HttpWebResponse)webrequest.GetResponse();
                StreamReader loResponseStream = new StreamReader(webresponse.GetResponseStream());
                Response = loResponseStream.ReadToEnd();

                if (Response.Length > 20)
                {
                    XmlDocument reader = new XmlDocument();
                    reader.LoadXml(Response);

                    XmlNode fieldNode = reader.SelectSingleNode("//GeocodeResponse/status");
                    if (fieldNode.InnerText.ToLower().Contains("ok"))
                    {
                        XmlNode latLocNode = reader.SelectSingleNode("//GeocodeResponse/result/geometry/location/lat");
                        lattitude = Double.Parse(latLocNode.InnerText);

                        XmlNode lngLocNode = reader.SelectSingleNode("//GeocodeResponse/result/geometry/location/lng");
                        longitude = Double.Parse(lngLocNode.InnerText);

                        succeeded = true;
                    }
                }
                else
                {
                }

                loResponseStream.Close();
                webresponse.Close();
            }
            catch (Exception)
            {

            }

        }

    }
}
