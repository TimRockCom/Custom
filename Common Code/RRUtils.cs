using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;

namespace Rimrock.Fieldboss.Common
{
    public class RRUtils
    {
        #region data extraction helpers

        public static DateTime? getDateValueAsLocal(Entity entity, String attribute, ContextWrapper contextWrapper)
        {
            DateTime? ret = getDateValue(entity, attribute);
            if (ret != null && ret.HasValue)
            {
                int? userTimeZoneCode = FBCommon.getUserTimeZoneCode(contextWrapper);
                DateTime dtLocal = FBCommon.UTCToLocalTime(ret.Value, userTimeZoneCode, contextWrapper);
                ret = new DateTime?(dtLocal);
            }
            return ret;
        }

        public static DateTime? getDateValue(Entity entity, String attribute)
        {
            DateTime? ret = null;
            if (entity.Contains(attribute) && entity[attribute] != null)
            {
                object val = entity[attribute];
                if (val is DateTime)
                    ret = (DateTime)val;
            }
            return ret;
        }

        public static Decimal getDecimalValue(Entity entity, String attribute)
        {
            Decimal ret = 0;
            if (entity.Contains(attribute) && entity[attribute] != null)
            {
                object val = entity[attribute];
                ret = getDecimalValue(val);
            }
            return ret;
        }

        public static Decimal getDecimalValue(object val)
        {
            Decimal ret = 0;

            if (val is Microsoft.Xrm.Sdk.Money)
                ret = ((Money)val).Value;
            else if (val is Decimal)
                ret = (Decimal)val;
            else if (val is AliasedValue)
                ret = getDecimalValue(((AliasedValue)val).Value);
            else
                ret = Convert.ToDecimal(val);

            return ret;
        }

        public static Double getDoubleValue(Entity entity, String attribute)
        {
            Double ret = 0;
            if (entity.Contains(attribute) && entity[attribute] != null)
            {
                object val = entity[attribute];
                ret = Convert.ToDouble(getDecimalValue(val));
            }
            return ret;
        }


        public static bool getBoolValue(Entity entity, String attribute)
        {
            bool ret = false;

            if (entity.Contains(attribute) && entity[attribute] != null)
                ret = entity.GetAttributeValue<bool>(attribute);

            return ret;
        }

        public static String getStringValue(Entity entity, String attribute)
        {
            String ret = String.Empty;

            if (entity.Contains(attribute) && entity[attribute] != null)
                ret = (String)entity[attribute];

            return ret;
        }

        public static int getIntValue(Entity entity, String attribute)
        {
            int ret = 0;

            if (entity.Contains(attribute) && entity[attribute] != null)
            {
                object val = entity[attribute];

                if (val is Microsoft.Xrm.Sdk.OptionSetValue)
                    ret = ((OptionSetValue)val).Value;
                else
                    ret = entity.GetAttributeValue<int>(attribute);
            }

            return ret;
        }

        public static OptionSetValue getOptionSetValue(Entity entity, String attribute)
        {
            OptionSetValue ret = null;

            if (entity.Contains(attribute) && entity[attribute] != null)
            {
                object val = entity[attribute];

                if (val is Microsoft.Xrm.Sdk.OptionSetValue)
                {
                    int opt = ((OptionSetValue)val).Value;
                    ret = new OptionSetValue(opt);
                }
            }

            return ret;
        }


        public static String getStringId(Entity entity, String attribute)
        {
            String ret = String.Empty;

            if (entity.Contains(attribute) && entity[attribute] != null)
                ret = getEntityId(entity, attribute).ToString();

            return ret;
        }

        public static Guid getEntityId(Entity entity, String attribute)
        {
            Guid ret = Guid.Empty;

            if (entity.Contains(attribute) && entity[attribute] != null)
            {
                object val = entity[attribute];
                if (val is EntityReference)
                    return ((EntityReference)entity[attribute]).Id;

                ret = (Guid)val;
            }
            return ret;
        }

        public static String getEntityName(Entity entity, String attribute)
        {
            String ret = String.Empty;
            if (entity.Contains(attribute) && entity[attribute] != null)
            {
                object val = entity[attribute];
                if (val is EntityReference)
                    ret = ((EntityReference)entity[attribute]).Name;
            }
            return ret;
        }

        public static EntityReference getEntityReference(Entity entity, String attribute)
        {
            EntityReference ret = null;
            if (entity.Contains(attribute) && entity[attribute] != null)
            {
                object val = entity[attribute];
                if (val is EntityReference)
                    ret = (EntityReference)entity[attribute];
            }
            return ret;
        }

        public static String getFormattedValue(Entity entity, String attribute)
        {
            String ret = null;
            if (entity.FormattedValues != null
                && entity.FormattedValues.Contains(attribute)
                && entity.FormattedValues[attribute] != null)
            {
                ret = RRUtils.stringVal(entity.FormattedValues[attribute]);
            }
            return ret;
        }

        // get a local option set value
        public static String getOptionName(Entity entity, string attributeName, IOrganizationService service)
        {
            int optionsetValue = getIntValue(entity, attributeName);

            string optionsetText = string.Empty;
            RetrieveAttributeRequest retrieveAttributeRequest = new RetrieveAttributeRequest();
            retrieveAttributeRequest.EntityLogicalName = entity.LogicalName;
            retrieveAttributeRequest.LogicalName = attributeName;
            retrieveAttributeRequest.RetrieveAsIfPublished = true;

            RetrieveAttributeResponse retrieveAttributeResponse =
              (RetrieveAttributeResponse)service.Execute(retrieveAttributeRequest);
            PicklistAttributeMetadata picklistAttributeMetadata =
              (PicklistAttributeMetadata)retrieveAttributeResponse.AttributeMetadata;

            OptionSetMetadata optionsetMetadata = picklistAttributeMetadata.OptionSet;

            foreach (OptionMetadata optionMetadata in optionsetMetadata.Options)
            {
                if (optionMetadata.Value == optionsetValue)
                {
                    optionsetText = optionMetadata.Label.UserLocalizedLabel.Label;
                    return optionsetText;
                }

            }
            return optionsetText;
        }

        public static String getOptionName(String optionsetName, int optionsetValue, IOrganizationService service)
        {
            string ret = string.Empty;

            RetrieveOptionSetRequest req = new RetrieveOptionSetRequest { Name = optionsetName };

            // Execute the request.
            RetrieveOptionSetResponse resp = (RetrieveOptionSetResponse)service.Execute(req);

            // Access the retrieved OptionSetMetadata.
            OptionSetMetadata retrievedOptionSetMetadata = (OptionSetMetadata)resp.OptionSetMetadata;

            // Get the current options list for the retrieved attribute.
            OptionMetadata[] optionList = retrievedOptionSetMetadata.Options.ToArray();
            foreach (OptionMetadata optionMetadata in optionList)
            {
                if (optionMetadata.Value == optionsetValue)
                {
                    ret = optionMetadata.Label.UserLocalizedLabel.Label.ToString();
                    break;
                }
            }

            return ret;
        }

        // get a global option set value
        public static String getGlobalOptionSetName(String optionsetName, int optionsetValue, IOrganizationService service)
        {
            string ret = string.Empty;

            RetrieveOptionSetRequest req = new RetrieveOptionSetRequest { Name = optionsetName };

            // Execute the request.
            RetrieveOptionSetResponse resp = (RetrieveOptionSetResponse)service.Execute(req);

            // Access the retrieved OptionSetMetadata.
            OptionSetMetadata retrievedOptionSetMetadata = (OptionSetMetadata)resp.OptionSetMetadata;

            // Get the current options list for the retrieved attribute.
            OptionMetadata[] optionList = retrievedOptionSetMetadata.Options.ToArray();
            foreach (OptionMetadata optionMetadata in optionList)
            {
                if (optionMetadata.Value == optionsetValue)
                {
                    ret = optionMetadata.Label.UserLocalizedLabel.Label.ToString();
                    break;
                }
            }

            return ret;
        }

        public static Entity combineEntities(Entity image, Entity target)
        {
            Entity entity = new Entity(target.LogicalName);

            // start with the image values
            entity.Attributes.AddRange(image.Attributes);

            // now map in the target values since they have the changed values
            foreach (KeyValuePair<string, object> attr in target.Attributes)
            {
                if (entity.Contains(attr.Key))
                    entity[attr.Key] = attr.Value;
                else
                    entity.Attributes.Add(attr);
            }

            return entity;
        }

        #endregion

        #region Trace Support

        public static void traceAttributes(ContextWrapper contextWrapper, Entity entity)
        {
            contextWrapper.trace.Trace(getAttributesAsString(entity, String.Empty));
        }

        public static void traceAttributes(ContextWrapper contextWrapper, Entity image, Entity target, Entity combined)
        {
            StringBuilder s = new StringBuilder();

            s.AppendLine(getAttributesAsString(image, "image"));
            s.AppendLine(getAttributesAsString(target, "target"));
            s.AppendLine(getAttributesAsString(combined, "combined"));

            contextWrapper.trace.Trace(s.ToString());
        }

        protected static String getAttributesAsString(Entity entity, String heading)
        {
            StringBuilder s = new StringBuilder();

            if (entity == null)
            {
                s.AppendFormat("---- Attributes for null ---- {0}", heading);
                s.AppendLine();
            }

            else
            {
                s.AppendFormat("---- Attributes for {0} ({1}) ---- {2}", entity.LogicalName, entity.Id, heading);
                s.AppendLine();

                SortedDictionary<string, object> sorted = new SortedDictionary<string, object>();
                foreach (KeyValuePair<string, object> pair in entity.Attributes)
                    if (!pair.Key.StartsWith("modified") && !pair.Key.StartsWith("created") && !pair.Key.StartsWith("own")) // not interested in these
                        sorted.Add(pair.Key, pair.Value);

                foreach (KeyValuePair<string, object> pair in sorted)
                    s.AppendLine(String.Format("{0}: {1} ", pair.Key, stringVal(pair.Value)));
            }

            return s.ToString();
        }

        public static String stringVal(object val)
        {
            string ret = "null";
            if (val != null)
            {
                if (val is Microsoft.Xrm.Sdk.EntityReference)
                {
                    EntityReference er = (Microsoft.Xrm.Sdk.EntityReference)val;
                    ret = String.Format("{0} ({1}) {2}", er.Name, er.LogicalName, er.Id);
                }

                else if (val is Microsoft.Xrm.Sdk.Money)
                    ret = ((Microsoft.Xrm.Sdk.Money)val).Value.ToString();

                else if (val is Microsoft.Xrm.Sdk.OptionSetValue)
                    ret = ((Microsoft.Xrm.Sdk.OptionSetValue)val).Value.ToString();

                else if (val is Microsoft.Xrm.Sdk.AliasedValue)
                    ret = stringVal(((AliasedValue)val).Value);

                else
                    ret = val.ToString();
            }

            return ret;
        }

        #endregion
        
        #region Sequential Numbers

        /// <summary>
        /// return the next sequence number for the given type "which".
        /// If this is the first time we've seen "which", create a new record in the "Sequence number" entity
        /// </summary>
        /// <param name="which">Type of sequence number to return, e.g. "line number", "SA", "milestone" etc.</param>
        /// <returns>the next sequence number (0-based)</returns>

        public static Object seqNumberLock = new Object();

        public static String getNextSequenceNumber(String which, ContextWrapper contextWrapper)
        {
            String nextSeqNum = String.Empty;

            String fetchXml = String.Format(@"
                <fetch distinct='false' mapping='logical' > 
                  <entity name='fsip_sequentialnumber'> 
                    <attribute name='fsip_prefix' /> 
                    <attribute name='fsip_nextnumber' /> 
                    <filter type='and'>
                        <condition attribute='fsip_name' operator='eq' value='{0}' />
                    </filter>
                  </entity> 
                </fetch>", which);

            // NOTE: This lock is useless on distributed servers
            lock (seqNumberLock)
            {
                EntityCollection seqnums = contextWrapper.service.RetrieveMultiple(new FetchExpression(fetchXml));
                if (seqnums.Entities.Count == 0)
                    throw new InvalidPluginExecutionException("Please set up an entry in the Sequential Number entity for " + which);

                if (seqnums.Entities.Count == 1)
                {
                    Entity seqnum = seqnums.Entities[0];

                    Decimal nextNumber = getDecimalValue(seqnum, "fsip_nextnumber");
                    String prefix = getStringValue(seqnum, "fsip_prefix");
                    nextSeqNum = String.Format("{0}-{1:0000000}", prefix, nextNumber);

                    seqnum["fsip_nextnumber"] = ++nextNumber;
                    contextWrapper.service.Update(seqnum);
                }
            }

            return nextSeqNum;
        }

        #endregion
    }
}
