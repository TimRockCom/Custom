using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Xrm.Sdk;

namespace Rimrock.Fieldboss.Common
{
    public class ContextWrapper
    {
        public IPluginExecutionContext context = null;
        public ITracingService trace = null;
        public IOrganizationService service = null;

        // plugin constructor
        public ContextWrapper(IPluginExecutionContext c, ITracingService t, IOrganizationService s)
        {
            context = c;
            trace = t;
            service = s;
        }

        // workflow contructor (omits the context)
        public ContextWrapper(ITracingService t, IOrganizationService s)
        {
            trace = t;
            service = s;
        }

        public Entity image
        {
            get
            {
                if (context.PreEntityImages != null && context.PreEntityImages.Contains("PreImage"))
                    return context.PreEntityImages["PreImage"];

                // if we are asking for an image and there isn't one, throw an exception
                throw new InvalidPluginExecutionException("No 'image' found for this entity.");
            }
        }

        public Entity target
        {
            get { return (Entity)context.InputParameters["Target"]; }
        }

        public EntityReference targetRef
        {
            get { return (EntityReference)context.InputParameters["Target"]; }
        }

        // Check this context, the parent and the grand-parent contexts 
        // to see if we are in a certain type of message (e.g. "Retrieve")
        public bool isInMessage(String msg)
        {
            bool ret = context.MessageName.Equals(msg);

            if (ret == false && context.ParentContext != null)
            {
                ret = context.ParentContext.MessageName.Equals(msg);

                if (ret == false && context.ParentContext.ParentContext != null)
                    ret = context.ParentContext.ParentContext.MessageName.Equals(msg);
            }

            return ret;
        }

        public void traceMessages()
        {
            trace.Trace("\nMessage {0} {1}", context.MessageName, context.PrimaryEntityName);

            if (context.ParentContext != null)
            {
                trace.Trace("\nParent {0} {1}", context.ParentContext.MessageName, context.ParentContext.PrimaryEntityName);

                if (context.ParentContext.ParentContext != null)
                    trace.Trace("\nGrandparent {0} {1}", context.ParentContext.ParentContext.MessageName,
                        context.ParentContext.ParentContext.PrimaryEntityName);
            }
        }

        public void traceContext()
        {
            StringBuilder s = new StringBuilder();

            s.Append(contextToString(context));

            if (context.ParentContext != null)
            {
                s.AppendLine("");
                s.AppendLine("Parent Context");
                s.AppendLine(contextToString(context.ParentContext));

                if (context.ParentContext.ParentContext != null)
                {
                    s.AppendLine("");
                    s.AppendLine("Grandparent Context");
                    s.AppendLine(contextToString(context.ParentContext.ParentContext));
                }
            }

            trace.Trace(s.ToString());
        }

        private static String contextToString(IPluginExecutionContext context)
        {
            StringBuilder s = new StringBuilder();
            s.AppendFormat(" Depth              : {0}", RRUtils.stringVal(context.Depth)); s.AppendLine();
            s.AppendFormat(" CorrelationId      : {0}", RRUtils.stringVal(context.CorrelationId)); s.AppendLine();
            s.AppendFormat(" InitiatingUserId   : {0}", RRUtils.stringVal(context.InitiatingUserId)); s.AppendLine();
            s.AppendFormat(" IsExecutingOffline : {0}", RRUtils.stringVal(context.IsExecutingOffline)); s.AppendLine();
            s.AppendFormat(" IsInTransaction    : {0}", RRUtils.stringVal(context.IsInTransaction)); s.AppendLine();
            s.AppendFormat(" IsOfflinePlayback  : {0}", RRUtils.stringVal(context.IsOfflinePlayback)); s.AppendLine();
            s.AppendFormat(" IsolationMode      : {0}", RRUtils.stringVal(context.IsolationMode)); s.AppendLine();
            s.AppendFormat(" MessageName        : {0}", RRUtils.stringVal(context.MessageName)); s.AppendLine();
            s.AppendFormat(" Mode               : {0}", RRUtils.stringVal(context.Mode)); s.AppendLine();
            s.AppendFormat(" OperationCreatedOn : {0}", RRUtils.stringVal(context.OperationCreatedOn)); s.AppendLine();
            s.AppendFormat(" OperationId        : {0}", RRUtils.stringVal(context.OperationId)); s.AppendLine();
            s.AppendFormat(" OrganizationId     : {0}", RRUtils.stringVal(context.OrganizationId)); s.AppendLine();
            s.AppendFormat(" OrganizationName   : {0}", RRUtils.stringVal(context.OrganizationName)); s.AppendLine();
            s.AppendFormat(" OwningExtension    : {0}", RRUtils.stringVal(context.OwningExtension)); s.AppendLine();
            s.AppendFormat(" ParentContext      : {0}", RRUtils.stringVal(context.ParentContext)); s.AppendLine();
            s.AppendFormat(" PrimaryEntityId    : {0}", RRUtils.stringVal(context.PrimaryEntityId)); s.AppendLine();
            s.AppendFormat(" PrimaryEntityName  : {0}", RRUtils.stringVal(context.PrimaryEntityName)); s.AppendLine();
            s.AppendFormat(" RequestId          : {0}", RRUtils.stringVal(context.RequestId)); s.AppendLine();
            s.AppendFormat(" SecondaryEntityName: {0}", RRUtils.stringVal(context.SecondaryEntityName)); s.AppendLine();
            s.AppendFormat(" Stage              : {0}", RRUtils.stringVal(context.Stage)); s.AppendLine();
            s.AppendFormat(" UserId             : {0}", RRUtils.stringVal(context.UserId)); s.AppendLine();
            s.AppendFormat(" BusinessUnitId     : {0}", RRUtils.stringVal(context.BusinessUnitId)); s.AppendLine();
            s.AppendFormat(" Type               : {0}", context.GetType().ToString()); s.AppendLine();

            s.AppendLine("");
            s.AppendLine("Input Parameters");
            foreach (KeyValuePair<String, Object> pair in context.InputParameters)
            {
                s.AppendFormat("{0} : {1}", RRUtils.stringVal(pair.Key), RRUtils.stringVal(pair.Value));
                s.AppendLine();
            }
            s.AppendLine("");
            s.AppendLine("Output Parameters");
            foreach (KeyValuePair<String, Object> pair in context.OutputParameters)
            {
                s.AppendFormat("{0} : {1}", RRUtils.stringVal(pair.Key), RRUtils.stringVal(pair.Value));
                s.AppendLine();
            }

            s.AppendLine("");
            s.AppendLine("Shared Variables");
            foreach (KeyValuePair<String, Object> pair in context.SharedVariables)
            {
                s.AppendFormat("{0} : {1}", RRUtils.stringVal(pair.Key), RRUtils.stringVal(pair.Value));
                s.AppendLine();
            }

            return s.ToString();
        }
    }
}

