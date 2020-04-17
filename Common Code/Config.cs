using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rimrock.Fieldboss.Common
{
    public class Config
    {
        private bool            _approvesawhentimecardsapproved = false;
        private bool            _autogeneratebuildingid = true;
        private bool            _createProjectNumbers = false;
        private bool            _createtimecardswhensaismarkeddone = true;
        private EntityReference _defaultPaymentTerm = null;
        private EntityReference _defaulttaxscheduleforaccounts = null;
        private EntityReference _defaulttaxscheduleforfreight = null;
        private bool            _enableinvoicetax = true;
        private bool            _enablequotetax = true;
        private bool            _enableworkordertax = true;
        private int             _erpIntegration = 0;
        private int             _externalBuildingLink = 0;
        private String          _googleApiKey = String.Empty;
        private bool            _mapsaactualtimestoscheduled = false;
        private Decimal         _overheadPct = -1;
        private EntityReference _retainageProduct = null;
        private bool            _syncsascheduledtimestoactual = false;
        private bool            _useoverridecostsinwototals = false;


        public bool approvesawhentimecardsapproved
        {
            get { return _approvesawhentimecardsapproved; }
        }
        public bool autogeneratebuildingid
        {
            get { return _autogeneratebuildingid; }
        }
        public bool createProjectNumbers
        {
            get { return _createProjectNumbers; }
        }
        public bool createtimecardswhensaismarkeddone
        {
            get { return _createtimecardswhensaismarkeddone; }
        }
        public int erpIntegration
        {
            get { return _erpIntegration; }
        }
        public int externalBuildingLink
        {
            get { return _externalBuildingLink; }
        }
        public EntityReference defaultPaymentTerm
        {
            get { return _defaultPaymentTerm; }
        }
        public EntityReference defaultTaxScheduleForAccounts
        {
            get { return _defaulttaxscheduleforaccounts; }
        }
        public EntityReference defaultTaxScheduleForFreight
        {
            get { return _defaulttaxscheduleforfreight; }
        }
        public bool enableInvoiceTax
        {
            get { return _enableinvoicetax; }
        }
        public bool enableQuoteTax
        {
            get { return _enablequotetax; }
        }
        public bool enableWorkOrderTax
        {
            get { return _enableworkordertax; }
        }
        public bool mapsaactualtimestoscheduled
        {
            get { return _mapsaactualtimestoscheduled; }
        }
        public Decimal overheadPct
        {
            get { return _overheadPct; }
        }
        public EntityReference retainageProduct
        {
            get { return _retainageProduct; }
        }
        public bool syncsascheduledtimestoactual
        {
            get { return _syncsascheduledtimestoactual; }
        }
        public bool useOverrideCostsInWOTotals
        {
            get { return _useoverridecostsinwototals; }
        }
        public String googleApiKey
        {
            get { return _googleApiKey; }
        }

        public Config(ContextWrapper contextWrapper)
        {
            string fetchXml = @"
            <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false' no-lock='true'>
              <entity name='new_mapconfiguration'>
                <attribute name='fsip_approvesawhentimecardsapproved' />
                <attribute name='fsip_autogeneratebuildingid' />
                <attribute name='fsip_createprojectnumbers' />
                <attribute name='fsip_createtimecardswhensaismarkeddone' />
                <attribute name='fsip_defaultpaymentterm' />
                <attribute name='fsip_defaulttaxscheduleforaccounts' />
                <attribute name='fsip_defaulttaxscheduleforfreight' />
                <attribute name='fsip_enableinvoicetax' />
                <attribute name='fsip_enablequotetax' />
                <attribute name='fsip_enableworkordertax' />
                <attribute name='fsip_erpintegration' />
                <attribute name='fsip_externalbuildinglink' />
                <attribute name='fsip_mapsaactualtimestoscheduled' />
                <attribute name='fsip_overheadpercent' />
                <attribute name='fsip_retainageproduct' />
                <attribute name='fsip_syncsascheduledtimestoactual' />
                <attribute name='fsip_useoverridecostsinwototals' />
                <attribute name='new_googlekey' />
              </entity>
            </fetch>";

            EntityCollection collection = contextWrapper.service.RetrieveMultiple(new FetchExpression(fetchXml));

            // there must be one and only one config record
            switch (collection.Entities.Count)
            {
                case 0:
                    throw new InvalidPluginExecutionException("There is no configuration record. Please set one up in the 'Settings' area.");

                case 1:
                    Entity config = collection.Entities[0];

                    _approvesawhentimecardsapproved = RRUtils.getBoolValue(config, "fsip_approvesawhentimecardsapproved");
                    _autogeneratebuildingid = RRUtils.getBoolValue(config, "fsip_autogeneratebuildingid");
                    _createProjectNumbers = RRUtils.getBoolValue(config, "fsip_createprojectnumbers");
                    _createtimecardswhensaismarkeddone = RRUtils.getBoolValue(config, "fsip_createtimecardswhensaismarkeddone");
                    _defaultPaymentTerm = RRUtils.getEntityReference(config, "fsip_defaultpaymentterm");
                    _defaulttaxscheduleforaccounts = RRUtils.getEntityReference(config, "fsip_defaulttaxscheduleforaccounts");
                    _defaulttaxscheduleforfreight = RRUtils.getEntityReference(config, "fsip_defaulttaxscheduleforfreight");
                    _enableinvoicetax = RRUtils.getBoolValue(config, "fsip_enableinvoicetax");
                    _enablequotetax = RRUtils.getBoolValue(config, "fsip_enablequotetax");
                    _enableworkordertax = RRUtils.getBoolValue(config, "fsip_enableworkordertax");
                    _erpIntegration = RRUtils.getIntValue(config, "fsip_erpintegration");
                    _externalBuildingLink = RRUtils.getIntValue(config, "fsip_externalbuildinglink");
                    _googleApiKey = RRUtils.getStringValue(config, "new_googlekey");
                    _mapsaactualtimestoscheduled = RRUtils.getBoolValue(config, "fsip_mapsaactualtimestoscheduled");
                    _overheadPct = RRUtils.getDecimalValue(config, "fsip_overheadpercent");
                    _retainageProduct = RRUtils.getEntityReference(config, "fsip_retainageproduct");
                    _syncsascheduledtimestoactual = RRUtils.getBoolValue(config, "fsip_syncsascheduledtimestoactual");
                    _useoverridecostsinwototals = RRUtils.getBoolValue(config, "fsip_useoverridecostsinwototals"); ;

                    break;

                default:
                    throw new InvalidPluginExecutionException("More than one configuration record found. Please ensure only one Config record exisits.");
            }
        }

    }
}
