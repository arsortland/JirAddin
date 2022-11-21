﻿using Inventor;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading.Tasks;
using System.ComponentModel;
using CadJiraForAll;
using System.Runtime.InteropServices;
using System.Collections;
using Newtonsoft.Json;
using RestSharp;
using RestSharp.Authenticators;

namespace JirAddin
{
    internal class JiraButton : Button
    {
        const string DesignTrackingGuid = "{32853F0F-3444-11D1-9E93-0060B03C1CA6}";
        const string UserDefinedGuid = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}";
        //const string ContentCenterGuid = "{B9600981-DEE8-4547-8D7C-E525B3A1727A}";
        //const string DocumentSummaryGuid = "{D5CDD502-2E9C-101B-9397-08002B2CF9AE}";
        const string SummaryGuid = "{F29F85E0-4FF9-1068-AB91-08002B27B3D9}";
        public JiraButton(string displayName, string internalName, CommandTypesEnum commandType, string clientId, string description, string tooltip, Icon standardIcon, Icon largeIcon, ButtonDisplayEnum buttonDisplayType)
    : base(displayName, internalName, commandType, clientId, description, tooltip, standardIcon, largeIcon, buttonDisplayType)
        {

        }

        public JiraButton(string displayName, string internalName, CommandTypesEnum commandType, string clientId, string description, string tooltip, ButtonDisplayEnum buttonDisplayType)
            : base(displayName, internalName, commandType, clientId, description, tooltip, buttonDisplayType)
        {

        }
        override protected async void ButtonDefinition_OnExecute(NameValueMap context)
        {
            try
            {
                //Accessing Inventor attribute values from active document in CAD:
                Document oDoc = Button.InventorApplication.ActiveDocument;
                PropertySets oPropSets = oDoc.PropertySets;
                PropertySet oPropSet = oPropSets[DesignTrackingGuid];
                Property prop = oPropSet.ItemByPropId[(int)PropertiesForDesignTrackingPropertiesEnum.kPartNumberDesignTrackingProperties];
                Property proptwo = oPropSet.ItemByPropId[(int)PropertiesForDesignTrackingPropertiesEnum.kVendorDesignTrackingProperties];

                string propstring = prop.Value.ToString(); // = Partnummer
                string proptwostring = proptwo.Value.ToString(); // = Vendor/MFG

                //SENDE ATTRIBUTTER FRA CAD/INVENTOR TIL FELLESKODE:
                CadJira felleskode = new CadJira();
                felleskode.DeliverAttributes(propstring, "Inventor");

                //FELLESKODEKJØRING:
                //felleskode.Formchoice();
                //await CadJira.API_Request(CadJira.redm_or_gcs);
                //RunAll runAll = new RunAll();
                //await runAll.NewMain();

            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }
        }
    }
}
