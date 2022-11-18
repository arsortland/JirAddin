using Inventor;
using Microsoft.Win32;
using System;
using System.Windows.Forms;
using System.Drawing;
using System.Runtime.InteropServices;

namespace JirAddin
{
    /// <summary>
    /// This is the primary AddIn Server class that implements the ApplicationAddInServer interface
    /// that all Inventor AddIns are required to implement. The communication between Inventor and
    /// the AddIn is via the methods on this interface.
    /// </summary>
    [GuidAttribute("7fc160c6-87b9-43a9-9a21-fafaa5fc5fe5")]
    public class StandardAddInServer : Inventor.ApplicationAddInServer
    {

        //Inventor application object
        private Inventor.Application m_inventorApplication;

        //button
        private JiraButton jiraButton;


        //user interface event
        private UserInterfaceEvents m_userInterfaceEvents;

        // ribbon panel
        RibbonPanel m_partSketchSlotRibbonPanel;

        //event handler delegates
        private Inventor.UserInterfaceEventsSink_OnResetCommandBarsEventHandler UserInterfaceEventsSink_OnResetCommandBarsEventDelegate;
        private Inventor.UserInterfaceEventsSink_OnResetEnvironmentsEventHandler UserInterfaceEventsSink_OnResetEnvironmentsEventDelegate;
        private Inventor.UserInterfaceEventsSink_OnResetRibbonInterfaceEventHandler UserInterfaceEventsSink_OnResetRibbonInterfaceEventDelegate;


        public StandardAddInServer()
        {
        }

        #region ApplicationAddInServer Members

        public void Activate(Inventor.ApplicationAddInSite addInSiteObject, bool firstTime)
        {
            try
            {
                //the Activate method is called by Inventor when it loads the addin
                //the AddInSiteObject provides access to the Inventor Application object
                //the FirstTime flag indicates if the addin is loaded for the first time

                //initialize AddIn members
                m_inventorApplication = addInSiteObject.Application;
                Button.InventorApplication = m_inventorApplication;

                //initialize event delegates
                m_userInterfaceEvents = m_inventorApplication.UserInterfaceManager.UserInterfaceEvents;

                UserInterfaceEventsSink_OnResetCommandBarsEventDelegate = new UserInterfaceEventsSink_OnResetCommandBarsEventHandler(UserInterfaceEvents_OnResetCommandBars);
                m_userInterfaceEvents.OnResetCommandBars += UserInterfaceEventsSink_OnResetCommandBarsEventDelegate;

                UserInterfaceEventsSink_OnResetEnvironmentsEventDelegate = new UserInterfaceEventsSink_OnResetEnvironmentsEventHandler(UserInterfaceEvents_OnResetEnvironments);
                m_userInterfaceEvents.OnResetEnvironments += UserInterfaceEventsSink_OnResetEnvironmentsEventDelegate;

                UserInterfaceEventsSink_OnResetRibbonInterfaceEventDelegate = new UserInterfaceEventsSink_OnResetRibbonInterfaceEventHandler(UserInterfaceEvents_OnResetRibbonInterface);
                m_userInterfaceEvents.OnResetRibbonInterface += UserInterfaceEventsSink_OnResetRibbonInterfaceEventDelegate;

                //load image icons for UI items
                Icon jiraButtonIcon = new Icon("Resources/JiraButton.ico");

                //retrieve the GUID for this class
                GuidAttribute jiraCLSID;
                jiraCLSID = (GuidAttribute)GuidAttribute.GetCustomAttribute(typeof(StandardAddInServer), typeof(GuidAttribute));
                string jiraCLSIDString;
                jiraCLSIDString = "{" + jiraCLSID.Value + "}";


                //create button
                jiraButton = new JiraButton(
                    "Create Ticket", "Autodesk:JIRAddIn:SendToJIRABtn", CommandTypesEnum.kShapeEditCmdType,
                    jiraCLSIDString, "Adds option for JIRA ticket creation",
                    "Jira Integration", jiraButtonIcon, jiraButtonIcon, ButtonDisplayEnum.kDisplayTextInLearningMode);


                //create the command category
                CommandCategory slotCmdCategory = m_inventorApplication.CommandManager.CommandCategories.Add("JirAddin", "Autodesk:JIRAddIn:SendToJIRABtn", jiraCLSIDString);

                slotCmdCategory.Add(jiraButton.ButtonDefinition);


                if (firstTime == true)
                {
                    //access user interface manager
                    UserInterfaceManager userInterfaceManager;
                    userInterfaceManager = m_inventorApplication.UserInterfaceManager;

                    InterfaceStyleEnum interfaceStyle;
                    interfaceStyle = userInterfaceManager.InterfaceStyle;

                    //create the UI for classic interface
                    if (interfaceStyle == InterfaceStyleEnum.kClassicInterface)
                    {
                        //create toolbar
                        CommandBar slotCommandBar;
                        slotCommandBar = userInterfaceManager.CommandBars.Add("JIRA", "Autodesk:JIRAddIn:SendToJIRABtn", CommandBarTypeEnum.kRegularCommandBar, jiraCLSIDString);

                        //add buttons to toolbar
                        slotCommandBar.Controls.AddButton(jiraButton.ButtonDefinition, 0);

                        //Get the 2d sketch environment base object
                        Inventor.Environment partSketchEnvironment;
                        partSketchEnvironment = userInterfaceManager.Environments["PMxPartSketchEnvironment"];

                        //make this command bar accessible in the panel menu for the 2d sketch environment.
                        partSketchEnvironment.PanelBar.CommandBarList.Add(slotCommandBar);
                    }
                    //create the UI for ribbon interface
                    else
                    {
                        //get the ribbon associated with part document
                        Inventor.Ribbons ribbons;
                        ribbons = userInterfaceManager.Ribbons;

                        Inventor.Ribbon partRibbon;
                        partRibbon = ribbons["Part"];

                        //get the tabls associated with part ribbon
                        RibbonTabs ribbonTabs;
                        ribbonTabs = partRibbon.RibbonTabs;

                        RibbonTab partSketchRibbonTab;
                        partSketchRibbonTab = ribbonTabs["id_TabTools"]; //Sets which UI tab which contains the button.

                        //create a new panel with the tab
                        RibbonPanels ribbonPanels;
                        ribbonPanels = partSketchRibbonTab.RibbonPanels;

                        m_partSketchSlotRibbonPanel = ribbonPanels.Add("JIRA", "Autodesk:JirAddin:SlotRibbonPanel", "{DB59D9A7-EE4C-434A-BB5A-F93E8866E872}", "", false);

                        //add controls to the slot panel
                        CommandControls partSketchSlotRibbonPanelCtrls;
                        partSketchSlotRibbonPanelCtrls = m_partSketchSlotRibbonPanel.CommandControls;

                        //add the buttons to the ribbon panel
                        CommandControl drawSlotCmdBtnCmdCtrl;
                        drawSlotCmdBtnCmdCtrl = partSketchSlotRibbonPanelCtrls.AddButton(jiraButton.ButtonDefinition, true, true, "", false);


                    }
                }

            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }
        }

        public void Deactivate()
        {
            //the Deactivate method is called by Inventor when the AddIn is unloaded
            //the AddIn will be unloaded either manually by the user or
            //when the Inventor session is terminated

            try
            {
                m_userInterfaceEvents.OnResetCommandBars -= UserInterfaceEventsSink_OnResetCommandBarsEventDelegate;
                m_userInterfaceEvents.OnResetEnvironments -= UserInterfaceEventsSink_OnResetEnvironmentsEventDelegate;

                UserInterfaceEventsSink_OnResetCommandBarsEventDelegate = null;
                UserInterfaceEventsSink_OnResetEnvironmentsEventDelegate = null;
                m_userInterfaceEvents = null;
                if (m_partSketchSlotRibbonPanel != null)
                {
                    m_partSketchSlotRibbonPanel.Delete();
                }

                //release inventor Application object
                Marshal.ReleaseComObject(m_inventorApplication);
                m_inventorApplication = null;

                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }
        }

        public void ExecuteCommand(int CommandID)
        {
            //this method was used to notify when an AddIn command was executed
            //the CommandID parameter identifies the command that was executed

            //Note:this method is now obsolete, you should use the new
            //ControlDefinition objects to implement commands, they have
            //their own event sinks to notify when the command is executed
        }

        public object Automation
        {
            //if you want to return an interface to another client of this addin,
            //implement that interface in a class and return that class object 
            //through this property

            get
            {
                return null;
            }
        }

        private void UserInterfaceEvents_OnResetCommandBars(ObjectsEnumerator commandBars, NameValueMap context)
        {
            try
            {
                CommandBar commandBar;
                for (int commandBarCt = 1; commandBarCt <= commandBars.Count; commandBarCt++)
                {
                    commandBar = (Inventor.CommandBar)commandBars[commandBarCt];
                    if (commandBar.InternalName == "Autodesk:JiraAddIn:SlotToolbar")
                    {


                        //add button to toolbar
                        commandBar.Controls.AddButton(jiraButton.ButtonDefinition, 0);

                        return;
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }
        }

        private void UserInterfaceEvents_OnResetEnvironments(ObjectsEnumerator environments, NameValueMap context)
        {
            try
            {
                Inventor.Environment environment;
                for (int environmentCt = 1; environmentCt <= environments.Count; environmentCt++)
                {
                    environment = (Inventor.Environment)environments[environmentCt];
                    if (environment.InternalName == "PMxPartSketchEnvironment") //Change here if other view is needed?
                    {
                        //make this command bar accessible in the panel menu for the 2d sketch environment.
                        environment.PanelBar.CommandBarList.Add(m_inventorApplication.UserInterfaceManager.CommandBars["Autodesk:JiraAddIn:SlotToolbar"]);

                        return;
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }
        }

        private void UserInterfaceEvents_OnResetRibbonInterface(NameValueMap context)
        {
            try
            {

                UserInterfaceManager userInterfaceManager;
                userInterfaceManager = m_inventorApplication.UserInterfaceManager;

                //get the ribbon associated with part document
                Inventor.Ribbons ribbons;
                ribbons = userInterfaceManager.Ribbons;

                Inventor.Ribbon partRibbon;
                partRibbon = ribbons["Part"];

                //get the tabls associated with part ribbon
                RibbonTabs ribbonTabs;
                ribbonTabs = partRibbon.RibbonTabs;

                RibbonTab partSketchRibbonTab;
                partSketchRibbonTab = ribbonTabs["id_TabSketch"];

                //create a new panel with the tab
                RibbonPanels ribbonPanels;
                ribbonPanels = partSketchRibbonTab.RibbonPanels;

                m_partSketchSlotRibbonPanel = ribbonPanels.Add("JIRA", "Autodesk:JiraAddIn:SlotRibbonPanel",
                                                             "{DB59D9A7-EE4C-434A-BB5A-F93E8866E872}", "", false);

                //add controls to the slot panel
                CommandControls partSketchSlotRibbonPanelCtrls;
                partSketchSlotRibbonPanelCtrls = m_partSketchSlotRibbonPanel.CommandControls;



                //add the buttons to the ribbon panel
                CommandControl drawSlotCmdBtnCmdCtrl;
                drawSlotCmdBtnCmdCtrl = partSketchSlotRibbonPanelCtrls.AddButton(jiraButton.ButtonDefinition, false, true, "", false);

            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }
        }

        #endregion
    }
}
