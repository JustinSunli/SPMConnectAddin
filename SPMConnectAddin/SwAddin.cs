using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorks.Interop.swpublished;
using SolidWorksTools;
using SolidWorksTools.File;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SPMConnectAddin
{
    /// <summary>
    /// Summary description for SPMConnectAddin.
    /// </summary>[Guid("666CAF40-D1A8-42C5-AD90-ADE271FFC4BC")]
    [Guid("666CAF40-D1A8-42C5-AD90-ADE271FFC4BC"), ComVisible(true)]
    [SwAddin(
         Description = "SPMConnect addin for macros",
         Title = "SPMConnect",
         LoadAtStartup = true
         )]
    public class SwAddin : ISwAddin
    {
        #region Local Variables

        private ISldWorks iSwApp = null;
        private ICommandManager iCmdMgr = null;
        private int addinID = 0;
        private BitmapHandler iBmp;
        private ConnectAPI connectapi;

        public const int mainCmdGroupID = 5;
        public const int flyoutGroupID = 91;

        private string[] mainIcons = new string[6];
        private string[] icons = new string[6];

        #region Event Handler Variables

        private Hashtable openDocs = new Hashtable();
        private SolidWorks.Interop.sldworks.SldWorks SwEventPtr = null;

        #endregion Event Handler Variables

        // Public Properties
        public ISldWorks SwApp
        {
            get { return iSwApp; }
        }

        public ICommandManager CmdMgr
        {
            get { return iCmdMgr; }
        }

        public Hashtable OpenDocs
        {
            get { return openDocs; }
        }

        #endregion Local Variables

        #region SolidWorks Registration

        [ComRegisterFunctionAttribute]
        public static void RegisterFunction(Type t)
        {
            #region Get Custom Attribute: SwAddinAttribute

            SwAddinAttribute SWattr = null;
            Type type = typeof(SwAddin);

            foreach (System.Attribute attr in type.GetCustomAttributes(false))
            {
                if (attr is SwAddinAttribute)
                {
                    SWattr = attr as SwAddinAttribute;
                    break;
                }
            }

            #endregion Get Custom Attribute: SwAddinAttribute

            try
            {
                Microsoft.Win32.RegistryKey hklm = Microsoft.Win32.Registry.LocalMachine;
                Microsoft.Win32.RegistryKey hkcu = Microsoft.Win32.Registry.CurrentUser;

                string keyname = "SOFTWARE\\SolidWorks\\Addins\\{" + t.GUID.ToString() + "}";
                Microsoft.Win32.RegistryKey addinkey = hklm.CreateSubKey(keyname);
                addinkey.SetValue(null, 0);

                addinkey.SetValue("Description", SWattr.Description);
                addinkey.SetValue("Title", SWattr.Title);

                keyname = "Software\\SolidWorks\\AddInsStartup\\{" + t.GUID.ToString() + "}";
                addinkey = hkcu.CreateSubKey(keyname);
                addinkey.SetValue(null, Convert.ToInt32(SWattr.LoadAtStartup), Microsoft.Win32.RegistryValueKind.DWord);
            }
            catch (System.NullReferenceException nl)
            {
                Console.WriteLine("There was a problem registering this dll: SWattr is null. \n\"" + nl.Message + "\"");
                System.Windows.Forms.MessageBox.Show("There was a problem registering this dll: SWattr is null.\n\"" + nl.Message + "\"");
            }
            catch (System.Exception e)
            {
                Console.WriteLine(e.Message);

                System.Windows.Forms.MessageBox.Show("There was a problem registering the function: \n\"" + e.Message + "\"");
            }
        }

        [ComUnregisterFunctionAttribute]
        public static void UnregisterFunction(Type t)
        {
            try
            {
                Microsoft.Win32.RegistryKey hklm = Microsoft.Win32.Registry.LocalMachine;
                Microsoft.Win32.RegistryKey hkcu = Microsoft.Win32.Registry.CurrentUser;

                string keyname = "SOFTWARE\\SolidWorks\\Addins\\{" + t.GUID.ToString() + "}";
                hklm.DeleteSubKey(keyname);

                keyname = "Software\\SolidWorks\\AddInsStartup\\{" + t.GUID.ToString() + "}";
                hkcu.DeleteSubKey(keyname);
            }
            catch (System.NullReferenceException nl)
            {
                Console.WriteLine("There was a problem unregistering this dll: " + nl.Message);
                System.Windows.Forms.MessageBox.Show("There was a problem unregistering this dll: \n\"" + nl.Message + "\"");
            }
            catch (System.Exception e)
            {
                Console.WriteLine("There was a problem unregistering this dll: " + e.Message);
                System.Windows.Forms.MessageBox.Show("There was a problem unregistering this dll: \n\"" + e.Message + "\"");
            }
        }

        #endregion SolidWorks Registration

        #region ISwAddin Implementation

        public SwAddin()
        {
        }

        public bool ConnectToSW(object ThisSW, int cookie)
        {
            iSwApp = (ISldWorks)ThisSW;
            addinID = cookie;

            //Setup callbacks
            iSwApp.SetAddinCallbackInfo(0, this, addinID);

            #region Setup the Command Manager

            iCmdMgr = iSwApp.GetCommandManager(cookie);
            AddCommandMgr();

            #endregion Setup the Command Manager

            #region Setup the Event Handlers

            SwEventPtr = (SolidWorks.Interop.sldworks.SldWorks)iSwApp;
            openDocs = new Hashtable();
            AttachEventHandlers();

            #endregion Setup the Event Handlers

            connectapi = new ConnectAPI(iSwApp);
            return true;
        }

        public bool DisconnectFromSW()
        {
            RemoveCommandMgr();

            DetachEventHandlers();

            System.Runtime.InteropServices.Marshal.ReleaseComObject(iCmdMgr);
            iCmdMgr = null;
            System.Runtime.InteropServices.Marshal.ReleaseComObject(iSwApp);
            iSwApp = null;
            //The addin _must_ call GC.Collect() here in order to retrieve all managed code pointers
            GC.Collect();
            GC.WaitForPendingFinalizers();

            GC.Collect();
            GC.WaitForPendingFinalizers();

            return true;
        }

        #endregion ISwAddin Implementation

        #region UI Methods

        public void AddCommandMgr()
        {
            ICommandGroup cmdGroup;
            if (iBmp == null)
                iBmp = new BitmapHandler();
            Assembly thisAssembly;
            int[] cmdIndex = new int[26];
            string Title = "SPM Connect", ToolTip = "SPM Connect Addin";

            int[] docTypes = new int[]{(int)swDocumentTypes_e.swDocASSEMBLY,
                                       (int)swDocumentTypes_e.swDocDRAWING,
                                       (int)swDocumentTypes_e.swDocPART};

            thisAssembly = System.Reflection.Assembly.GetAssembly(this.GetType());

            int cmdGroupErr = 0;
            bool ignorePrevious = false;

            object registryIDs;
            //get the ID information stored in the registry
            bool getDataResult = iCmdMgr.GetGroupDataFromRegistry(mainCmdGroupID, out registryIDs);

            int[] knownIDs = new int[26] { 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25 };

            if (getDataResult)
            {
                if (!CompareIDs((int[])registryIDs, knownIDs)) //if the IDs don't match, reset the commandGroup
                {
                    ignorePrevious = true;
                }
            }

            int smallImage = 0;
            int mediumImage = 0;
            int largeImage = 0;
            int imageSizeToUse = 0;

            imageSizeToUse = iSwApp.GetImageSize(out smallImage, out mediumImage, out largeImage);

            cmdGroup = iCmdMgr.CreateCommandGroup2(mainCmdGroupID, Title, ToolTip, "", -1, ignorePrevious, ref cmdGroupErr);

            icons[0] = Path.Combine(GetAssemblyLocation(), @"icon20.png");
            icons[1] = Path.Combine(GetAssemblyLocation(), @"icon32.png");
            icons[2] = Path.Combine(GetAssemblyLocation(), @"icon40.png");
            icons[3] = Path.Combine(GetAssemblyLocation(), @"icon64.png");
            icons[4] = Path.Combine(GetAssemblyLocation(), @"icon96.png");
            icons[5] = Path.Combine(GetAssemblyLocation(), @"icon128.png");

            mainIcons[0] = Path.Combine(GetAssemblyLocation(), @"main20.png");
            mainIcons[1] = Path.Combine(GetAssemblyLocation(), @"main32.png");
            mainIcons[2] = Path.Combine(GetAssemblyLocation(), @"main40.png");
            mainIcons[3] = Path.Combine(GetAssemblyLocation(), @"main64.png");
            mainIcons[4] = Path.Combine(GetAssemblyLocation(), @"main96.png");
            mainIcons[5] = Path.Combine(GetAssemblyLocation(), @"main128.png");

            cmdGroup.IconList = icons;
            cmdGroup.MainIconList = mainIcons;

            int menuToolbarOption = (int)(swCommandItemType_e.swMenuItem | swCommandItemType_e.swToolbarItem);
            cmdIndex[0] = cmdGroup.AddCommandItem2("Save to Server", -1, "Save current model to spmconnect", "Save to Server", 20, "SavetoServer", "NeedEnabledPA", 0, menuToolbarOption);
            cmdIndex[1] = cmdGroup.AddCommandItem2("Open Model", -1, "Open Solidworks model", "Open Model", 11, "OpenModel", "", 1, menuToolbarOption);
            cmdIndex[2] = cmdGroup.AddCommandItem2("Open Read Only", -1, "Open Solidworks model in read only mode", "Open Read Only", 9, "OpenReadOnly", "", 2, menuToolbarOption);
            cmdIndex[3] = cmdGroup.AddCommandItem2("Copy Model", -1, "Save as copy of current model", "Copy Model", 2, "SaveasCopy", "NeedEnabledPA", 3, menuToolbarOption);
            cmdIndex[4] = cmdGroup.AddCommandItem2("Export Parasolid", -1, "Export current model as parasolid", "Export Parasolid", 4, "ExportasParasolid", "NeedEnabledPA", 4, menuToolbarOption);
            cmdIndex[5] = cmdGroup.AddCommandItem2("Export STEP", -1, "Export current model as STEP", "Export STEP", 5, "ExportasStep", "NeedEnabledPA", 5, menuToolbarOption);
            cmdIndex[6] = cmdGroup.AddCommandItem2("Export IGES", -1, "Export current model as IGES", "Export IGES", 6, "ExportasIges", "NeedEnabledPA", 6, menuToolbarOption);
            cmdIndex[7] = cmdGroup.AddCommandItem2("Export Dxf", -1, "Export current model as DXF", "Export Dxf", 8, "ExportasDXF", "NeedEnabledP", 7, menuToolbarOption);
            cmdIndex[8] = cmdGroup.AddCommandItem2("Export IGES CNC", -1, "Export current model to CNC folder as IGES", "Export IGES CNC", 12, "SaveIgesToCnc", "NeedEnabledPA", 8, menuToolbarOption);
            cmdIndex[9] = cmdGroup.AddCommandItem2("Delete Dangling", -1, "Delete dangling dimensions in drawing", "Delete Dangling", 7, "DelDangling", "NeedEnabledD", 9, menuToolbarOption);
            cmdIndex[10] = cmdGroup.AddCommandItem2("Reload Sheet Format", -1, "Reload current sheet format", "Reload Sheet Format", 14, "ReloadSheetformat", "NeedEnabledD", 10, menuToolbarOption);
            cmdIndex[11] = cmdGroup.AddCommandItem2("Dynamic Higlight", -1, "Turn on or off dynamic highlight", "Dynamic Highlight", 18, "DynamicHighlight", "NeedEnabledD", 11, menuToolbarOption);
            cmdIndex[12] = cmdGroup.AddCommandItem2("Export PDF", -1, "Export drawing as pdf", "Export PDF", 17, "Savedrawingaspdf", "NeedEnabledD", 12, menuToolbarOption);
            cmdIndex[13] = cmdGroup.AddCommandItem2("M6 FPT Sketch Block", -1, "Insert M6 FPT sketch block", "M6 FPT Sketch Block", 0, "Msixfptsketch", "NeedEnabledP", 13, menuToolbarOption);
            cmdIndex[14] = cmdGroup.AddCommandItem2("Close Inactive", -1, "Close all inactive documents in background", "Close Inactive", 1, "CloseInactive", "NeedEnabledPAD", 14, menuToolbarOption);
            cmdIndex[15] = cmdGroup.AddCommandItem2("Random Color", -1, "Assign random color to part", "Random Color", 16, "Randomcolor", "NeedEnabledP", 15, menuToolbarOption);
            cmdIndex[16] = cmdGroup.AddCommandItem2("Print Drawings", -1, "Print all drawings for assembly", "Print Drawings", 13, "Printall", "NeedEnabledPA", 16, menuToolbarOption);
            cmdIndex[17] = cmdGroup.AddCommandItem2("Export Parasolid CNC", -1, "Export current model to CNC folder as Parasolid", "Export Parasolid CNC", 19, "SaveParaToCnc", "NeedEnabledPA", 17, menuToolbarOption);
            cmdIndex[18] = cmdGroup.AddCommandItem2("Create Cube", -1, "Create a cube to start in a new part", "Create Cube", 21, "CreateCube", "NeedEnabledPA", 18, menuToolbarOption);
            cmdIndex[19] = cmdGroup.AddCommandItem2("SPM Connect", -1, "Show SPM Connect application", "SPM Connect", 22, "SPMConnect", "NeedEnabledPA", 19, menuToolbarOption);
            cmdIndex[20] = cmdGroup.AddCommandItem2("Add to Favorites", -1, "Add current item to favorites", "Add to Favorites", 23, "AddtoFav", "CheckNotFav", 20, menuToolbarOption);
            cmdIndex[21] = cmdGroup.AddCommandItem2("Remove from Favorites", -1, "Remove current item from favorites", "Remove from Favorites", 24, "RemoveFav", "CheckAllReadyFav", 21, menuToolbarOption);
            cmdIndex[22] = cmdGroup.AddCommandItem2("Mark Drawing Checked", -1, "Mark active part/assy drawings checked", "Mark Checked", 25, "MarkChecked", "CheckDrawingEnable", 22, menuToolbarOption);
            cmdIndex[23] = cmdGroup.AddCommandItem2("Mark Drawing Approved", -1, "Mark active part/assy drawings approved", "Mark Approved", 26, "MarkApproved", "ApproveDrawingEnable", 23, menuToolbarOption);
            cmdIndex[24] = cmdGroup.AddCommandItem2("Get My Next Item No", -1, "Get your next item no on your clipboard", "Next ItemNo", 27, "NextPartNo", "NeedEnabledPA", 24, menuToolbarOption);
            cmdIndex[25] = cmdGroup.AddCommandItem2("Check for updates", -1, "Check for new update on SPM Connect AddIn", "Check for updates", 28, "UpdateCheck", "NeedEnabledPA", 25, menuToolbarOption);

            cmdGroup.HasToolbar = true;
            cmdGroup.HasMenu = true;
            cmdGroup.Activate();

            FlyoutGroup flyGroup = iCmdMgr.CreateFlyoutGroup(flyoutGroupID, "Rapid Sketch", "Rapid Sketch", "Create Rapid Sketches",
          mainIcons[0], mainIcons[2], Path.Combine(GetAssemblyLocation(), @"icons_16.png"), Path.Combine(GetAssemblyLocation(), @"icons_24.png"), "FlyoutCallback", "FlyoutEnable");

            string t;
            t = locale.LProfileSketch;
            flyGroup.AddCommandItem(t, "", 0, "L_ProfileSketch", "FlyoutEnableCommandItem1");
            t = locale.UProfileSketch;
            flyGroup.AddCommandItem(t, "", 1, "U_ProfileSketch", "FlyoutEnableCommandItem1");
            t = locale.TProfileSketchClosed;
            flyGroup.AddCommandItem(t, "", 3, "T_ProfileClosedSketch", "FlyoutEnableCommandItem1");
            t = locale.HexagonSketch;
            flyGroup.AddCommandItem(t, "", 4, "HexagonSketch", "FlyoutEnableCommandItem1");
            t = locale.CircleSketch;
            flyGroup.AddCommandItem(t, "", 5, "CircleSketch", "FlyoutEnableCommandItem1");
            t = locale.RectangleSketch;
            flyGroup.AddCommandItem(t, "", 8, "RectangleSketch", "FlyoutEnableCommandItem1");
            t = locale.RectangleWithRevolveAxisSketch;
            flyGroup.AddCommandItem(t, "", 9, "RectangleWithRevolveAxisSketch", "FlyoutEnableCommandItem1");
            t = locale.CircleWithRevolveAxisSketch;
            flyGroup.AddCommandItem(t, "", 10, "CircleWithRevolveAxisSketch", "FlyoutEnableCommandItem1");

            flyGroup.FlyoutType = (int)swCommandFlyoutStyle_e.swCommandFlyoutStyle_Simple;

            foreach (int type in docTypes)
            {
                CommandTab cmdTab;

                cmdTab = iCmdMgr.GetCommandTab(type, Title);

                if (cmdTab != null & !getDataResult | ignorePrevious)//if tab exists, but we have ignored the registry info (or changed command group ID), re-create the tab.  Otherwise the ids won't matchup and the tab will be blank
                {
                    bool res = iCmdMgr.RemoveCommandTab(cmdTab);
                    cmdTab = null;
                }

                //if cmdTab is null, must be first load (possibly after reset), add the commands to the tabs
                if (cmdTab == null)
                {
                    cmdTab = iCmdMgr.AddCommandTab(type, Title);

                    CommandTabBox cmdBox = cmdTab.AddCommandTabBox();

                    int[] cmdIDs = new int[4];
                    int[] TextType = new int[4];

                    //Save
                    cmdIDs[0] = cmdGroup.get_CommandID(cmdIndex[1]);
                    TextType[0] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow;
                    //Open
                    cmdIDs[1] = cmdGroup.get_CommandID(cmdIndex[0]);
                    TextType[1] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow;
                    //Open Read only
                    cmdIDs[2] = cmdGroup.get_CommandID(cmdIndex[2]);
                    TextType[2] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow;
                    if (type == 1 || type == 2)
                    {
                        //Save as copy
                        cmdIDs[3] = cmdGroup.get_CommandID(cmdIndex[3]);
                        TextType[3] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow;
                    }

                    // Group 2 - Export
                    CommandTabBox cmdBox2 = cmdTab.AddCommandTabBox();
                    int[] cmdIDs2 = new int[6];
                    int[] TextType2 = new int[6];
                    if (type == 1 || type == 2)
                    {
                        cmdIDs2[0] = cmdGroup.get_CommandID(cmdIndex[4]);
                        TextType2[0] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal;
                        cmdIDs2[1] = cmdGroup.get_CommandID(cmdIndex[5]);
                        TextType2[1] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal;
                        cmdIDs2[2] = cmdGroup.get_CommandID(cmdIndex[6]);
                        TextType2[2] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal;
                        cmdIDs2[3] = cmdGroup.get_CommandID(cmdIndex[7]);
                        TextType2[3] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal;
                        cmdIDs2[4] = cmdGroup.get_CommandID(cmdIndex[8]);
                        TextType2[4] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal;
                        cmdIDs2[5] = cmdGroup.get_CommandID(cmdIndex[17]);
                        TextType2[5] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal;
                    }

                    // Group 3 Drawings
                    CommandTabBox cmdBox3 = cmdTab.AddCommandTabBox();
                    int[] cmdIDs3 = new int[4];
                    int[] TextType3 = new int[4];
                    if (type == 3)
                    {
                        cmdIDs3[0] = cmdGroup.get_CommandID(cmdIndex[9]);
                        TextType3[0] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow;
                        cmdIDs3[1] = cmdGroup.get_CommandID(cmdIndex[10]);
                        TextType3[1] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow;
                        cmdIDs3[2] = cmdGroup.get_CommandID(cmdIndex[11]);
                        TextType3[2] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow;
                        cmdIDs3[3] = cmdGroup.get_CommandID(cmdIndex[12]);
                        TextType3[3] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow;
                    }

                    // Group 5 Extras
                    CommandTabBox cmdBox5 = cmdTab.AddCommandTabBox();
                    int[] cmdIDs5 = new int[5];
                    int[] TextType5 = new int[5];
                    if (type == 1 || type == 2)
                    {
                        if (type == 1)
                        {
                            cmdIDs5[0] = cmdGroup.get_CommandID(cmdIndex[13]);
                            TextType5[0] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow;
                            cmdIDs5[3] = cmdGroup.get_CommandID(cmdIndex[15]);
                            TextType5[3] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow;
                            cmdIDs5[4] = cmdGroup.get_CommandID(cmdIndex[18]);
                            TextType5[4] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow;
                        }

                        cmdIDs5[1] = cmdGroup.get_CommandID(cmdIndex[14]);
                        TextType5[1] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow;
                        cmdIDs5[2] = cmdGroup.get_CommandID(cmdIndex[16]);
                        TextType5[2] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow;
                    }

                    // Group 4
                    CommandTabBox cmdBox4 = cmdTab.AddCommandTabBox();
                    int[] cmdIDs4 = new int[8];
                    int[] TextType4 = new int[8];

                    if (type == 1 || type == 2)
                    {
                        if (type == 1)
                        {
                            cmdIDs4[0] = flyGroup.CmdID;
                            TextType4[0] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow | (int)swCommandTabButtonFlyoutStyle_e.swCommandTabButton_ActionFlyout;
                        }
                        cmdIDs4[1] = cmdGroup.get_CommandID(cmdIndex[19]);
                        TextType4[1] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow;
                        cmdIDs4[2] = cmdGroup.get_CommandID(cmdIndex[20]);
                        TextType4[2] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow;
                        cmdIDs4[3] = cmdGroup.get_CommandID(cmdIndex[21]);
                        TextType4[3] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow;
                        cmdIDs4[4] = cmdGroup.get_CommandID(cmdIndex[22]);
                        TextType4[4] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow;
                        cmdIDs4[5] = cmdGroup.get_CommandID(cmdIndex[23]);
                        TextType4[5] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow;
                        cmdIDs4[6] = cmdGroup.get_CommandID(cmdIndex[24]);
                        TextType4[6] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow;
                        cmdIDs4[7] = cmdGroup.get_CommandID(cmdIndex[25]);
                        TextType4[7] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow;
                    }

                    // Add the commands
                    cmdBox.AddCommands(cmdIDs, TextType);
                    if (type == 1 || type == 2)
                    {
                        cmdBox2.AddCommands(cmdIDs2, TextType2);
                        cmdBox4.AddCommands(cmdIDs4, TextType4);
                        cmdBox5.AddCommands(cmdIDs5, TextType5);
                    }
                    if (type == 3)
                        cmdBox3.AddCommands(cmdIDs3, TextType3);

                    // Add separators
                    if (type == 1 || type == 2)
                    {
                        cmdTab.AddSeparator(cmdBox2, cmdIDs2[0]);
                        cmdTab.AddSeparator(cmdBox5, cmdIDs5[0]);
                        cmdTab.AddSeparator(cmdBox4, cmdIDs4[0]);
                    }

                    if (type == 3)
                        cmdTab.AddSeparator(cmdBox3, cmdIDs3[0]);
                }
            }

            thisAssembly = null;
        }

        private string GetAssemblyLocation()
        {
            return System.Reflection.Assembly.GetExecutingAssembly().Location.Remove(System.Reflection.Assembly.GetExecutingAssembly().Location.LastIndexOf(@"\"));
        }

        public void RemoveCommandMgr()
        {
            iBmp.Dispose();
            iCmdMgr.RemoveCommandGroup(mainCmdGroupID);
        }

        public bool CompareIDs(int[] storedIDs, int[] addinIDs)
        {
            List<int> storedList = new List<int>(storedIDs);
            List<int> addinList = new List<int>(addinIDs);

            addinList.Sort();
            storedList.Sort();

            if (addinList.Count != storedList.Count)
            {
                return false;
            }
            else
            {
                for (int i = 0; i < addinList.Count; i++)
                {
                    if (addinList[i] != storedList[i])
                    {
                        return false;
                    }
                }
            }
            return true;
        }

        #endregion UI Methods

        #region UI Callbacks

        public void CreateCube() => connectapi.CreateCube();

        public void FlyoutCommandItem1()
        {
            iSwApp.SendMsgToUser("Flyout command 1");
        }

        public void SavetoServer()
        {
            connectapi.SaveToServer(true);
        }

        public void OpenModel()
        {
            SPMConnectAddin.OpenItem openItem = new SPMConnectAddin.OpenItem();
            openItem.BringToFront();
            openItem.TopMost = true;
            openItem.Focus();
            string input = "";
            if (openItem.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                input = openItem.ValueIWant;
                if (input.Length == 6 && !String.IsNullOrEmpty(input) && Char.IsLetter(input[0]))
                {
                    connectapi.Checkforspmfile(input, false);
                }
                else
                {
                    if (!(input == ""))
                        MessageBox.Show("Not a valid part number. Please enter a valid six digit SPM item number (starting with 'A', 'B', 'C') to open solidworks model.", "SPM Connect open model", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        public void OpenReadOnly()
        {
            SPMConnectAddin.OpenItem openItem = new SPMConnectAddin.OpenItem();
            openItem.BringToFront();
            openItem.TopMost = true;
            openItem.Focus();
            string input = "";
            if (openItem.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                input = openItem.ValueIWant;
                if (input.Length == 6 && !String.IsNullOrEmpty(input) && Char.IsLetter(input[0]))
                {
                    connectapi.Checkforspmfile(input, true);
                }
                else
                {
                    if (!(input == ""))
                        MessageBox.Show("Not a valid part number. Please enter a valid six digit SPM item number (starting with 'A', 'B', 'C') to open solidworks model.", "SPM Connect open model", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        public void SaveasCopy()
        {
            ModelDoc2 swModel = iSwApp.ActiveDoc as ModelDoc2;
            if (swModel == null)
            {
                MessageBox.Show("No active model found", "SPM Connect", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }

            if (swModel.GetType() != (int)swDocumentTypes_e.swDocPART && swModel.GetType() != (int)swDocumentTypes_e.swDocASSEMBLY)
            {
                // Tell user
                MessageBox.Show("Active model is not a part or assembly to copy", "SPM Connect", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            DialogResult result = MessageBox.Show("Are you sure want to copy item no. " + connectapi.Getfilename() + " to a new item?", "SPM Connect - Copy Item?", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                string user = connectapi.UserName();
                string activeblock = connectapi.Getactiveblock();
                string lastnumber = connectapi.Getlastnumber();
                if (activeblock.ToString().Length > 0)
                {
                    if (connectapi.Validnumber(lastnumber.ToString()) == true)
                    {
                        connectapi.Prepareforcopy(activeblock.ToString(), connectapi.Getfilename(), lastnumber);
                    }
                }
                else
                {
                    MessageBox.Show("User block number has not been assigned. Please contact the admin.", "SPM Connect", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        public void ExportasStep() => connectapi.ExportModelAsStep();

        public void ExportasDXF() => connectapi.ExportPartAsDxf();

        public void ExportasParasolid() => connectapi.ExportModelAsParasolid();

        public void ExportasIges() => connectapi.ExportModelAsIGES();

        public void SaveIgesToCnc() => connectapi.ExportModelAsIGESToCNC();

        public void SaveParaToCnc() => connectapi.ExportModelAsParasolidToCNC();

        public void DelDangling() => connectapi.DeleteDanglingDimensions();

        public void ReloadSheetformat() => connectapi.ReloadSheetformat();

        public void DynamicHighlight()
        {
            iSwApp.SetUserPreferenceToggle(119, true);
            MessageBox.Show("Dynamic highlight turned on.", "SPM Connect", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public void Savedrawingaspdf() => connectapi.ExportDrawingAsPdf();

        public void Msixfptsketch() => connectapi.InsertM6FPT();

        public void CloseInactive() => connectapi.CloseInactive();

        public void Randomcolor() => connectapi.Randomcolor();

        private Form1 progressBarForm = new Form1();

        public void UpdateCheck()
        {
            SPMConnectAddin.UpdateChecker.start(@"\\spm-adfs\SDBASE\SPM Connect Addin\update.xml", "SPM Connect AddIn", "spm", false);
        }

        public void Printall()
        {
            //iSwApp.SendMsgToUser("Coming soon!")
            Task task = new Task(PrintDrawings);
            task.Start();
            progressBarForm.ShowDialog();
            progressBarForm.UpdateProgressBar(1, "Fetching components");
        }

        private void PrintDrawings()
        {
            connectapi.PrintDrawings(progressBarForm);
        }

        [DllImport("USER32.DLL")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImportAttribute("User32.DLL")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        public void SPMConnect()
        {
            IntPtr handle;
            Process[] processName = Process.GetProcessesByName("SPM Connect");
            if (processName.Length == 0)
            {
                //Start application here
                Process.Start(@"\\spm-adfs\SDBASE\SPM Connect SQL\SPM Connect.application");
            }
            else
            {
                //Set foreground window
                handle = processName[0].MainWindowHandle;
                ShowWindow(handle, 9);
                SetForegroundWindow(handle);
            }
        }

        public void AddtoFav()
        {
            connectapi.Addtofavorites("");
        }

        public int CheckNotFav()
        {
            ModelDoc2 swModel = iSwApp.ActiveDoc as ModelDoc2;
            if (swModel == null)
            {
                return 0;
            }
            else if (String.IsNullOrEmpty(swModel.GetPathName()))
            {
                return 0;
            }
            if (swModel.GetType() == (int)swDocumentTypes_e.swDocASSEMBLY || swModel.GetType() == (int)swDocumentTypes_e.swDocPART)
            {
                bool result = connectapi.CheckitempresentonFavorites("");

                return result ? 0 : 1;
            }
            return 0;
        }

        public void RemoveFav()
        {
            connectapi.Removefromfavorites("");
        }

        public int CheckAllReadyFav()
        {
            ModelDoc2 swModel = iSwApp.ActiveDoc as ModelDoc2;
            if (swModel == null)
            {
                return 0;
            }
            else if (String.IsNullOrEmpty(swModel.GetPathName()))
            {
                return 0;
            }
            if (swModel.GetType() == (int)swDocumentTypes_e.swDocASSEMBLY || swModel.GetType() == (int)swDocumentTypes_e.swDocPART)
            {
                bool result = connectapi.CheckitempresentonFavorites("");
                return result ? 1 : 0;
            }
            return 0;
        }

        public int NeedEnabledPA()
        {
            ModelDoc2 swModel = iSwApp.ActiveDoc as ModelDoc2;
            if (swModel == null)
                return 0;

            if (swModel.GetType() == (int)swDocumentTypes_e.swDocASSEMBLY || swModel.GetType() == (int)swDocumentTypes_e.swDocPART)
            {
                return 1;
            }
            return 0;
        }

        public int NeedEnabledPAD()
        {
            ModelDoc2 swModel = iSwApp.ActiveDoc as ModelDoc2;
            if (swModel == null)
                return 0;

            if (swModel.GetType() == (int)swDocumentTypes_e.swDocASSEMBLY || swModel.GetType() == (int)swDocumentTypes_e.swDocPART || swModel.GetType() == (int)swDocumentTypes_e.swDocDRAWING)
            {
                return 1;
            }
            return 0;
        }

        public int NeedEnabledD()
        {
            ModelDoc2 swModel = iSwApp.ActiveDoc as ModelDoc2;
            if (swModel == null)
                return 0;

            if (swModel.GetType() == (int)swDocumentTypes_e.swDocDRAWING)
            {
                return 1;
            }
            return 0;
        }

        public int NeedEnabledP()
        {
            ModelDoc2 swModel = iSwApp.ActiveDoc as ModelDoc2;
            if (swModel == null)
                return 0;

            if (swModel.GetType() == (int)swDocumentTypes_e.swDocPART)
            {
                return 1;
            }
            return 0;
        }

        public int CheckDrawingEnable()
        {
            ModelDoc2 swModel = iSwApp.ActiveDoc as ModelDoc2;
            if (swModel == null)
            {
                return 0;
            }
            else if (String.IsNullOrEmpty(swModel.GetPathName()))
            {
                return 0;
            }
            if (swModel.GetType() == (int)swDocumentTypes_e.swDocASSEMBLY || swModel.GetType() == (int)swDocumentTypes_e.swDocPART)
            {
                if (connectapi.CheckingDrawingRights())
                {
                    return 1;
                }
                else
                {
                    return 0;
                }

                // TODO: do learning on the drawing approval table to see the drawing has been marked checked or not
            }
            return 0;
        }

        public int ApproveDrawingEnable()
        {
            ModelDoc2 swModel = iSwApp.ActiveDoc as ModelDoc2;
            if (swModel == null)
            {
                return 0;
            }
            else if (String.IsNullOrEmpty(swModel.GetPathName()))
            {
                return 0;
            }
            if (swModel.GetType() == (int)swDocumentTypes_e.swDocASSEMBLY || swModel.GetType() == (int)swDocumentTypes_e.swDocPART)
            {
                if (connectapi.ApproveDrawingRights())
                {
                    bool result = connectapi.CheckMarkedDrawingExists("");
                    return result ? 1 : 0;
                }
                else
                {
                    return 0;
                }
            }
            return 0;
        }

        public void MarkChecked()
        {
            //iSwApp.SendMsgToUser("Coming soon!")
            Task task = new Task(MarkDrawingChecked);
            task.Start();
            progressBarForm.ShowDialog();
            progressBarForm.UpdateProgressBar(1, "Fetching components");
        }

        private void MarkDrawingChecked()
        {
            connectapi.MarkDrawings(progressBarForm, false);
        }

        public void MarkApproved()
        {
            //iSwApp.SendMsgToUser("Coming soon!")
            Task task = new Task(MarkDrawApproved);
            task.Start();
            progressBarForm.ShowDialog();
            progressBarForm.UpdateProgressBar(1, "Fetching components");
        }

        private void MarkDrawApproved()
        {
            connectapi.MarkDrawings(progressBarForm, true);
        }

        public void NextPartNo()
        {
            connectapi.GetMyNextPartNo();
        }

        #region Sketch CallBacks

        public void L_ProfileSketch()
        {
            ModelDoc2 swDoc = (ModelDoc2)iSwApp.ActiveDoc;
            if (swDoc.GetType() == 1)
            {
                iSwApp.SendMsgToUser2("Active model is not a part", 0, 2);
            }
            // Create sketch
            swDoc.SketchManager.InsertSketch(false);
            // Draw the lines
            SketchSegment line1, line2;
            line1 = (SketchSegment)(swDoc.SketchManager.CreateLine(0.0, 0.0, 0.0, 0.05, 0.0, 0.0)); // horizontal
            line2 = (SketchSegment)(swDoc.SketchManager.CreateLine(0.0, 0.0, 0.0, 0.0, 0.05, 0.0)); // vertical
                                                                                                    // Add dimensions
            line1.Select4(false, null);
            swDoc.AddDimension2(0.01, -0.01, 0.0);
            line2.Select4(false, null);
            swDoc.AddDimension2(-0.01, 0.01, 0.0);
            // Exit sketch
            swDoc.ClearSelection2(true);
            swDoc.SketchManager.InsertSketch(true);
            swDoc.ViewZoomtofit2();
        }

        public void U_ProfileSketch()
        {
            ModelDoc2 swDoc = (ModelDoc2)iSwApp.ActiveDoc;
            // Create sketch
            swDoc.SketchManager.InsertSketch(false);
            // Draw the lines
            SketchSegment line1, line2, line3;
            SketchPoint origin;
            line1 = (SketchSegment)(swDoc.SketchManager.CreateLine(-0.05, 0.0, 0.0, 0.05, 0.0, 0.0)); // horizontal
            line2 = (SketchSegment)(swDoc.SketchManager.CreateLine(-0.05, 0.0, 0.0, -0.05, 0.05, 0.0)); // left
            line3 = (SketchSegment)(swDoc.SketchManager.CreateLine(0.05, 0.0, 0.0, 0.05, 0.05, 0.0)); // right
                                                                                                      // left and right line same length
            line2.Select4(false, null);
            line3.Select4(true, null);
            swDoc.SketchAddConstraints("sgSAMELENGTH");
            // line 1 center on datum origin
            //Feature f = (Feature)swDoc.FirstFeature();
            //while (f != null)
            //{
            //    if (f.GetTypeName2() == "OriginProfileFeature")
            //    {
            //        f.Select2(false, 0);
            //        break;
            //    }
            //    f = (Feature)f.GetNextFeature();
            //}
            //swDoc.Extension.SelectByID2("", swSelectType_e.swSelDATUMPOINTS.ToString(), 0.0, 0.0, 0.0, false, 0, null, 0);
            origin = swDoc.SketchManager.CreatePoint(0.0, 0.0, 0.0);
            origin.Select4(false, null);
            swDoc.SketchAddConstraints("sgFIXED");
            line1.Select4(true, null);
            swDoc.SketchAddConstraints("sgATMIDDLE");
            // Add dimensions
            line1.Select4(false, null);
            swDoc.AddDimension2(0.0, -0.05, 0.0);
            line2.Select4(false, null);
            swDoc.AddDimension2(-0.06, 0.025, 0.0);
            // Exit sketch
            swDoc.ClearSelection2(true);
            swDoc.SketchManager.InsertSketch(true);
            swDoc.ViewZoomtofit2();
        }

        public void U_ProfileFlangeSketch()
        {
            ModelDoc2 swDoc = (ModelDoc2)iSwApp.ActiveDoc;
            // Create sketch
            swDoc.SketchManager.InsertSketch(false);
            // Draw the lines
            SketchSegment line1, line2, line3, line4, line5;
            SketchPoint origin;
            line1 = (SketchSegment)(swDoc.SketchManager.CreateLine(-0.05, 0.0, 0.0, 0.05, 0.0, 0.0)); // horizontal
            line2 = (SketchSegment)(swDoc.SketchManager.CreateLine(-0.05, 0.0, 0.0, -0.05, 0.05, 0.0)); // left vertical
            line3 = (SketchSegment)(swDoc.SketchManager.CreateLine(0.05, 0.0, 0.0, 0.05, 0.05, 0.0)); // right vertical
            line4 = (SketchSegment)(swDoc.SketchManager.CreateLine(-0.1, 0.05, 0.0, -0.05, 0.05, 0.0)); // left horizontal
            line5 = (SketchSegment)(swDoc.SketchManager.CreateLine(0.05, 0.05, 0.0, 0.1, 0.05, 0.0)); // right horizontal
                                                                                                      // left and right lines same length
            line2.Select4(false, null);
            line3.Select4(true, null);
            swDoc.SketchAddConstraints("sgSAMELENGTH");
            line4.Select4(false, null);
            line5.Select4(true, null);
            swDoc.SketchAddConstraints("sgSAMELENGTH");
            // line 1 center on datum origin
            origin = swDoc.SketchManager.CreatePoint(0.0, 0.0, 0.0);
            origin.Select4(false, null);
            swDoc.SketchAddConstraints("sgFIXED");
            line1.Select4(true, null);
            swDoc.SketchAddConstraints("sgATMIDDLE");
            // Add dimensions
            line1.Select4(false, null);
            swDoc.AddDimension2(0.0, -0.05, 0.0);
            line2.Select4(false, null);
            swDoc.AddDimension2(-0.06, 0.025, 0.0);
            line4.Select4(false, null);
            swDoc.AddDimension2(-0.09, 0.025, 0.0);
            // Exit sketch
            swDoc.ClearSelection2(true);
            swDoc.SketchManager.InsertSketch(true);
            swDoc.ViewZoomtofit2();
        }

        public void T_ProfileClosedSketch()
        {
            ModelDoc2 swDoc = (ModelDoc2)iSwApp.ActiveDoc;
            // Create sketch
            swDoc.SketchManager.InsertSketch(false);
            // Draw the lines
            SketchSegment line1, line2, line3, line4, line5, line6, line7, line8, axis;
            SketchPoint origin;
            // Add the origin point
            origin = swDoc.SketchManager.CreatePoint(0.0, 0.0, 0.0);
            origin.Select4(false, null);
            swDoc.SketchAddConstraints("sgFIXED");
            // now the lines
            line1 = swDoc.SketchManager.CreateLine(-0.05, 0.0, 0.0, 0.05, 0.0, 0.0);
            line2 = swDoc.SketchManager.CreateLine(0.05, 0.0, 0.0, 0.05, 0.01, 0.0);
            line3 = swDoc.SketchManager.CreateLine(0.05, 0.01, 0.0, 0.01, 0.01, 0.0);
            line4 = swDoc.SketchManager.CreateLine(0.01, 0.01, 0.0, 0.01, 0.1, 0.0);
            line5 = swDoc.SketchManager.CreateLine(0.01, 0.1, 0.0, -0.01, 0.1, 0.0);
            line6 = swDoc.SketchManager.CreateLine(-0.01, 0.1, 0.0, -0.01, 0.01, 0.0);
            line7 = swDoc.SketchManager.CreateLine(-0.01, 0.01, 0.0, -0.05, 0.01, 0.0);
            line8 = swDoc.SketchManager.CreateLine(-0.05, 0.01, 0.0, -0.05, 0.0, 0.0);
            axis = swDoc.SketchManager.CreateCenterLine(0.0, 0.0, 0.0, 0.0, 0.1, 0.0);
            // add constraints
            axis.Select4(false, null);
            line4.Select4(true, null);
            line6.Select4(true, null);
            swDoc.SketchAddConstraints("sgSYMMETRIC");
            swDoc.ClearSelection2(true);
            line3.Select4(false, null);
            line7.Select4(true, null);
            swDoc.SketchAddConstraints("sgSAMELENGTH");
            swDoc.SketchAddConstraints("sgCOLINEAR");
            // Add dimensions
            swDoc.ClearSelection2(true);
            line1.Select4(false, null);
            swDoc.AddDimension2(0.0, -0.01, 0.0);
            line8.Select4(false, null);
            swDoc.AddDimension2(-0.051, 0.005, 0.0);
            line1.Select4(false, null);
            line5.Select4(true, null);
            swDoc.AddDimension2(0.055, 0.05, 0.0);
            swDoc.ClearSelection2(true);
            line5.Select4(false, null);
            swDoc.AddDimension2(0.0, 0.11, 0.0);
            // Exit sketch
            swDoc.ClearSelection2(true);
            swDoc.SketchManager.InsertSketch(true);
            swDoc.ViewZoomtofit2();
        }

        public void HexagonSketch()
        {
            ModelDoc2 swDoc = (ModelDoc2)iSwApp.ActiveDoc;
            // Create sketch
            swDoc.SketchManager.InsertSketch(false);
            object[] hexagon;
            hexagon = (object[])swDoc.SketchManager.CreatePolygon(0.0, 0.0, 0.0, 0.0, 0.05, 0.0, 6, true);
            swDoc.ClearSelection2(true);
            foreach (object x in hexagon)
            {
                // select one line and make it horizontal
                SketchSegment y = (SketchSegment)x;
                if (y.GetType() == (int)swSketchSegments_e.swSketchLINE)
                {
                    y.Select4(false, null);
                    swDoc.SketchAddConstraints("sgHORIZONTAL2D");
                    break;
                }
            }
            foreach (object x in hexagon)
            {
                // select the circle and add dimension to it
                SketchSegment y = (SketchSegment)x;
                if (y.GetType() == (int)swSketchSegments_e.swSketchARC)
                {
                    y.Select4(false, null);
                    swDoc.AddDimension2(-0.075, 0.075, 0.0);
                    break;
                }
            }
            swDoc.SketchManager.InsertSketch(true);
            swDoc.ViewZoomtofit2();
        }

        public void CircleSketch()
        {
            ModelDoc2 swDoc = (ModelDoc2)iSwApp.ActiveDoc;
            // Create sketch
            swDoc.SketchManager.InsertSketch(false);
            SketchSegment circle;
            circle = (SketchSegment)swDoc.SketchManager.CreateCircle(0.0, 0.0, 0.0, 0.05, 0.0, 0.0);
            circle.Select4(false, null);
            swDoc.AddDimension2(0.075, 0.075, 0.0);
            swDoc.ClearSelection2(true);
            swDoc.SketchManager.InsertSketch(true);
            swDoc.ViewZoomtofit2();
        }

        public void RectangleSketch()
        {
            ModelDoc2 swDoc = (ModelDoc2)iSwApp.ActiveDoc;
            // Create sketch
            swDoc.SketchManager.InsertSketch(false);
            object[] rectangle;
            rectangle = (object[])swDoc.SketchManager.CreateCenterRectangle(0.0, 0.0, 0.0, 0.05, 0.025, 0.0);
            SketchSegment l1, l2;
            l1 = (SketchSegment)rectangle[0];
            l2 = (SketchSegment)rectangle[1];
            l1.Select4(false, null);
            swDoc.AddDimension2(0.0, 0.075, 0.0);
            l2.Select4(false, null);
            swDoc.AddDimension2(0.075, 0.0125, 0.0);
            swDoc.SketchManager.InsertSketch(true);
            swDoc.ViewZoomtofit2();
        }

        public void RectangleWithRevolveAxisSketch()
        {
            ModelDoc2 swDoc = (ModelDoc2)iSwApp.ActiveDoc;
            // Create sketch
            swDoc.SketchManager.InsertSketch(false);
            SketchSegment axis, l0, l1, l2, l3;
            object[] rectangle;
            SketchPoint origin;
            // Add the origin point
            origin = swDoc.SketchManager.CreatePoint(0.0, 0.0, 0.0);
            origin.Select4(false, null);
            swDoc.SketchAddConstraints("sgFIXED");
            axis = swDoc.SketchManager.CreateLine(0.0, 0.0, 0.0, 0.0, 0.05, 0.0);
            axis.ConstructionGeometry = true;
            rectangle = (object[])swDoc.SketchManager.CreateCornerRectangle(0.05, 0.0, 0.0, 0.075, 0.05, 0.0);
            l0 = (SketchSegment)rectangle[0];
            l1 = (SketchSegment)rectangle[1];
            l2 = (SketchSegment)rectangle[2];
            l3 = (SketchSegment)rectangle[3];
            axis.Select4(false, null);
            l1.Select4(true, null);
            swDoc.AddDimension2(0.025, -0.05, 0.0);
            swDoc.ClearSelection2(true);
            l2.Select4(false, null);
            swDoc.AddDimension2(0.0625, 0.075, 0.0);
            swDoc.ClearSelection2(true);
            l3.Select4(false, null);
            swDoc.AddDimension2(0.1, 0.025, 0.0);
            swDoc.ClearSelection2(true);
            l0.Select4(false, null);
            origin.Select4(true, null);
            swDoc.SketchAddConstraints("sgCOINCIDENT");
            swDoc.ClearSelection2(true);
            swDoc.SketchManager.InsertSketch(true);
            swDoc.ViewZoomtofit2();
        }

        public void CircleWithRevolveAxisSketch()
        {
            ModelDoc2 swDoc = (ModelDoc2)iSwApp.ActiveDoc;
            // Create sketch
            swDoc.SketchManager.InsertSketch(false);
            SketchSegment axis, circle;
            SketchPoint origin, c_origin;
            // Add the origin point
            origin = swDoc.SketchManager.CreatePoint(0.0, 0.0, 0.0);
            origin.Select4(false, null);
            swDoc.SketchAddConstraints("sgFIXED");
            swDoc.ClearSelection2(true);
            axis = swDoc.SketchManager.CreateLine(0.0, 0.0, 0.0, 0.0, 0.05, 0.0);
            axis.ConstructionGeometry = true;
            c_origin = swDoc.SketchManager.CreatePoint(0.05, 0.0, 0.0);
            c_origin.Select4(false, null);
            origin.Select4(true, null);
            swDoc.SketchAddConstraints("sgHORIZONTALPOINTS2D");
            circle = swDoc.SketchManager.CreateCircleByRadius(0.05, 0.0, 0.0, 0.025);
            axis.Select4(false, null);
            circle.Select4(true, null);
            swDoc.ClearSelection2(true);
            circle.Select4(false, null);
            origin.Select4(true, null);
            swDoc.AddDimension2(0.025, 0.05, 0.0);
            swDoc.ClearSelection2(true);
            circle.Select4(false, null);
            swDoc.AddDimension2(0.075, 0.05, 0.0);
            swDoc.ClearSelection2(true);
            swDoc.SketchManager.InsertSketch(true);
            swDoc.ViewZoomtofit2();
        }

        #endregion Sketch CallBacks

        public void FlyoutCallback()
        {
            FlyoutGroup flyGroup = iCmdMgr.GetFlyoutGroup(flyoutGroupID);
            flyGroup.RemoveAllCommandItems();
            string t;
            t = locale.LProfileSketch;
            flyGroup.AddCommandItem(t, "", 0, "L_ProfileSketch", "FlyoutEnableCommandItem1");
            t = locale.UProfileSketch;
            flyGroup.AddCommandItem(t, "", 1, "U_ProfileSketch", "FlyoutEnableCommandItem1");
            t = locale.TProfileSketchClosed;
            flyGroup.AddCommandItem(t, "", 3, "T_ProfileClosedSketch", "FlyoutEnableCommandItem1");
            t = locale.HexagonSketch;
            flyGroup.AddCommandItem(t, "", 4, "HexagonSketch", "FlyoutEnableCommandItem1");
            t = locale.CircleSketch;
            flyGroup.AddCommandItem(t, "", 5, "CircleSketch", "FlyoutEnableCommandItem1");
            t = locale.RectangleSketch;
            flyGroup.AddCommandItem(t, "", 8, "RectangleSketch", "FlyoutEnableCommandItem1");
            t = locale.RectangleWithRevolveAxisSketch;
            flyGroup.AddCommandItem(t, "", 9, "RectangleWithRevolveAxisSketch", "FlyoutEnableCommandItem1");
            t = locale.CircleWithRevolveAxisSketch;
            flyGroup.AddCommandItem(t, "", 10, "CircleWithRevolveAxisSketch", "FlyoutEnableCommandItem1");
        }

        public int FlyoutEnable()
        {
            return 1;
        }

        public int FlyoutEnableCommandItem1()
        {
            return 1;
        }

        #endregion UI Callbacks

        #region Event Methods

        public bool AttachEventHandlers()
        {
            AttachSwEvents();
            //Listen for events on all currently open docs
            AttachEventsToAllDocuments();
            return true;
        }

        private bool AttachSwEvents()
        {
            try
            {
                SwEventPtr.ActiveDocChangeNotify += new DSldWorksEvents_ActiveDocChangeNotifyEventHandler(OnDocChange);
                SwEventPtr.DocumentLoadNotify2 += new DSldWorksEvents_DocumentLoadNotify2EventHandler(OnDocLoad);
                SwEventPtr.FileNewNotify2 += new DSldWorksEvents_FileNewNotify2EventHandler(OnFileNew);
                SwEventPtr.ActiveModelDocChangeNotify += new DSldWorksEvents_ActiveModelDocChangeNotifyEventHandler(OnModelChange);
                SwEventPtr.FileOpenPostNotify += new DSldWorksEvents_FileOpenPostNotifyEventHandler(FileOpenPostNotify);
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return false;
            }
        }

        private bool DetachSwEvents()
        {
            try
            {
                SwEventPtr.ActiveDocChangeNotify -= new DSldWorksEvents_ActiveDocChangeNotifyEventHandler(OnDocChange);
                SwEventPtr.DocumentLoadNotify2 -= new DSldWorksEvents_DocumentLoadNotify2EventHandler(OnDocLoad);
                SwEventPtr.FileNewNotify2 -= new DSldWorksEvents_FileNewNotify2EventHandler(OnFileNew);
                SwEventPtr.ActiveModelDocChangeNotify -= new DSldWorksEvents_ActiveModelDocChangeNotifyEventHandler(OnModelChange);
                SwEventPtr.FileOpenPostNotify -= new DSldWorksEvents_FileOpenPostNotifyEventHandler(FileOpenPostNotify);
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return false;
            }
        }

        public void AttachEventsToAllDocuments()
        {
            ModelDoc2 modDoc = (ModelDoc2)iSwApp.GetFirstDocument();
            while (modDoc != null)
            {
                if (!openDocs.Contains(modDoc))
                {
                    AttachModelDocEventHandler(modDoc);
                }
                modDoc = (ModelDoc2)modDoc.GetNext();
            }
        }

        public bool AttachModelDocEventHandler(ModelDoc2 modDoc)
        {
            if (modDoc == null)
                return false;

            DocumentEventHandler docHandler = null;

            if (!openDocs.Contains(modDoc))
            {
                switch (modDoc.GetType())
                {
                    case (int)swDocumentTypes_e.swDocPART:
                        {
                            docHandler = new PartEventHandler(modDoc, this, connectapi);

                            break;
                        }
                    case (int)swDocumentTypes_e.swDocASSEMBLY:
                        {
                            docHandler = new AssemblyEventHandler(modDoc, this, connectapi);
                            break;
                        }
                    case (int)swDocumentTypes_e.swDocDRAWING:
                        {
                            docHandler = new DrawingEventHandler(modDoc, this, connectapi);
                            break;
                        }
                    default:
                        {
                            return false; //Unsupported document type
                        }
                }
                docHandler.AttachEventHandlers();
                openDocs.Add(modDoc, docHandler);
            }
            return true;
        }

        public bool DetachModelEventHandler(ModelDoc2 modDoc)
        {
            DocumentEventHandler docHandler;
            docHandler = (DocumentEventHandler)openDocs[modDoc];
            openDocs.Remove(modDoc);
            modDoc = null;
            docHandler = null;
            return true;
        }

        public bool DetachEventHandlers()
        {
            DetachSwEvents();

            //Close events on all currently open docs
            DocumentEventHandler docHandler;
            int numKeys = openDocs.Count;
            object[] keys = new Object[numKeys];

            //Remove all document event handlers
            openDocs.Keys.CopyTo(keys, 0);
            foreach (ModelDoc2 key in keys)
            {
                docHandler = (DocumentEventHandler)openDocs[key];
                docHandler.DetachEventHandlers(); //This also removes the pair from the hash
                docHandler = null;
            }
            return true;
        }

        #endregion Event Methods

        #region Event Handlers

        //Events
        public int OnDocChange()
        {
            return 0;
        }

        public int OnDocLoad(string docTitle, string docPath)
        {
            return 0;
        }

        private int FileOpenPostNotify(string FileName)
        {
            AttachEventsToAllDocuments();
            return 0;
        }

        public int OnFileNew(object newDoc, int docType, string templateName)
        {
            AttachEventsToAllDocuments();
            return 0;
        }

        public int OnModelChange()
        {
            return 0;
        }

        #endregion Event Handlers
    }
}