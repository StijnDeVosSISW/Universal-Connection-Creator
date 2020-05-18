using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using NXOpen;
using NXOpen.BlockStyler;

namespace UCCreator
{
    public class UniversalConnectionCreator
    {
        //class members
        private static Session theSession = null;
        private static UI theUI = null;
        private static ListingWindow lw = null;
        private string theDlxFileName;
        private NXOpen.BlockStyler.BlockDialog theDialog;
        private NXOpen.BlockStyler.Group group0;// Block type: Group
        private NXOpen.BlockStyler.Label label0;// Block type: Label
        private NXOpen.BlockStyler.Separator separator0;// Block type: Separator
        private NXOpen.BlockStyler.Tree tree_control0;// Block type: Tree Control
        private NXOpen.BlockStyler.Group group1;// Block type: Group
        private NXOpen.BlockStyler.Group group_SavedLists;// Block type: Group
        private NXOpen.BlockStyler.Enumeration enum_SavedLists;// Block type: Enumeration
        private NXOpen.BlockStyler.Button button_IMPORT_SavedLists;// Block type: Button
        private NXOpen.BlockStyler.Group group;// Block type: Group
        private NXOpen.BlockStyler.FileSelection nativeFileBrowser0;// Block type: NativeFileBrowser
        private NXOpen.BlockStyler.Button button_IMPORT;// Block type: Button
        private NXOpen.BlockStyler.Enumeration enum0;// Block type: Enumeration
        private NXOpen.BlockStyler.Button button_CREATE;// Block type: Button

        private static List<NXOpen.BlockStyler.Node> allNodes = new List<Node>();
        private static List<MODELS.BoltDefinition> allBoltDefinitions = new List<MODELS.BoltDefinition>();
        private enum MenuID { AddNode, DeleteNode };

        private enum TargEnv { Production, Debug, Siemens };

        private static List<NXOpen.CAE.AssyFemPart> allAFEMs = new List<NXOpen.CAE.AssyFemPart>();
        private static List<NXOpen.CAE.FemPart> allFEMs = new List<NXOpen.CAE.FemPart>();

        private static string StorageFileName = "UCCreator_SavedBoltDefinitions";  // Name of Excel file in which content of Universal Conn Def tree will be stored for later use
        private static string StoragePath_server = null;
        private static string StoragePath_user = null;

        private static bool ProcessAll = true;
        private static NXOpen.BasePart currWork = null;

        //------------------------------------------------------------------------------
        //Constructor for NX Styler class
        //------------------------------------------------------------------------------
        public UniversalConnectionCreator()
        {
            try
            {
                theSession = Session.GetSession();
                theUI = UI.GetUI();
                lw = theSession.ListingWindow;

                // Set path to GUI .dlx file
                //TargEnv targEnv = TargEnv.Production;
                //TargEnv targEnv = TargEnv.Debug;
                TargEnv targEnv = TargEnv.Siemens;

                switch (targEnv)
                {
                    case TargEnv.Production:
                        theDlxFileName = @"D:\NX\CAE\UBC\ABC\UniversalConnectionCreator\UniversalConnectionCreator.dlx";  // IN CPP TC environment as Production tool
                        break;
                    case TargEnv.Debug:
                        theDlxFileName = @"C:\sdevos\ABC NXOpen applications\ABC applications\application\UniversalConnectionCreator.dlx";  // Debug by Stijn in CPP TC environment
                        break;
                    case TargEnv.Siemens:
                        theDlxFileName = "UniversalConnectionCreator.dlx";  // In Siemens TC environment
                        break;
                    default:
                        break;
                }
                
                theDialog = theUI.CreateDialog(theDlxFileName);
                theDialog.AddUpdateHandler(new NXOpen.BlockStyler.BlockDialog.Update(update_cb));
                theDialog.AddInitializeHandler(new NXOpen.BlockStyler.BlockDialog.Initialize(initialize_cb));
                theDialog.AddDialogShownHandler(new NXOpen.BlockStyler.BlockDialog.DialogShown(dialogShown_cb));
            }
            catch (Exception ex)
            {
                //---- Enter your exception handling code here -----
                throw ex;
            }
        }
        //------------------------------- DIALOG LAUNCHING ---------------------------------
        //
        //    Before invoking this application one needs to open any part/empty part in NX
        //    because of the behavior of the blocks.
        //
        //    Make sure the dlx file is in one of the following locations:
        //        1.) From where NX session is launched
        //        2.) $UGII_USER_DIR/application
        //        3.) For released applications, using UGII_CUSTOM_DIRECTORY_FILE is highly
        //            recommended. This variable is set to a full directory path to a file 
        //            containing a list of root directories for all custom applications.
        //            e.g., UGII_CUSTOM_DIRECTORY_FILE=$UGII_BASE_DIR\ugii\menus\custom_dirs.dat
        //
        //    You can create the dialog using one of the following way:
        //
        //    1. Journal Replay
        //
        //        1) Replay this file through Tool->Journal->Play Menu.
        //
        //    2. USER EXIT
        //
        //        1) Create the Shared Library -- Refer "Block UI Styler programmer's guide"
        //        2) Invoke the Shared Library through File->Execute->NX Open menu.
        //
        //------------------------------------------------------------------------------
        public static void Main()
        {
            UniversalConnectionCreator theUniversalConnectionCreator = null;
            try
            {
                theUniversalConnectionCreator = new UniversalConnectionCreator();
                // The following method shows the dialog immediately
                theUniversalConnectionCreator.Show();
            }
            catch (Exception ex)
            {
                //---- Enter your exception handling code here -----
                theUI.NXMessageBox.Show("Block Styler", NXMessageBox.DialogType.Error, ex.ToString());
            }
            finally
            {
                if (theUniversalConnectionCreator != null)
                    theUniversalConnectionCreator.Dispose();
                theUniversalConnectionCreator = null;
            }
        }
        //------------------------------------------------------------------------------
        // This method specifies how a shared image is unloaded from memory
        // within NX. This method gives you the capability to unload an
        // internal NX Open application or user  exit from NX. Specify any
        // one of the three constants as a return value to determine the type
        // of unload to perform:
        //
        //
        //    Immediately : unload the library as soon as the automation program has completed
        //    Explicitly  : unload the library from the "Unload Shared Image" dialog
        //    AtTermination : unload the library when the NX session terminates
        //
        //
        // NOTE:  A program which associates NX Open applications with the menubar
        // MUST NOT use this option since it will UNLOAD your NX Open application image
        // from the menubar.
        //------------------------------------------------------------------------------
        public static int GetUnloadOption(string arg)
        {
            //return System.Convert.ToInt32(Session.LibraryUnloadOption.Explicitly);
            return System.Convert.ToInt32(Session.LibraryUnloadOption.Immediately);
            // return System.Convert.ToInt32(Session.LibraryUnloadOption.AtTermination);
        }

        //------------------------------------------------------------------------------
        // Following method cleanup any housekeeping chores that may be needed.
        // This method is automatically called by NX.
        //------------------------------------------------------------------------------
        public static void UnloadLibrary(string arg)
        {
            try
            {
                //---- Enter your code here -----
            }
            catch (Exception ex)
            {
                //---- Enter your exception handling code here -----
                theUI.NXMessageBox.Show("Block Styler", NXMessageBox.DialogType.Error, ex.ToString());
            }
        }

        //------------------------------------------------------------------------------
        //This method shows the dialog on the screen
        //------------------------------------------------------------------------------
        public NXOpen.UIStyler.DialogResponse Show()
        {
            try
            {
                theDialog.Show();
            }
            catch (Exception ex)
            {
                //---- Enter your exception handling code here -----
                theUI.NXMessageBox.Show("Block Styler", NXMessageBox.DialogType.Error, ex.ToString());
            }
            return 0;
        }

        //------------------------------------------------------------------------------
        //Method Name: Dispose
        //------------------------------------------------------------------------------
        public void Dispose()
        {
            if (theDialog != null)
            {
                // Store current Tree List content to use in next session
                //StoreUnivConnList();

                theDialog.Dispose();
                theDialog = null;
            }
        }

        //------------------------------------------------------------------------------
        //---------------------Block UI Styler Callback Functions--------------------------
        //------------------------------------------------------------------------------

        //------------------------------------------------------------------------------
        //Callback Name: initialize_cb
        //------------------------------------------------------------------------------
        public void initialize_cb()
        {
            try
            {
                group0 = (NXOpen.BlockStyler.Group)theDialog.TopBlock.FindBlock("group0");
                label0 = (NXOpen.BlockStyler.Label)theDialog.TopBlock.FindBlock("label0");
                separator0 = (NXOpen.BlockStyler.Separator)theDialog.TopBlock.FindBlock("separator0");
                tree_control0 = (NXOpen.BlockStyler.Tree)theDialog.TopBlock.FindBlock("tree_control0");
                group1 = (NXOpen.BlockStyler.Group)theDialog.TopBlock.FindBlock("group1");
                group_SavedLists = (NXOpen.BlockStyler.Group)theDialog.TopBlock.FindBlock("group_SavedLists");
                enum_SavedLists = (NXOpen.BlockStyler.Enumeration)theDialog.TopBlock.FindBlock("enum_SavedLists");
                button_IMPORT_SavedLists = (NXOpen.BlockStyler.Button)theDialog.TopBlock.FindBlock("button_IMPORT_SavedLists");
                group = (NXOpen.BlockStyler.Group)theDialog.TopBlock.FindBlock("group");
                nativeFileBrowser0 = (NXOpen.BlockStyler.FileSelection)theDialog.TopBlock.FindBlock("nativeFileBrowser0");
                button_IMPORT = (NXOpen.BlockStyler.Button)theDialog.TopBlock.FindBlock("button_IMPORT");
                enum0 = (NXOpen.BlockStyler.Enumeration)theDialog.TopBlock.FindBlock("enum0");
                button_CREATE = (NXOpen.BlockStyler.Button)theDialog.TopBlock.FindBlock("button_CREATE");

                // Initialize storage paths
                StoragePath_server = System.Reflection.Assembly.GetExecutingAssembly().CodeBase;
                StoragePath_server = StoragePath_server.Remove(StoragePath_server.LastIndexOf("/")) + "\\" + StorageFileName + ".txt";
                StoragePath_server = StoragePath_server.Substring(StoragePath_server.IndexOf("/")).Substring(3).Replace("/", "\\");

                StoragePath_user = Environment.GetEnvironmentVariable("USERPROFILE") + "\\AppData\\Local\\UniversalConnectionCreator\\" + StorageFileName + ".txt"; 

                //------------------------------------------------------------------------------
                //Registration of Treelist specific callbacks
                //------------------------------------------------------------------------------
                //tree_control0.SetOnExpandHandler(new NXOpen.BlockStyler.Tree.OnExpandCallback(OnExpandCallback));

                //tree_control0.SetOnInsertColumnHandler(new NXOpen.BlockStyler.Tree.OnInsertColumnCallback(OnInsertColumnCallback));

                //tree_control0.SetOnInsertNodeHandler(new NXOpen.BlockStyler.Tree.OnInsertNodeCallback(OnInsertNodecallback));

                //tree_control0.SetOnDeleteNodeHandler(new NXOpen.BlockStyler.Tree.OnDeleteNodeCallback(OnDeleteNodecallback));

                //tree_control0.SetOnPreSelectHandler(new NXOpen.BlockStyler.Tree.OnPreSelectCallback(OnPreSelectcallback));

                tree_control0.SetOnSelectHandler(new NXOpen.BlockStyler.Tree.OnSelectCallback(OnSelectcallback));

                //tree_control0.SetOnStateChangeHandler(new NXOpen.BlockStyler.Tree.OnStateChangeCallback(OnStateChangecallback));

                //tree_control0.SetToolTipTextHandler(new NXOpen.BlockStyler.Tree.ToolTipTextCallback(ToolTipTextcallback));

                //tree_control0.SetColumnSortHandler(new NXOpen.BlockStyler.Tree.ColumnSortCallback(ColumnSortcallback));

                //tree_control0.SetStateIconNameHandler(new NXOpen.BlockStyler.Tree.StateIconNameCallback(StateIconNameCallback));

                tree_control0.SetOnBeginLabelEditHandler(new NXOpen.BlockStyler.Tree.OnBeginLabelEditCallback(OnBeginLabelEditCallback));

                tree_control0.SetOnEndLabelEditHandler(new NXOpen.BlockStyler.Tree.OnEndLabelEditCallback(OnEndLabelEditCallback));

                tree_control0.SetOnEditOptionSelectedHandler(new NXOpen.BlockStyler.Tree.OnEditOptionSelectedCallback(OnEditOptionSelectedCallback));

                tree_control0.SetAskEditControlHandler(new NXOpen.BlockStyler.Tree.AskEditControlCallback(AskEditControlCallback));

                tree_control0.SetOnMenuHandler(new NXOpen.BlockStyler.Tree.OnMenuCallback(OnMenuCallback));

                tree_control0.SetOnMenuSelectionHandler(new NXOpen.BlockStyler.Tree.OnMenuSelectionCallback(OnMenuSelectionCallback));

                tree_control0.SetIsDropAllowedHandler(new NXOpen.BlockStyler.Tree.IsDropAllowedCallback(IsDropAllowedCallback));

                tree_control0.SetIsDragAllowedHandler(new NXOpen.BlockStyler.Tree.IsDragAllowedCallback(IsDragAllowedCallback));

                tree_control0.SetOnDropHandler(new NXOpen.BlockStyler.Tree.OnDropCallback(OnDropCallback));

                //tree_control0.SetOnDropMenuHandler(new NXOpen.BlockStyler.Tree.OnDropMenuCallback(OnDropMenuCallback));

                //tree_control0.SetOnDefaultActionHandler(new NXOpen.BlockStyler.Tree.OnDefaultActionCallback(OnDefaultActionCallback));

                //------------------------------------------------------------------------------
            }
            catch (Exception ex)
            {
                //---- Enter your exception handling code here -----
                theUI.NXMessageBox.Show("Block Styler", NXMessageBox.DialogType.Error, ex.ToString());
            }
        }

        //------------------------------------------------------------------------------
        //Callback Name: dialogShown_cb
        //This callback is executed just before the dialog launch. Thus any value set 
        //here will take precedence and dialog will be launched showing that value. 
        //------------------------------------------------------------------------------
        public void dialogShown_cb()
        {
            try
            {
                // Initialize ListingWindow
                lw.WriteFullline(
                " ------------------------------ " + Environment.NewLine +
                " ------------------------------ " + Environment.NewLine +
                "| UNIVERSAL CONNECTION CREATOR |" + Environment.NewLine +
                " ------------------------------ " + Environment.NewLine +
                " ------------------------------ " + Environment.NewLine);

                lw.Open();

                // Initialize GUI
                nativeFileBrowser0.Path = "";
                button_IMPORT_SavedLists.Enable = true;
                enum_SavedLists.SetBalloonTooltipTexts(new string[]{ StoragePath_user, StoragePath_server });

                // Initialize Tree Control (List of predefined Universal Bolt Connections)
                int default_width = 150;
                tree_control0.InsertColumn(0, "Name", default_width);
                tree_control0.InsertColumn(1, "Shank Diameter [mm]", default_width);
                tree_control0.InsertColumn(2, "Head Diameter [mm]", default_width);
                tree_control0.InsertColumn(3, "Maximum Connection Length [mm]", default_width);
                tree_control0.InsertColumn(4, "Material", default_width);

                // Import stored Bolt Definitions
                if (File.Exists(StoragePath_user))
                {
                    ImportStoredBoltDefinitions(StoragePath_user);
                }
                else
                {
                    ImportStoredBoltDefinitions(StoragePath_server);
                }

                // Check value of process level switch
                switch (enum0.ValueAsString)
                {
                    case "This level and all sub-levels":
                        ProcessAll = true;
                        break;

                    case "This level only":
                        ProcessAll = false;
                        break;

                    default:
                        break;
                }

                // Get current working object
                currWork = theSession.Parts.BaseWork;
            }
            catch (Exception ex)
            {
                //---- Enter your exception handling code here -----
                theUI.NXMessageBox.Show("Block Styler", NXMessageBox.DialogType.Error, ex.ToString());
            }
        }

        //------------------------------------------------------------------------------
        //Callback Name: update_cb
        //------------------------------------------------------------------------------
        public int update_cb(NXOpen.BlockStyler.UIBlock block)
        {
            try
            {
                if (block == label0)
                {
                    //---------Enter your code here-----------
                }
                else if (block == separator0)
                {
                    //---------Enter your code here-----------
                }
                else if (block == enum_SavedLists)
                {
                    //---------Enter your code here-----------
                }
                else if (block == button_IMPORT_SavedLists)
                {
                    switch (enum_SavedLists.ValueAsString)
                    {
                        case "Last used by you":
                            if (File.Exists(StoragePath_user))
                            {
                                ImportStoredBoltDefinitions(StoragePath_user);
                            }
                            else
                            {
                                theUI.NXMessageBox.Show("Universal Connection Creator", NXMessageBox.DialogType.Error, "No saved User list found to import!");
                            }
                            break;

                        case "Default list":
                            if (File.Exists(StoragePath_server))
                            {
                                ImportStoredBoltDefinitions(StoragePath_server);
                            }
                            else
                            {
                                theUI.NXMessageBox.Show("Universal Connection Creator", NXMessageBox.DialogType.Error, "No saved Default list found to import!");
                            }
                            break;

                        default:
                            break;
                    }
                    if (nativeFileBrowser0.Path != "" && File.Exists(nativeFileBrowser0.Path))
                    {
                        ImportDefsFromExcel(nativeFileBrowser0.Path);
                    }
                }
                else if (block == nativeFileBrowser0)
                {
                    button_IMPORT.Enable = false;

                    if (nativeFileBrowser0.Path != null)
                    {
                        if (File.Exists(nativeFileBrowser0.Path))
                        {
                            button_IMPORT.Enable = true;
                        }
                    }
                }
                else if (block == button_IMPORT)
                {
                    if (nativeFileBrowser0.Path != "" && File.Exists(nativeFileBrowser0.Path))
                    {
                        ImportDefsFromExcel(nativeFileBrowser0.Path);
                    }
                }
                else if (block == enum0)
                {
                    //lw.WriteFullline("enum0.ValueAsString = " + enum0.ValueAsString);

                    switch (enum0.ValueAsString)
                    {
                        case "This level and all sub-levels":
                            ProcessAll = true;
                            break;

                        case "This level only":
                            ProcessAll = false;
                            break;

                        default:
                            break;
                    }
                }
                else if (block == button_CREATE)
                {
                    // Execute Universal Bolt Connection creation
                    ExecuteBoltGeneration();

                    // Store current Tree List content to use in next session
                    StoreUnivConnList();
                }
            }
            catch (Exception ex)
            {
                //---- Enter your exception handling code here -----
                theUI.NXMessageBox.Show("Block Styler", NXMessageBox.DialogType.Error, ex.ToString());
            }
            return 0;
        }

        //------------------------------------------------------------------------------
        //Treelist specific callbacks
        //------------------------------------------------------------------------------
        //public void OnExpandCallback(NXOpen.BlockStyler.Tree tree, NXOpen.BlockStyler.Node node)
        //{
        //}

        //public void OnInsertColumnCallback(NXOpen.BlockStyler.Tree tree, NXOpen.BlockStyler.Node node, int columnID)
        //{
        //}

        //public void OnInsertNodecallback(NXOpen.BlockStyler.Tree tree, NXOpen.BlockStyler.Node node)
        //{
        //}

        //public void OnDeleteNodecallback(NXOpen.BlockStyler.Tree tree, NXOpen.BlockStyler.Node node)
        //{
        //}

        //public void OnPreSelectcallback(NXOpen.BlockStyler.Tree tree, NXOpen.BlockStyler.Node node, int columnID, bool Selected)
        //{
        //}

        public void OnSelectcallback(NXOpen.BlockStyler.Tree tree, NXOpen.BlockStyler.Node node, int columnID, bool Selected)
        {
            
        }

        //public void OnStateChangecallback(NXOpen.BlockStyler.Tree tree, NXOpen.BlockStyler.Node node, int State)
        //{
        //}

        //public string ToolTipTextcallback(NXOpen.BlockStyler.Tree tree, NXOpen.BlockStyler.Node node, int columnID)
        //{
        //}

        //public int ColumnSortcallback(NXOpen.BlockStyler.Tree tree, int columnID, NXOpen.BlockStyler.Node node1, NXOpen.BlockStyler.Node node2)
        //{
        //}

        //public string StateIconNameCallback(NXOpen.BlockStyler.Tree tree, NXOpen.BlockStyler.Node node, int state)
        //{
        //}

        public Tree.BeginLabelEditState OnBeginLabelEditCallback(NXOpen.BlockStyler.Tree tree, NXOpen.BlockStyler.Node node, int columnID)
        {
            return Tree.BeginLabelEditState.Allow;
        }

        public Tree.EndLabelEditState OnEndLabelEditCallback(NXOpen.BlockStyler.Tree tree, NXOpen.BlockStyler.Node node, int columnID, string editedText)
        {
            return Tree.EndLabelEditState.AcceptText;
        }

        public Tree.EditControlOption OnEditOptionSelectedCallback(NXOpen.BlockStyler.Tree tree, NXOpen.BlockStyler.Node node, int columnID, int selectedOptionID, string selectedOptionText, Tree.ControlType type)
        {
            Tree.EditControlOption OnEditOptionSelected = NXOpen.BlockStyler.Tree.EditControlOption.Accept;
            return OnEditOptionSelected;
        }

        public Tree.ControlType AskEditControlCallback(NXOpen.BlockStyler.Tree tree, NXOpen.BlockStyler.Node node, int columnID)
        {
            return Tree.ControlType.ComboBox;
        }

        public void OnMenuCallback(NXOpen.BlockStyler.Tree tree, NXOpen.BlockStyler.Node node, int columnID)
        {
            NXOpen.BlockStyler.TreeListMenu treeMenu = tree.CreateMenu();
            treeMenu.AddMenuItem((int)MenuID.DeleteNode, "Delete");
            treeMenu.AddSeparator();
            treeMenu.AddMenuItem((int)MenuID.AddNode, "Add new Bolt Definition below");
            tree_control0.SetMenu(treeMenu);
            treeMenu.Dispose();
        }

        public void OnMenuSelectionCallback(NXOpen.BlockStyler.Tree tree, NXOpen.BlockStyler.Node node, int menuItemID)
        {
            switch (menuItemID)
            {
                case (int)MenuID.DeleteNode:
                    tree.DeleteNode(node);
                    allNodes.Remove(node);
                    break;

                case (int)MenuID.AddNode:
                    NXOpen.BlockStyler.Node newNode = tree_control0.CreateNode("<name>");

                    tree_control0.InsertNode(newNode, null, node, Tree.NodeInsertOption.First);
                    allNodes.Insert(allNodes.IndexOf(node) + 1, newNode);

                    newNode.SetColumnDisplayText(0, "<name>");
                    newNode.SetColumnDisplayText(1, "<shank diameter>");
                    newNode.SetColumnDisplayText(2, "<head diameter>");
                    newNode.SetColumnDisplayText(3, "<max connection length>");
                    newNode.SetColumnDisplayText(4, "<material name>");
                    break;

                default:
                    break;
            }
        }

        public Node.DropType IsDropAllowedCallback(NXOpen.BlockStyler.Tree tree, NXOpen.BlockStyler.Node node, int columnID, NXOpen.BlockStyler.Node targetNode, int targetColumnID)
        {
            return Node.DropType.After;
        }

        public Node.DragType IsDragAllowedCallback(NXOpen.BlockStyler.Tree tree, NXOpen.BlockStyler.Node node, int columnID)
        {
            return Node.DragType.All;
        }

        public bool OnDropCallback(NXOpen.BlockStyler.Tree tree, NXOpen.BlockStyler.Node[] node, int columnID, NXOpen.BlockStyler.Node targetNode, int targetColumnID, Node.DropType dropType, int dropMenuItemId)
        {
            switch (dropType)
            {
                case Node.DropType.None:
                    break;
                case Node.DropType.On:
                    break;
                case Node.DropType.Before:
                    break;
                case Node.DropType.After:
                    NXOpen.BlockStyler.Node movedNode = tree.CreateNode("new");

                    tree.InsertNode(movedNode, null, targetNode, Tree.NodeInsertOption.First);

                    movedNode.SetColumnDisplayText(0, node[0].GetColumnDisplayText(0));
                    movedNode.SetColumnDisplayText(1, node[0].GetColumnDisplayText(1));
                    movedNode.SetColumnDisplayText(2, node[0].GetColumnDisplayText(2));
                    movedNode.SetColumnDisplayText(3, node[0].GetColumnDisplayText(3));
                    movedNode.SetColumnDisplayText(4, node[0].GetColumnDisplayText(4));

                    allNodes.Insert(allNodes.IndexOf(node[0]), movedNode);

                    tree.DeleteNode(node[0]);
                    allNodes.Remove(node[0]);
                    
                    break;
                case Node.DropType.BeforeAndAfter:
                    break;
                default:
                    break;
            }
            return true;
        }

        //public void OnDropMenuCallback(NXOpen.BlockStyler.Tree tree, NXOpen.BlockStyler.Node node, int columnID, NXOpen.BlockStyler.Node targetNode, int targetColumnID)
        //{
        //}

        //public void OnDefaultActionCallback(NXOpen.BlockStyler.Tree tree, NXOpen.BlockStyler.Node node, int columnID)
        //{
        //}

        //------------------------------------------------------------------------------
        //StringBlock specific callbacks
        //------------------------------------------------------------------------------
        //public int KeystrokeCallback(NXOpen.BlockStyler.StringBlock string_block, string uncommitted_value)
        //{
        //}

        //------------------------------------------------------------------------------

        //------------------------------------------------------------------------------
        //Function Name: GetBlockProperties
        //Returns the propertylist of the specified BlockID
        //------------------------------------------------------------------------------
        public PropertyList GetBlockProperties(string blockID)
        {
            PropertyList plist = null;
            try
            {
                plist = theDialog.GetBlockProperties(blockID);
            }
            catch (Exception ex)
            {
                //---- Enter your exception handling code here -----
                theUI.NXMessageBox.Show("Block Styler", NXMessageBox.DialogType.Error, ex.ToString());
            }
            return plist;
        }


        #region CUSTOM METHODS
        /// <summary>
        /// Import predefined Bolt Definitions from an Excel file
        /// </summary>
        /// <param name="filePath">Full path to target Excel file</param>
        private void ImportDefsFromExcel(string filePath)
        {
            try
            {
                lw.WriteFullline(Environment.NewLine +
                    " ----------------------------------------- " + Environment.NewLine +
                    "| IMPORT BOLT DEFINITIONS FROM EXCEL FILE |" + Environment.NewLine +
                    " ----------------------------------------- ");

                lw.WriteFullline("Input Excel file  :  " + filePath);

                // Clear all existing nodes and BoltDefinition objects
                foreach (NXOpen.BlockStyler.Node myNode in allNodes)
                {
                    tree_control0.DeleteNode(myNode);
                }
                allNodes.Clear();

                lw.WriteFullline(Environment.NewLine + "Delete existing Bolt Definitions :  SUCCESS");

                // Create COM objects to use Excel to read the input Excel file
                Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(filePath);
                Microsoft.Office.Interop.Excel.Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                Microsoft.Office.Interop.Excel.Range targRange = xlWorksheet.UsedRange;

                // Extract each BoltDefinition object from Used Range of Excel sheet
                // -> start at i = 2, because first row in Excel sheet are the column headers! (Excel is NOT zero-based)
                for (int i = 2; i < (targRange.Rows.Count+1); i++)
                {
                    lw.WriteFullline(Environment.NewLine + "IMPORTING: Bolt Definition " + (i - 1).ToString());
                    // Add new node to Tree List
                    NXOpen.BlockStyler.Node newNode = tree_control0.CreateNode("<new>");
                    tree_control0.InsertNode(newNode, null, null, Tree.NodeInsertOption.Last);

                    newNode.SetColumnDisplayText(0, (string)targRange.Cells[i, 1].Text);
                    newNode.SetColumnDisplayText(1, (string)targRange.Cells[i, 2].Text);
                    newNode.SetColumnDisplayText(2, (string)targRange.Cells[i, 3].Text);
                    newNode.SetColumnDisplayText(3, (string)targRange.Cells[i, 4].Text);
                    newNode.SetColumnDisplayText(4, (string)targRange.Cells[i, 5].Text);

                    allNodes.Add(newNode);

                    lw.WriteFullline("   New node created with:" + Environment.NewLine +
                        "   Name                           = " + newNode.GetColumnDisplayText(0) + Environment.NewLine +
                        "   Shank Diameter [mm]            = " + newNode.GetColumnDisplayText(1) + Environment.NewLine +
                        "   Head Diameter [mm]             = " + newNode.GetColumnDisplayText(2) + Environment.NewLine +
                        "   Maximum Connection Length [mm] = " + newNode.GetColumnDisplayText(3) + Environment.NewLine +
                        "   Material Name                  = " + newNode.GetColumnDisplayText(4) + Environment.NewLine);

                    //for (int j = 1; j < (targRange.Columns.Count+1); j++)
                    //{
                    //    lw.WriteFullline("  Cell[" + i.ToString() + "," + j.ToString() + "] = " + targRange.Cells[i, j].Value);
                    //    //newNode.SetColumnDisplayText(j, targRange.Cells[i, j].Value);
                    //}
                }

                //cleanup
                GC.Collect();
                GC.WaitForPendingFinalizers();

                //release com objects to fully kill excel process from running in the background
                Marshal.ReleaseComObject(targRange);
                Marshal.ReleaseComObject(xlWorksheet);

                //close and release
                xlWorkbook.Close();
                Marshal.ReleaseComObject(xlWorkbook);

                //quit and release
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);

                lw.WriteFullline("Released: targRange, xlWorksheet, xlWorkbook, xlApp");
            }
            catch (Exception e)
            {
                lw.WriteFullline("!ERROR occurred: " + Environment.NewLine +
                    e.ToString());
            }
        }

        /// <summary>
        /// Import last used Bolt Definitions, stored by the previous session
        /// </summary>
        private void ImportStoredBoltDefinitions(string filePath)
        {
            try
            {
                //lw.WriteFullline(Environment.NewLine +
                //    " -------------------------------- " + Environment.NewLine +
                //    "| IMPORT STORED BOLT DEFINITIONS |" + Environment.NewLine +
                //    " -------------------------------- ");

                // Get target file
                // ---------------
                //// Get target file path to stored Excel file
                //string filePath = System.Reflection.Assembly.GetExecutingAssembly().CodeBase;
                //filePath = filePath.Remove(filePath.LastIndexOf("/")) + "\\" + StorageFileName + ".txt";
                //filePath = filePath.Substring(filePath.IndexOf("/")).Substring(3).Replace("/", "\\");

                //lw.WriteFullline("File path = " + Environment.NewLine +
                //    filePath);

                // Check if it exists
                if (!File.Exists(filePath))
                {
                    lw.WriteFullline("No stored bolt definitions could be found:  continue...");
                    return;
                }

                // Read all content of target file
                string fileContent = "";
                using (StreamReader sr = File.OpenText(filePath))
                {
                    fileContent = sr.ReadToEnd();
                }


                // Process target file content
                // ---------------------------
                // Split fileContent on NewLine characters
                string[] allLines = fileContent.Split(Environment.NewLine.ToCharArray());

                // Create a list of BoltDefinition objects based on info of each line
                List<MODELS.BoltDefinition> currBoltDefinitions = new List<MODELS.BoltDefinition>();
                foreach (string line in allLines)
                {
                    if (!line.Contains("DIAMETER") && line != "") // Makes sure we don't process the header content and that line content is not nothing
                    {
                        //lw.WriteFullline("   PROCESSING:  " + line);
                        string currLine = line;
                        MODELS.BoltDefinition newBoltDefinition = new MODELS.BoltDefinition();

                        // Name
                        newBoltDefinition.Name = currLine.Remove(currLine.IndexOf("|") - 1);
                        currLine = currLine.Substring(currLine.IndexOf("|") + 2);

                        // Shank Diameter
                        newBoltDefinition.ShankDiam = Convert.ToInt32(currLine.Remove(currLine.IndexOf("|") - 1));
                        currLine = currLine.Substring(currLine.IndexOf("|") + 2);

                        // Head Diameter
                        newBoltDefinition.HeadDiam = Convert.ToInt32(currLine.Remove(currLine.IndexOf("|") - 1));
                        currLine = currLine.Substring(currLine.IndexOf("|") + 2);

                        // Maximum Connection Length
                        newBoltDefinition.MaxConnLength = Convert.ToInt32(currLine.Remove(currLine.IndexOf("|") - 1));
                        currLine = currLine.Substring(currLine.IndexOf("|") + 2);

                        // Material
                        newBoltDefinition.MaterialName = currLine;

                        currBoltDefinitions.Add(newBoltDefinition);
                    }
                }


                // Import as new Nodes in Tree List
                // --------------------------------
                // Clear all existing nodes and BoltDefinition objects
                foreach (NXOpen.BlockStyler.Node myNode in allNodes)
                {
                    tree_control0.DeleteNode(myNode);
                }
                allNodes.Clear();

                //lw.WriteFullline(Environment.NewLine + "Delete existing Bolt Definitions :  SUCCESS");

                // Import new Bolt Definitions
                int i = 1;
                foreach (MODELS.BoltDefinition boltDefinition in currBoltDefinitions)
                {
                    //lw.WriteFullline(Environment.NewLine + "IMPORTING: Bolt Definition " + i.ToString());

                    // Add new node to Tree List
                    NXOpen.BlockStyler.Node newNode = tree_control0.CreateNode("<new>");
                    tree_control0.InsertNode(newNode, null, null, Tree.NodeInsertOption.Last);

                    newNode.SetColumnDisplayText(0, boltDefinition.Name);
                    newNode.SetColumnDisplayText(1, boltDefinition.ShankDiam.ToString());
                    newNode.SetColumnDisplayText(2, boltDefinition.HeadDiam.ToString());
                    newNode.SetColumnDisplayText(3, boltDefinition.MaxConnLength.ToString());
                    newNode.SetColumnDisplayText(4, boltDefinition.MaterialName);

                    allNodes.Add(newNode);

                    //lw.WriteFullline("   New node created with:" + Environment.NewLine +
                    //    "   Name                           = " + newNode.GetColumnDisplayText(0) + Environment.NewLine +
                    //    "   Shank Diameter [mm]            = " + newNode.GetColumnDisplayText(1) + Environment.NewLine +
                    //    "   Head Diameter [mm]             = " + newNode.GetColumnDisplayText(2) + Environment.NewLine +
                    //    "   Maximum Connection Length [mm] = " + newNode.GetColumnDisplayText(3) + Environment.NewLine +
                    //    "   Material Name                  = " + newNode.GetColumnDisplayText(4) + Environment.NewLine);

                    i++;
                }

                lw.WriteFullline("Load stored Bolt Definitions:  SUCCESS");
            }
            catch (Exception e)
            {
                lw.WriteFullline("!ERROR occurred: " + Environment.NewLine +
                    e.ToString());
            }
        }

        /// <summary>
        /// Store current Universal Bolt Connection definitions in an Excel file
        /// </summary>
        private void StoreUnivConnList()
        {
            try
            {
                lw.WriteFullline(Environment.NewLine +
                    " ----------------------------------------- " + Environment.NewLine +
                    "| SAVE BOLT DEFINITION LIST FOR LATER USE |" + Environment.NewLine +
                    " ----------------------------------------- ");

                // Get target Excel file path
                //string filePath = System.Reflection.Assembly.GetExecutingAssembly().CodeBase;
                //filePath = filePath.Remove(filePath.LastIndexOf("/")) + "\\"+ StorageFileName + ".txt";
                //filePath = filePath.Substring(filePath.IndexOf("/")).Substring(3).Replace("/", "\\");
                string filePath = StoragePath_user;

                //lw.WriteFullline("Target file path :  " + filePath);

                // Create new Text file
                if (File.Exists(filePath)) { File.Delete(filePath); }
                //lw.WriteFullline("Write content...");

                // Check if target Directory exists, if not, try to create it
                Directory.CreateDirectory(filePath.Remove(filePath.LastIndexOf(@"\")));

                using (StreamWriter sw = File.CreateText(filePath))
                {
                    // Write Headers
                    sw.WriteLine("NAME | SHANK DIAMETER [mm] | HEAD DIAMETER [mm] | MAXIMUM CONNECTION LENGTH [mm] | MATERIAL");
                    //lw.WriteFullline("   Added Headers");

                    // Write content
                    int i = 1;
                    foreach (NXOpen.BlockStyler.Node node in allNodes)
                    {
                        sw.WriteLine(
                            node.GetColumnDisplayText(0) + " | " +
                            node.GetColumnDisplayText(1) + " | " +
                            node.GetColumnDisplayText(2) + " | " +
                            node.GetColumnDisplayText(3) + " | " +
                            node.GetColumnDisplayText(4));

                        //lw.WriteFullline("   Added Node " + i.ToString());

                        i++;
                    }
                }

                lw.WriteFullline("Saved");
            }
            catch (Exception e)
            {
                lw.WriteFullline("!ERROR occurred: " + Environment.NewLine +
                    e.ToString());
            }
        }

        /// <summary>
        /// Start Bolt generation execution
        /// </summary>
        private void ExecuteBoltGeneration()
        {
            try
            {
                lw.WriteFullline(Environment.NewLine + Environment.NewLine +
                    " ------------------------- " + Environment.NewLine +
                    "| EXECUTE BOLT GENERATION |" + Environment.NewLine +
                    " ------------------------- ");

                // CREATE LIST OF PREDEFINED BOLT CONNECTIONS
                lw.WriteFullline("Generate list of predefined Bolt Connections...");

                allBoltDefinitions.Clear();

                foreach (NXOpen.BlockStyler.Node node in allNodes)
                {
                    allBoltDefinitions.Add(new MODELS.BoltDefinition()
                    {
                        Name = node.GetColumnDisplayText(0),
                        ShankDiam = Convert.ToInt32(node.GetColumnDisplayText(1)),
                        HeadDiam = Convert.ToInt32(node.GetColumnDisplayText(2)),
                        MaxConnLength = Convert.ToInt32(node.GetColumnDisplayText(3)),
                        MaterialName = node.GetColumnDisplayText(4)
                    }); ;

                    lw.WriteFullline("   Added Bolt Definition:  " + node.GetColumnDisplayText(0).ToUpper());
                }

                // Loop through all (A)FEM objects
                lw.WriteFullline(Environment.NewLine +
                    " --------------------------------- " + Environment.NewLine +
                    "| Loop through all (A)FEM objects |" + Environment.NewLine +
                    " --------------------------------- ");

                currWork = theSession.Parts.BaseWork;
                lw.WriteFullline("Current working object :  " + currWork.ToString());

                switch (theSession.Parts.BaseWork.GetType().ToString())
                {
                    case "NXOpen.CAE.SimPart":
                        lw.WriteFullline("---> Recognized as SIM");
                        NXOpen.CAE.SimPart mySIM = (NXOpen.CAE.SimPart)theSession.Parts.BaseWork;

                        switch (mySIM.FemPart.GetType().ToString())
                        {
                            case "NXOpen.CAE.AssyFemPart":
                                lw.WriteFullline("---> Underlying CAE object = recognized as AFEM");
                                ProcessFromAFEM((NXOpen.CAE.AssyFemPart)mySIM.FemPart);
                                break;

                            case "NXOpen.CAE.FemPart":
                                lw.WriteFullline("---> Underlying CAE object = recognized as FEM");
                                ProcessFromFEM((NXOpen.CAE.FemPart)mySIM.FemPart);
                                break;

                            default:
                                lw.WriteFullline("---> Underlying CAE object = recognized as " + mySIM.FemPart.GetType().ToString() + " -> SKIPPED");
                                break;
                        }

                        break;

                    case "NXOpen.CAE.AssyFemPart":
                        lw.WriteFullline("---> Recognized as AFEM");
                        ProcessFromAFEM((NXOpen.CAE.AssyFemPart)theSession.Parts.BaseWork);
                        break;

                    case "NXOpen.CAE.FemPart":
                        lw.WriteFullline("---> Recognized as FEM");
                        ProcessFromFEM((NXOpen.CAE.FemPart)theSession.Parts.BaseWork);
                        break;

                    default:
                        lw.WriteFullline("---> Not recognized as SIM, AFEM or FEM, but as:  " + theSession.Parts.BaseWork.GetType().ToString() + Environment.NewLine +
                            "=> exiting...");
                        return;
                        break;
                }


                // UPDATE EACH AFEM, IF NECESSARY
                // ------------------------------
                lw.WriteFullline(Environment.NewLine +
                    " ---------------------------------------------- " + Environment.NewLine +
                    "| Update AFEM objects that have pending update |" + Environment.NewLine +
                    " ---------------------------------------------- ");

                //allAFEMs.Reverse(); // To iterate bottom-up, instead of top-down

                foreach (NXOpen.CAE.AssyFemPart myAFEM in allAFEMs)
                {
                    // Set current AFEM to working
                    theSession.Parts.SetWork(myAFEM);

                    // Force "update" status for each Universal Bolt Connection
                    foreach (NXOpen.CAE.Connections.IConnection myConn in myAFEM.BaseFEModel.ConnectionsContainer.GetAllConnections())
                    {
                        try
                        {
                            // Check if it is a Universal Bolt Connection
                            NXOpen.CAE.Connections.Bolt myBoltConn = (NXOpen.CAE.Connections.Bolt)myConn;

                            // Force an "update" of the Universal Bolt Connection
                            myBoltConn.MaxBoltLength.Value++;
                            myBoltConn.MaxBoltLength.Value--;

                            lw.WriteFullline("   Update forced of:  " + myBoltConn.Name.ToUpper());
                        }
                        catch (Exception)
                        {
                            // Not a Universal Bolt Connection
                        }
                    }

                    // Update AFEM to realize all Universal Bolt Connections
                    //if (myAFEM.BaseFEModel.AskUpdatePending())
                    //{
                        myAFEM.BaseFEModel.UpdateFemodel();
                        lw.WriteFullline("   UPDATED:  " + myAFEM.Name.ToUpper());
                    //}
                }


                // UPDATE EACH (normal) FEM, IF NECESSARY
                // --------------------------------------
                lw.WriteFullline(Environment.NewLine +
                    " ------------------------------------------------------ " + Environment.NewLine +
                    "| Update (normal) FEM objects that have pending update |" + Environment.NewLine +
                    " ------------------------------------------------------ ");

                foreach (NXOpen.CAE.FemPart myFEM in allFEMs)
                {
                    // Set current AFEM to working
                    theSession.Parts.SetWork(myFEM);

                    // Force "update" status for each Universal Bolt Connection
                    foreach (NXOpen.CAE.Connections.IConnection myConn in myFEM.BaseFEModel.ConnectionsContainer.GetAllConnections())
                    {
                        try
                        {
                            // Check if it is a Universal Bolt Connection
                            NXOpen.CAE.Connections.Bolt myBoltConn = (NXOpen.CAE.Connections.Bolt)myConn;

                            // Force an "update" of the Universal Bolt Connection
                            myBoltConn.MaxBoltLength.Value++;
                            myBoltConn.MaxBoltLength.Value--;

                            lw.WriteFullline("   Update forced of:  " + myBoltConn.Name.ToUpper());
                        }
                        catch (Exception)
                        {
                            // Not a Universal Bolt Connection
                        }
                    }

                    // Update AFEM to realize all Universal Bolt Connections
                    myFEM.BaseFEModel.UpdateFemodel();
                    lw.WriteFullline("   UPDATED:  " + myFEM.Name.ToUpper());
                }

                // Set initial working object to working again
                // -------------------------------------------
                theSession.Parts.SetWork(currWork);

                //lw.Open();
            }
            catch (Exception e)
            {
                lw.WriteFullline("!ERROR occurred: " + Environment.NewLine +
                    e.ToString());

                lw.Open();
            }
        }

        /// <summary>
        /// Process target AFEM object and optionally loop through its child components
        /// </summary>
        /// <param name="myAFEM">Target AFEM object</param>
        private static void ProcessFromAFEM(NXOpen.CAE.AssyFemPart myAFEM)
        {
            lw.WriteFullline(Environment.NewLine +
                "AFEM : " + myAFEM.Name.ToString());

            allAFEMs.Add(myAFEM);

            // CREATE SELECTION RECIPES
            CreateSelectionRecipes(myAFEM);

            // CREATE UNIVERSAL BOLT CONNECTION DEFINITIONS
            CreateUniversalBoltConnections(myAFEM, null);


            if (ProcessAll)
            {
                // Cycle through all underlying FEM/AFEM objects and act appropriately
                NXOpen.Assemblies.Component myRoot = myAFEM.ComponentAssembly.RootComponent;

                if (myRoot != null)
                {
                    ProcessChildrenAFEM(myRoot);
                }
            }
        }
        

        /// <summary>
        /// Loop through child components of a target Assembly object and perform an action based on the child component's object type
        /// </summary>
        /// <param name="myComp">Target Assembly object</param>
        private static void ProcessChildrenAFEM(NXOpen.Assemblies.Component myComp)
        {
            //lw.WriteFullline(Environment.NewLine +
            //    "Processing children of AFEM : " + myComp.Name.ToString());

            try
            {
                // Loop through all Child components of AFEM object
                foreach (NXOpen.Assemblies.Component myChild in myComp.GetChildren())
                {
                    //lw.WriteFullline("CHILD : " + myChild.Name.ToString());

                    // Get OwningPart object of Child component
                    NXOpen.BasePart myBasePart = myChild.Prototype.OwningPart;

                    Type TargType = myBasePart.GetType();

                    if (TargType != null)
                    {
                        switch (TargType.ToString())
                        {
                            case "NXOpen.CAE.AssyFemPart":
                                //lw.WriteFullline("Recognized as AFEM");

                                ProcessFromAFEM((NXOpen.CAE.AssyFemPart)myBasePart);
                                //ProcessChildrenAFEM(myChild);
                                break;

                            case "NXOpen.CAE.FemPart":
                                //lw.WriteFullline("Recognized as FEM");
                                ProcessFromFEM((NXOpen.CAE.FemPart)myBasePart);
                                break;

                            default:
                                //lw.WriteFullline("Recognized as " + TargType.ToString() + " -> SKIPPED");
                                break;
                        }
                    }
                }
            }
            catch (Exception e)
            {
                lw.WriteFullline("!ERROR - while processing children of AFEM : " + Environment.NewLine +
                    e.ToString());
            }
        }


        /// <summary>
        /// Process a FEM object
        /// </summary>
        /// <param name="myFEM"></param>
        private static void ProcessFromFEM(NXOpen.CAE.FemPart myFEM)
        {
            lw.WriteFullline(Environment.NewLine +
                "FEM : " + myFEM.Name.ToString());

            allFEMs.Add(myFEM);

            // CREATE SELECTION RECIPES
            CreateSelectionRecipes(myFEM);

            // CREATE UNIVERSAL BOLT CONNECTION DEFINITIONS
            CreateUniversalBoltConnections(null, myFEM);
        }


        /// <summary>
        /// Create predefined Selection Recipes
        /// </summary>
        /// <param name="myAFEM">Target AFEM object</param>
        private static void CreateSelectionRecipes(NXOpen.CAE.CaePart myAFEM)
        {
            try
            {
                lw.WriteFullline(Environment.NewLine +
                "   CREATE SELECTION RECIPES" + Environment.NewLine);

                theSession.Parts.SetWork((NXOpen.BasePart)myAFEM);

                // Create "Get all meshes" Selection Recipe
                // ----------------------------------------
                // Get target name
                string targSelRecipeName = "Get all meshes";

                // Check if not existing yet
                foreach (NXOpen.CAE.SelectionRecipe selRecipe in myAFEM.SelectionRecipes)
                {
                    if (selRecipe.Name.ToUpper() == targSelRecipeName.ToUpper())
                    {
                        lw.WriteFullline("      Selection Recipe:  " + targSelRecipeName.ToUpper() + "  --> exists (skipped)");
                        goto otherRecipes;
                    }
                }

                //lw.WriteFullline("   Selection Recipe:  " + targSelRecipeName.ToUpper());

                // Set target entity types
                NXOpen.CAE.CaeSetGroupFilterType[] entitytypes = new NXOpen.CAE.CaeSetGroupFilterType[1];
                entitytypes[0] = NXOpen.CAE.CaeSetGroupFilterType.CaeMesh;

                // Set points for Bounding Box
                double box_offset = 999999;
                NXOpen.Point leftPoint = myAFEM.Points.CreatePoint(new Point3d(-box_offset, -box_offset, -box_offset));
                NXOpen.Point rightPoint = myAFEM.Points.CreatePoint(new Point3d(box_offset, box_offset, box_offset));

                //// Create Selection Recipe (NX1899)
                //NXOpen.CAE.SelRecipeBuilder selRecipeBuilder = myAFEM.SelectionRecipes.CreateSelRecipeBuilder();
                //NXOpen.CAE.SelRecipeBoundingVolumeStrategy selRecipeBoundingVolumeStrategy = selRecipeBuilder.AddBoxBoundingVolumeStrategy(leftPoint, rightPoint, entitytypes, NXOpen.CAE.SelRecipeBuilder.InputFilterType.EntireModel, null);
                //selRecipeBoundingVolumeStrategy.BoundingVolume.Containment = NXOpen.CAE.CaeBoundingVolumePrimitiveContainment.Inside;

                //selRecipeBuilder.RecipeName = "Get all Meshes";
                //selRecipeBuilder.Commit();


                // Create Seletion Recipe (NX12)
                NXOpen.CAE.BoundingVolumeSelectionRecipe SelRec_GetAllMeshes;
                SelRec_GetAllMeshes = myAFEM.SelectionRecipes.CreateBoxBoundingVolumeRecipe("Get all meshes", leftPoint, rightPoint, entitytypes);
                SelRec_GetAllMeshes.BoundingVolume.Containment = NXOpen.CAE.CaeBoundingVolumePrimitiveContainment.Inside;

                lw.WriteFullline("      Selection Recipe:  " + targSelRecipeName.ToUpper() + "  --> CREATED");

            otherRecipes:;
                // Bolt Curve-related Selection Recipes
                // ------------------------------------
                foreach (MODELS.BoltDefinition boltDefinition in allBoltDefinitions)
                {
                    // Get target name
                    targSelRecipeName = boltDefinition.Name + " Curves";

                    // Check if not existing yet
                    foreach (NXOpen.CAE.SelectionRecipe selRecipe in myAFEM.SelectionRecipes)
                    {
                        if (selRecipe.Name.ToUpper() == targSelRecipeName.ToUpper())
                        {
                            lw.WriteFullline("      Selection Recipe:  " + targSelRecipeName.ToUpper() + "  --> skipped (exists)");
                            goto nextBoltDef;
                        }
                    }

                    // Create Seletion Recipe (NX12)
                    NXOpen.CAE.AttributeSelectionRecipe myAttributeSelRecipe = myAFEM.SelectionRecipes.CreateAttributeRecipe(
                        targSelRecipeName,
                        NXOpen.CAE.CaeSetGroupFilterType.GeomCurve,
                        false,
                        (NXOpen.CAE.CaeSetGroupFilterType)(-1));

                    myAttributeSelRecipe.SetUserAttributes(true, "Curve_" + boltDefinition.Name, false, 0, new string[0], new NXObject.AttributeInformation[0], new NXObject.AttributeInformation[0]);

                    //lw.WriteFullline("   Selection Recipe:  " + targSelRecipeName.ToUpper() + "  --> created");

                    // If Selection Recipe contains 0 entities, delete again
                    if (myAttributeSelRecipe.GetEntities().Length == 0)
                    {
                        theSession.UpdateManager.AddObjectsToDeleteList(new TaggedObject[] { myAttributeSelRecipe });
                        theSession.UpdateManager.DoUpdate(new Session.UndoMarkId());
                        lw.WriteFullline("      Selection Recipe:  " + targSelRecipeName.ToUpper() + "  --> skipped (0 entities)");
                    }
                    else
                    {
                        lw.WriteFullline("      Selection Recipe:  " + targSelRecipeName.ToUpper() + "  --> created (" + myAttributeSelRecipe.GetEntities().Length.ToString() + " entities)");
                    }

                nextBoltDef:;
                }
            }
            catch (Exception e)
            {
                lw.WriteFullline("!ERROR occurred: " + Environment.NewLine +
                    e.ToString());
            }
        }


        /// <summary>
        /// Create predefined Universal Bolt Connection definitions
        /// </summary>
        /// <param name="myAFEM">Target AFEM object</param>
        private static void CreateUniversalBoltConnections(NXOpen.CAE.AssyFemPart myAFEM, NXOpen.CAE.FemPart myFEM)
        {
            try
            {
                lw.WriteFullline(Environment.NewLine +
                "   CREATE UNIVERSAL BOLT CONNECTIONS" + Environment.NewLine);

                // Check whether input is an AFEM or not
                bool isAFEM = myAFEM != null ? true : false;

                // Set target AFEM to working
                // --------------------------
                if (isAFEM) { theSession.Parts.SetWork((NXOpen.BasePart)myAFEM); }
                else { theSession.Parts.SetWork((NXOpen.BasePart)myFEM); }
                

                // Initializations
                // ---------------
                List<NXOpen.CAE.Connections.IConnection> newBoltConnections = new List<NXOpen.CAE.Connections.IConnection>();

                // Get existing Universal Connections
                // ----------------------------------
                List<string> existingUnivConnNames = isAFEM 
                    ? myAFEM.BaseFEModel.ConnectionsContainer.GetAllConnections().Select(x => x.Name).ToList()
                    : myFEM.BaseFEModel.ConnectionsContainer.GetAllConnections().Select(x => x.Name).ToList();

                // Loop through all predefined Bolt Definitions
                foreach (MODELS.BoltDefinition boltDefinition in allBoltDefinitions)
                {
                    try
                    {
                        // Check if not existing yet
                        if (existingUnivConnNames.Contains(boltDefinition.Name.ToUpper()))
                        {
                            // Delete existing Bolt Connection, to make sure:
                            // - it can be adapted if the desired properties are different
                            // - it can be re-generated, but only if the related Selection Recipe has entities in it
                            theSession.UpdateManager.AddObjectsToDeleteList(new TaggedObject[] {
                                myAFEM.BaseFEModel.ConnectionsContainer.GetAllConnections().ToList().Single(x => x.Name == boltDefinition.Name)
                            });
                            theSession.UpdateManager.DoUpdate(new Session.UndoMarkId());
                            //continue;

                            lw.WriteFullline("      BOLT DEFINITION:  " + boltDefinition.Name.ToUpper() + "  --> exists (deleted for re-creation)");
                        }
                        else
                        {
                            lw.WriteFullline("      BOLT DEFINITION:  " + boltDefinition.Name.ToUpper());
                        }


                        // Check if target Selection Recipe exists and contains any Curves at all
                        // ----------------------------------------------------------------------
                        NXOpen.CAE.SelectionRecipe targSelRecipe = null;

                        try
                        {
                            targSelRecipe = isAFEM
                                ? myAFEM.SelectionRecipes.ToArray().Single(x => x.Name.ToUpper().Contains(boltDefinition.Name.ToUpper()) && !x.Name.ToLower().Contains("unique"))
                                : myFEM.SelectionRecipes.ToArray().Single(x => x.Name.ToUpper().Contains(boltDefinition.Name.ToUpper()) && !x.Name.ToLower().Contains("unique"));
                        }
                        catch (Exception e)
                        {
                            lw.WriteFullline("         Related Selection Recipe not found --> skipped");
                            continue;
                        }

                        if (targSelRecipe.GetEntities().Length == 0)
                        {
                            lw.WriteFullline("         Related Selection Recipe (" + targSelRecipe.Name + ") :  " + targSelRecipe.GetEntities().Length.ToString() + " entitities " +
                                "--> skipped");
                            continue;
                        }
                        else
                        {
                            lw.WriteFullline("         Related Selection Recipe (" + targSelRecipe.Name + ") :  " + targSelRecipe.GetEntities().Length.ToString() + " entitities ");
                        }

                        // Create Universal Bolt Connection definition (NX12)
                        // --------------------------------------------------
                        NXOpen.CAE.Connections.Bolt newBoltConn = isAFEM
                            ? (NXOpen.CAE.Connections.Bolt)myAFEM.BaseFEModel.ConnectionsContainer.CreateConnection(NXOpen.CAE.Connections.ConnectionType.Bolt, boltDefinition.Name)
                            : (NXOpen.CAE.Connections.Bolt)myFEM.BaseFEModel.ConnectionsContainer.CreateConnection(NXOpen.CAE.Connections.ConnectionType.Bolt, boltDefinition.Name);

                        // Set Name
                        newBoltConn.SetName(boltDefinition.Name);
                        lw.WriteFullline("         Name            : " + boltDefinition.Name);

                        // Set Targets (Flanges)
                        if (isAFEM) 
                        {
                            newBoltConn.AddFlangeEntities(0, myAFEM.SelectionRecipes.ToArray().Single(x => x.Name.ToUpper() == "GET ALL MESHES").GetEntities());
                            lw.WriteFullline("         Targets         : " + myAFEM.SelectionRecipes.ToArray().Single(x => x.Name.ToUpper() == "GET ALL MESHES").Name + " (Selection Recipe)");
                        }
                        else 
                        {
                            newBoltConn.AddFlangeEntities(0, myFEM.SelectionRecipes.ToArray().Single(x => x.Name.ToUpper() == "GET ALL MESHES").GetEntities());
                            lw.WriteFullline("         Targets         : " + myFEM.SelectionRecipes.ToArray().Single(x => x.Name.ToUpper() == "GET ALL MESHES").Name + " (Selection Recipe)");
                        }
                        

                        // Set Locations
                        newBoltConn.AddLocationSelectionRecipe(0, targSelRecipe);
                        lw.WriteFullline("         Locations       : " + targSelRecipe.Name + " (Selection Recipe)");

                        // Set Physicals
                        newBoltConn.DiameterType = NXOpen.CAE.Connections.DiameterType.User;
                        newBoltConn.Diameter.Value = boltDefinition.ShankDiam;
                        newBoltConn.HeadDiameterType = NXOpen.CAE.Connections.HeadDiameterType.User;
                        newBoltConn.HeadDiameter.Value = boltDefinition.HeadDiam;
                        newBoltConn.MaxBoltLength.Value = boltDefinition.MaxConnLength;

                        lw.WriteFullline("         Shank Diameter  : " + newBoltConn.Diameter.Value.ToString());
                        lw.WriteFullline("         Head Diameter   : " + newBoltConn.HeadDiameter.Value.ToString());
                        lw.WriteFullline("         Max Bolt Length : " + newBoltConn.MaxBoltLength.Value.ToString());

                        // Set Material
                        NXOpen.PhysicalMaterial targMaterial = null;
                        bool isPhysicalMaterial = isAFEM
                            ? myAFEM.MaterialManager.PhysicalMaterials.ToArray().Select(x => x.Name).Contains(boltDefinition.MaterialName)
                            : myFEM.MaterialManager.PhysicalMaterials.ToArray().Select(x => x.Name).Contains(boltDefinition.MaterialName);
                        
                        if (isPhysicalMaterial)
                        {
                            targMaterial = isAFEM
                                ? (NXOpen.PhysicalMaterial)myAFEM.MaterialManager.PhysicalMaterials.ToArray().Single(x => x.Name.ToUpper() == boltDefinition.MaterialName.ToUpper())
                                : (NXOpen.PhysicalMaterial)myFEM.MaterialManager.PhysicalMaterials.ToArray().Single(x => x.Name.ToUpper() == boltDefinition.MaterialName.ToUpper());
                        }
                        else
                        {
                            try
                            {
                                targMaterial = isAFEM
                                    ? myAFEM.MaterialManager.PhysicalMaterials.LoadFromNxmatmllibrary(boltDefinition.MaterialName)
                                    : myFEM.MaterialManager.PhysicalMaterials.LoadFromNxmatmllibrary(boltDefinition.MaterialName);
                                goto matfound;
                            }
                            catch (Exception) { }
                            try
                            {
                                targMaterial = isAFEM
                                    ? myAFEM.MaterialManager.PhysicalMaterials.LoadFromLegacynxlibrary(boltDefinition.MaterialName)
                                    : myFEM.MaterialManager.PhysicalMaterials.LoadFromLegacynxlibrary(boltDefinition.MaterialName);
                                goto matfound;
                            }
                            catch (Exception) { }
                            try
                            {
                                targMaterial = isAFEM
                                    ? myAFEM.MaterialManager.PhysicalMaterials.LoadFromNxlibrary(boltDefinition.MaterialName)
                                    : myFEM.MaterialManager.PhysicalMaterials.LoadFromNxlibrary(boltDefinition.MaterialName);
                                goto matfound;
                            }
                            catch (Exception) { }

                            lw.WriteFullline("         ! MATERIAL COULD NOT BE FOUND :  " + boltDefinition.MaterialName);
                            continue;
                        }


                    matfound:;
                        newBoltConn.Material = targMaterial;
                        lw.WriteFullline("         Material        : " + targMaterial.Name);

                        lw.WriteFullline("         CREATE  =  success");


                        // Add newly created Bolt Connection to list for realization step at end
                        newBoltConnections.Add(newBoltConn);
                    }
                    catch (Exception er)
                    {
                        lw.WriteFullline("!ERROR occurred while creating Bolt Connection: " + Environment.NewLine +
                        er.ToString());
                    }
                }

                // Realize all new Universal Bolt Connection definitions
                // -----------------------------------------------------
                lw.WriteFullline(Environment.NewLine +
                    "   REALIZE UNIVERSAL BOLT CONNECTIONS");

                NXOpen.CAE.Connections.Element boltConnElement = isAFEM
                    ? myAFEM.BaseFEModel.ConnectionElementCollection.Create(NXOpen.CAE.Connections.ElementType.E1DSpider, "Element - BOLT DEFINITIONS", newBoltConnections.ToArray())
                    : myFEM.BaseFEModel.ConnectionElementCollection.Create(NXOpen.CAE.Connections.ElementType.E1DSpider, "Element - BOLT DEFINITIONS", newBoltConnections.ToArray());

                boltConnElement.GenerateElements();
                lw.WriteFullline("      Elements generated");

                //if (myAFEM.BaseFEModel.AskUpdatePending())
                //{
                //    myAFEM.BaseFEModel.UpdateFemodel();
                //    lw.WriteFullline("      FEModel updated");
                //}

                lw.WriteFullline("      REALIZATION =  success");
            }
            catch (Exception e)
            {
                lw.WriteFullline("!ERROR occurred: " + Environment.NewLine +
                    e.ToString());
            }
        }
        #endregion
    }
}
