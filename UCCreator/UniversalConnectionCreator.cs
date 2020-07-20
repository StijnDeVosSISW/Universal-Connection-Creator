//=================================================================================
//        
//        UNIVERSAL CONNECTION CREATOR
//        
//        Description:
//        ------------
//        NXOpen application that creates pre-defined Universal Bolt Connections, 
//        based on the presence of dedicated Curve objects representing the bolts.
//
//        Created by: Stijn De Vos (stijn.de_vos@siemens.com)
//              Version: NX 12.0.2
//              Date: 07-13-2020  (Format: mm-dd-yyyy)
//
//=================================================================================

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
        // Initializations
        private static Session theSession = null;
        private static UI theUI = null;
        private static NXOpen.UF.UFSession theUfSession = null;
        private static ListingWindow lw = null;

        // GUI variables
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
        private NXOpen.BlockStyler.FileSelection nativeFileBrowser0;// Block type: NativeFileBrowser'
        private NXOpen.BlockStyler.Button button_IMPORT;// Block type: Button
        private NXOpen.BlockStyler.Enumeration enum0;// Block type: Enumeration
        private NXOpen.BlockStyler.Group group2;// Block type: Group
        private NXOpen.BlockStyler.Toggle toggle_Validator;// Block type: Toggle
        private NXOpen.BlockStyler.Button button_CREATE;// Block type: Button
        private NXOpen.BlockStyler.Separator separator01;// Block type: Separator

        // TreeList variables
        private static List<NXOpen.BlockStyler.Node> allNodes = new List<Node>();
        private static List<MODELS.BoltDefinition> allBoltDefinitions = new List<MODELS.BoltDefinition>();
        private enum MenuID { AddNode, DeleteNode };

        // Application specific variables
        private enum TargEnv { Production, Debug, Siemens };
        private static TargEnv targEnv;

        private static List<NXOpen.NXObject> allTargObjects = new List<NXObject>();
        private static List<NXOpen.NXObject> objectsToUpdate = new List<NXObject>();

        //private static List<NXOpen.Part> underlyingCAD = new List<NXOpen.Part>();
        //private static List<NXOpen.CAE.CaePart> underlyingCAE = new List<NXOpen.CAE.CaePart>();

        private static string StorageFileName = "UCCreator_SavedBoltDefinitions";  // Name of Excel file in which content of Universal Conn Def tree will be stored for later use
        private static string StoragePath_server = null;
        private static string StoragePath_user = null;

        private static bool ProcessAll = true;
        private static bool RunValidatorAfter = false;
        private static string PathToValidatorExe = "";
        private static string targReferenceSet = "";
        private static NXOpen.BasePart currWork = null;

        private enum CurveSearchingMethod { SelectionRecipe, LineOccurrence};
        private static CurveSearchingMethod targCurveSearching;

        // Diagnostic variables
        private static System.Diagnostics.Stopwatch myStopwatch = null;
        private List<double> ExecutionTimes = new List<double>();
        private static string log = "";
        private static string baseMsg = "";

        //------------------------------------------------------------------------------
        //Constructor for NX Styler class
        //------------------------------------------------------------------------------
        public UniversalConnectionCreator()
        {
            try
            {
                theSession = Session.GetSession();
                theUI = UI.GetUI();
                theUfSession = NXOpen.UF.UFSession.GetUFSession();
                lw = theSession.ListingWindow;
                myStopwatch = new System.Diagnostics.Stopwatch();

                // Set path to GUI .dlx file 
                //targEnv = TargEnv.Production;
                //targEnv = TargEnv.Debug;
                targEnv = TargEnv.Siemens;

                // Set Curve Searching method
                targCurveSearching = CurveSearchingMethod.LineOccurrence;

                switch (targEnv)
                {
                    case TargEnv.Production:
                        theDlxFileName = @"D:\NX\CAE\UBC\ABC\UniversalConnectionCreator\UniversalConnectionCreator.dlx";  // IN CPP TC environment as Production tool
                        PathToValidatorExe = @"D:\NX\CAE\UBC\ABC\UniversalConnectionValidator\UCValidator.dll";
                        targReferenceSet = "CAE";
                        break;
                    case TargEnv.Debug:
                        theDlxFileName = @"C:\sdevos\ABC NXOpen applications\ABC applications\application\UniversalConnectionCreator.dlx";  // Debug by Stijn in CPP TC environment
                        PathToValidatorExe = @"C:\sdevos\ABC NXOpen applications\ABC applications\application\UCValidator.dll";
                        targReferenceSet = "CAE";
                        break;
                    case TargEnv.Siemens:
                        theDlxFileName = "UniversalConnectionCreator.dlx";  // In Siemens TC environment
                        PathToValidatorExe = @"D:\3__TEAMCENTER\2_Projects\2_OCE_TCSimRollOut\4_Automatic_Bolt_Connections__Part_Families\UNIVERSAL CONNECTION VALIDATER\INSTALL\UniversalConnectionValidater\application\UCValidator.dll";
                        targReferenceSet = "Entire Part";
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
                group2 = (NXOpen.BlockStyler.Group)theDialog.TopBlock.FindBlock("group2");
                toggle_Validator = (NXOpen.BlockStyler.Toggle)theDialog.TopBlock.FindBlock("toggle_Validator");
                button_CREATE = (NXOpen.BlockStyler.Button)theDialog.TopBlock.FindBlock("button_CREATE");
                separator01 = (NXOpen.BlockStyler.Separator)theDialog.TopBlock.FindBlock("separator01");

                // Initialize storage paths
                StoragePath_server = System.Reflection.Assembly.GetExecutingAssembly().CodeBase;
                StoragePath_server = StoragePath_server.Remove(StoragePath_server.LastIndexOf("/")) + "\\" + StorageFileName + ".txt";
                StoragePath_server = StoragePath_server.Substring(StoragePath_server.IndexOf("/")).Substring(3).Replace("/", "\\");

                StoragePath_user = Environment.GetEnvironmentVariable("USERPROFILE") + "\\AppData\\Local\\UniversalConnectionCreator\\" + StorageFileName + ".txt";

                // Get initial value for "Run Validator after Creator setting"
                RunValidatorAfter = toggle_Validator.Value;


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
                log += 
                " ------------------------------ " + Environment.NewLine +
                " ------------------------------ " + Environment.NewLine +
                "| UNIVERSAL CONNECTION CREATOR |" + Environment.NewLine +
                " ------------------------------ " + Environment.NewLine +
                " ------------------------------ " + Environment.NewLine + Environment.NewLine;

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
                UpdateProcessAll();

                // Check value of RunValidatorAfter
                UpdateRunValidatorAfter();

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
                    UpdateProcessAll();
                }
                else if (block == toggle_Validator)
                {
                    UpdateRunValidatorAfter();
                }
                else if (block == button_CREATE)
                {
                    // Hide all objects
                    HideAllObjects();

                    // Execute Universal Bolt Connection creation
                    ExecuteBoltGeneration();

                    // Store current Tree List content to use in next session
                    StoreUnivConnList();

                    // Show all objects
                    ShowAllObjects();

                    // Show log
                    if (lw.IsOpen) { lw.Close(); }
                    lw.Open();
                    lw.WriteFullline(log);

                    // Run VALIDATOR, if selected
                    RunValidator();
                }
                else if (block == separator01)
                {
                    //---------Enter your code here-----------
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
        #region TREELIST CALLBACKS
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
        #endregion


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
        /// Writes the defined message in the NX status bar (at the bottom)
        /// </summary>
        /// <param name="msg"></param>
        private static void SetNXstatusMessage(string msg)
        {
            //theUfSession.Ui.SetStatus(msg);
            theUfSession.Ui.SetPrompt("ABC CREATOR | " + msg);
        }

        /// <summary>
        /// Write to log content
        /// </summary>
        /// <param name="msg"></param>
        private void WriteToLog(string msg)
        {
            log += msg + Environment.NewLine;
        }

        /// <summary>
        /// Hides all objects in the NX/SC session
        /// </summary>
        private void HideAllObjects()
        {
            theSession.DisplayManager.HideByType("SHOW_HIDE_TYPE_ALL", NXOpen.DisplayManager.ShowHideScope.AnyInAssembly);
            theSession.UpdateManager.DoUpdate(new Session.UndoMarkId());
        }

        /// <summary>
        /// Shows all objects in the NX/SC session
        /// </summary>
        private void ShowAllObjects()
        {
            theSession.DisplayManager.ShowByType("SHOW_HIDE_TYPE_ALL", NXOpen.DisplayManager.ShowHideScope.AnyInAssembly);
            theSession.UpdateManager.DoUpdate(new Session.UndoMarkId());
        }

        /// <summary>
        /// Import predefined Bolt Definitions from an Excel file
        /// </summary>
        /// <param name="filePath">Full path to target Excel file</param>
        private void ImportDefsFromExcel(string filePath)
        {
            try
            {
                log += Environment.NewLine +
                    " ----------------------------------------- " + Environment.NewLine +
                    "| IMPORT BOLT DEFINITIONS FROM EXCEL FILE |" + Environment.NewLine +
                    " ----------------------------------------- " + Environment.NewLine;

                log += "Input Excel file  :  " + filePath + Environment.NewLine;

                // Clear all existing nodes and BoltDefinition objects
                foreach (NXOpen.BlockStyler.Node myNode in allNodes)
                {
                    tree_control0.DeleteNode(myNode);
                }
                allNodes.Clear();

                log += Environment.NewLine + "Delete existing Bolt Definitions :  SUCCESS" + Environment.NewLine;

                // Create COM objects to use Excel to read the input Excel file
                Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(filePath);
                Microsoft.Office.Interop.Excel.Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                Microsoft.Office.Interop.Excel.Range targRange = xlWorksheet.UsedRange;

                // Extract each BoltDefinition object from Used Range of Excel sheet
                // -> start at i = 2, because first row in Excel sheet are the column headers! (Excel is NOT zero-based)
                for (int i = 2; i < (targRange.Rows.Count+1); i++)
                {
                    log += Environment.NewLine + "IMPORTING: Bolt Definition " + (i - 1).ToString() + Environment.NewLine;
                    // Add new node to Tree List
                    NXOpen.BlockStyler.Node newNode = tree_control0.CreateNode("<new>");
                    tree_control0.InsertNode(newNode, null, null, Tree.NodeInsertOption.Last);

                    newNode.SetColumnDisplayText(0, (string)targRange.Cells[i, 1].Text);
                    newNode.SetColumnDisplayText(1, (string)targRange.Cells[i, 2].Text);
                    newNode.SetColumnDisplayText(2, (string)targRange.Cells[i, 3].Text);
                    newNode.SetColumnDisplayText(3, (string)targRange.Cells[i, 4].Text);
                    newNode.SetColumnDisplayText(4, (string)targRange.Cells[i, 5].Text);

                    allNodes.Add(newNode);

                    log += "   New node created with:" + Environment.NewLine +
                        "   Name                           = " + newNode.GetColumnDisplayText(0) + Environment.NewLine +
                        "   Shank Diameter [mm]            = " + newNode.GetColumnDisplayText(1) + Environment.NewLine +
                        "   Head Diameter [mm]             = " + newNode.GetColumnDisplayText(2) + Environment.NewLine +
                        "   Maximum Connection Length [mm] = " + newNode.GetColumnDisplayText(3) + Environment.NewLine +
                        "   Material Name                  = " + newNode.GetColumnDisplayText(4) + Environment.NewLine + Environment.NewLine;

                    //for (int j = 1; j < (targRange.Columns.Count+1); j++)
                    //{
                    //    log += "  Cell[" + i.ToString() + "," + j.ToString() + "] = " + targRange.Cells[i, j].Value);
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

                log += "Released: targRange, xlWorksheet, xlWorkbook, xlApp" + Environment.NewLine;
            }
            catch (Exception e)
            {
                log += "!ERROR occurred: " + Environment.NewLine +
                    e.ToString() + Environment.NewLine;
            }
        }

        /// <summary>
        /// Import last used Bolt Definitions, stored by the previous session
        /// </summary>
        private void ImportStoredBoltDefinitions(string filePath)
        {
            try
            {
                // Get target file
                // ---------------
                // Check if it exists
                if (!File.Exists(filePath))
                {
                    log += "No stored bolt definitions could be found:  continue..." + Environment.NewLine;
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
                        //log += "   PROCESSING:  " + line + Environment.NewLine;
                        string currLine = line;
                        MODELS.BoltDefinition newBoltDefinition = new MODELS.BoltDefinition();

                        // Name
                        newBoltDefinition.Name = currLine.Remove(currLine.IndexOf("|") - 1);
                        currLine = currLine.Substring(currLine.IndexOf("|") + 2);

                        // Shank Diameter
                        newBoltDefinition.ShankDiam = Convert.ToDouble(currLine.Remove(currLine.IndexOf("|") - 1));
                        currLine = currLine.Substring(currLine.IndexOf("|") + 2);

                        // Head Diameter
                        newBoltDefinition.HeadDiam = Convert.ToDouble(currLine.Remove(currLine.IndexOf("|") - 1));
                        currLine = currLine.Substring(currLine.IndexOf("|") + 2);

                        // Maximum Connection Length
                        newBoltDefinition.MaxConnLength = Convert.ToDouble(currLine.Remove(currLine.IndexOf("|") - 1));
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

                //log += Environment.NewLine + "Delete existing Bolt Definitions :  SUCCESS" + Environment.NewLine;

                // Import new Bolt Definitions
                int i = 1;
                foreach (MODELS.BoltDefinition boltDefinition in currBoltDefinitions)
                {
                    //log += Environment.NewLine + "IMPORTING: Bolt Definition " + i.ToString() + Environment.NewLine;

                    // Add new node to Tree List
                    NXOpen.BlockStyler.Node newNode = tree_control0.CreateNode("<new>");
                    tree_control0.InsertNode(newNode, null, null, Tree.NodeInsertOption.Last);

                    newNode.SetColumnDisplayText(0, boltDefinition.Name);
                    newNode.SetColumnDisplayText(1, boltDefinition.ShankDiam.ToString());
                    newNode.SetColumnDisplayText(2, boltDefinition.HeadDiam.ToString());
                    newNode.SetColumnDisplayText(3, boltDefinition.MaxConnLength.ToString());
                    newNode.SetColumnDisplayText(4, boltDefinition.MaterialName);

                    allNodes.Add(newNode);

                    //log += "   New node created with:" + Environment.NewLine +
                    //    "   Name                           = " + newNode.GetColumnDisplayText(0) + Environment.NewLine +
                    //    "   Shank Diameter [mm]            = " + newNode.GetColumnDisplayText(1) + Environment.NewLine +
                    //    "   Head Diameter [mm]             = " + newNode.GetColumnDisplayText(2) + Environment.NewLine +
                    //    "   Maximum Connection Length [mm] = " + newNode.GetColumnDisplayText(3) + Environment.NewLine +
                    //    "   Material Name                  = " + newNode.GetColumnDisplayText(4) + Environment.NewLine + Environment.NewLine;

                    i++;
                }

                log += "Load stored Bolt Definitions:  SUCCESS" + Environment.NewLine;
            }
            catch (Exception e)
            {
                log += "!ERROR occurred: " + Environment.NewLine +
                    e.ToString() + Environment.NewLine;
            }
        }

        /// <summary>
        /// Store current Universal Bolt Connection definitions in an Excel file
        /// </summary>
        private void StoreUnivConnList()
        {
            try
            {
                log += Environment.NewLine +
                    " ----------------------------------------- " + Environment.NewLine +
                    "| SAVE BOLT DEFINITION LIST FOR LATER USE |" + Environment.NewLine +
                    " ----------------------------------------- " + Environment.NewLine;

                // Get target Excel file path
                string filePath = StoragePath_user;

                //log += "Target file path :  " + filePath);

                // Create new Text file
                if (File.Exists(filePath)) { File.Delete(filePath); }
                //log += "Write content...");

                // Check if target Directory exists, if not, try to create it
                Directory.CreateDirectory(filePath.Remove(filePath.LastIndexOf(@"\")));

                using (StreamWriter sw = File.CreateText(filePath))
                {
                    // Write Headers
                    sw.WriteLine("NAME | SHANK DIAMETER [mm] | HEAD DIAMETER [mm] | MAXIMUM CONNECTION LENGTH [mm] | MATERIAL");
                    //log += "   Added Headers");

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

                        //log += "   Added Node " + i.ToString() + Environment.NewLine;

                        i++;
                    }
                }

                log += "Saved" + Environment.NewLine;
            }
            catch (Exception e)
            {
                log += "!ERROR occurred: " + Environment.NewLine +
                    e.ToString() + Environment.NewLine;
            }
        }

        /// <summary>
        /// Updates value of ProcessAll, based on input coming from GUI
        /// </summary>
        private void UpdateProcessAll()
        {
            switch (enum0.ValueAsString)
            {
                case "All assyFEM levels":
                    ProcessAll = true;
                    //theUI.NXMessageBox.Show("Selection of target (A)FEM level changed:", NXMessageBox.DialogType.Information, "Set to:  All assyFEM levels");
                    break;

                case "This level only":
                    ProcessAll = false;
                    //theUI.NXMessageBox.Show("Selection of target (A)FEM level changed:", NXMessageBox.DialogType.Information, "Set to:  This level only");
                    break;

                default:
                    //theUI.NXMessageBox.Show("Selection of target (A)FEM level changed:", NXMessageBox.DialogType.Information, "NOT FOUND!" + Environment.NewLine + 
                    //    "Set to:  " + enum0.ValueAsString);
                    break;
            }
        }

        /// <summary>
        /// Updates value of RunValidatorAfter, based on input coming from GUI
        /// </summary>
        private void UpdateRunValidatorAfter()
        {
            RunValidatorAfter = toggle_Validator.Value;
        }

        /// <summary>
        /// Start Bolt generation execution
        /// </summary>
        private void ExecuteBoltGeneration()
        {
            try
            {
                log += Environment.NewLine + Environment.NewLine +
                    " ------------------------- " + Environment.NewLine +
                    "| EXECUTE BOLT GENERATION |" + Environment.NewLine +
                    " ------------------------- " + Environment.NewLine;

                #region CREATE LIST OF PREDEFINED BOLT CONNECTIONS
                // ------------------------------------------
                // CREATE LIST OF PREDEFINED BOLT CONNECTIONS
                // ------------------------------------------
                myStopwatch.Restart();
                log += Environment.NewLine + Environment.NewLine +
                    " -------------------------------------------- " + Environment.NewLine +
                    "| Create list of Predefined Bolt Connections |" + Environment.NewLine +
                    " -------------------------------------------- " + Environment.NewLine;

                log += "Generate list of predefined Bolt Connections..." + Environment.NewLine;

                allBoltDefinitions.Clear();

                foreach (NXOpen.BlockStyler.Node node in allNodes)
                {
                    allBoltDefinitions.Add(new MODELS.BoltDefinition()
                    {
                        Name = node.GetColumnDisplayText(0),
                        ShankDiam = Convert.ToDouble(node.GetColumnDisplayText(1)),
                        HeadDiam = Convert.ToDouble(node.GetColumnDisplayText(2)),
                        MaxConnLength = Convert.ToDouble(node.GetColumnDisplayText(3)),
                        MaterialName = node.GetColumnDisplayText(4)
                    }); ;

                    log += "   Added Bolt Definition:  " + node.GetColumnDisplayText(0).ToUpper() + Environment.NewLine;
                }

                SetNXstatusMessage("Generated list of predefined bolt connections ...");

                myStopwatch.Stop();
                log += Environment.NewLine + 
                    "[" + myStopwatch.Elapsed.TotalSeconds.ToString() + " seconds]" + Environment.NewLine;
                ExecutionTimes.Add(myStopwatch.Elapsed.TotalSeconds);
                #endregion


                #region GATHER ALL (A)FEM OBJECTS TO PROCESS
                // ------------------------------------
                // GATHER ALL (A)FEM OBJECTS TO PROCESS
                // ------------------------------------
                myStopwatch.Restart();
                log += Environment.NewLine +
                    " --------------------------- " + Environment.NewLine +
                    "| Gather all (A)FEM objects |" + Environment.NewLine +
                    " --------------------------- " + Environment.NewLine;

                currWork = theSession.Parts.BaseWork;
                log += "Current working object :  " + currWork.ToString() + Environment.NewLine;

                if (ProcessAll)
                {
                    log += Environment.NewLine +
                        "SCENARIO =   ALL ASSYFEM LEVELS" + Environment.NewLine;
                }
                else
                {
                    log += Environment.NewLine +
                        "SCENARIO =   THIS LEVEL ONLY" + Environment.NewLine;
                }

                log += Environment.NewLine +
                    "Getting all FEM and AFEM objects to process..." + Environment.NewLine;
                SetNXstatusMessage("Getting all FEM and AFEM objects to process...");

                allTargObjects.Clear();
                switch (theSession.Parts.BaseWork.GetType().ToString())
                {
                    case "NXOpen.CAE.SimPart":
                        //log += "---> Recognized as SIM" + Environment.NewLine;
                        NXOpen.CAE.SimPart mySIM = (NXOpen.CAE.SimPart)theSession.Parts.BaseWork;

                        switch (mySIM.FemPart.GetType().ToString())
                        {
                            case "NXOpen.CAE.AssyFemPart":
                                //log += "---> Underlying CAE object = recognized as AFEM" + Environment.NewLine;
                                ProcessFromAFEM((NXOpen.CAE.AssyFemPart)mySIM.FemPart);
                                break;

                            case "NXOpen.CAE.FemPart":
                                //log += "---> Underlying CAE object = recognized as FEM" + Environment.NewLine;
                                ProcessFromFEM((NXOpen.CAE.FemPart)mySIM.FemPart);
                                break;

                            default:
                                //log += "---> Underlying CAE object = recognized as " + mySIM.FemPart.GetType().ToString() + " -> SKIPPED" + Environment.NewLine;
                                break;
                        }

                        break;

                    case "NXOpen.CAE.AssyFemPart":
                        //log += "---> Recognized as AFEM" + Environment.NewLine;
                        ProcessFromAFEM((NXOpen.CAE.AssyFemPart)theSession.Parts.BaseWork);
                        break;

                    case "NXOpen.CAE.FemPart":
                        //log += "---> Recognized as FEM" + Environment.NewLine;
                        ProcessFromFEM((NXOpen.CAE.FemPart)theSession.Parts.BaseWork);
                        break;

                    default:
                        log += "---> Not recognized as SIM, AFEM or FEM, but as:  " + theSession.Parts.BaseWork.GetType().ToString() + Environment.NewLine +
                            "=> exiting..." + Environment.NewLine;
                        return;
                        break;
                }

                log += "   -> # (A)FEM objects to process = " + allTargObjects.Count.ToString() + Environment.NewLine;



                // Remove any duplicates from List
                allTargObjects = allTargObjects.Distinct().ToList();

                log += Environment.NewLine +
                    "Removing duplicates... " + Environment.NewLine +
                     "   -> # (A)FEM objects to process = " + allTargObjects.Count.ToString() + Environment.NewLine;

                foreach (NXObject obj in allTargObjects)
                {
                    log += "   " + obj.Name.ToUpper() + Environment.NewLine;
                }



                myStopwatch.Stop();
                log += Environment.NewLine + 
                    "[" + myStopwatch.Elapsed.TotalSeconds.ToString() + " seconds]" + Environment.NewLine;
                ExecutionTimes.Add(myStopwatch.Elapsed.TotalSeconds);
                #endregion


                #region PROCESS EACH (A)FEM OBJECT
                // --------------------------
                // PROCESS EACH (A)FEM OBJECT
                // --------------------------
                myStopwatch.Restart();
                log += Environment.NewLine +
                    " ---------------------------- " + Environment.NewLine +
                    "| Process each (A)FEM object |" + Environment.NewLine +
                    " ---------------------------- " + Environment.NewLine;
                int i = 1;
                int tot = allTargObjects.Count;
                bool SelRecipesHaveCurves = false;

                System.Diagnostics.Stopwatch detailStopwatch = new System.Diagnostics.Stopwatch();

                foreach (NXObject targObj in allTargObjects)
                {
                    // Initializations
                    List<double> DetailExecutionTimes = new List<double>() { 0, 0, 0, 0 };

                    baseMsg = "Processing unique (A)FEM objects :   " + i.ToString() + @"/" + tot.ToString() + "  (" + Math.Round(((double)i / tot) * 100) + "%)    " +
                        "[" + targObj.Name + "]";
                    SetNXstatusMessage(baseMsg);

                    log += Environment.NewLine +
                        "=================================================================================" + Environment.NewLine +
                        targObj.Name.ToUpper() + Environment.NewLine +
                        "=================================================================================" + Environment.NewLine;

                    // Check whether it is a FEM or an Assembly FEM object
                    bool isFEM = targObj.GetType().ToString() == "NXOpen.CAE.FemPart" ? true : false;

                    if (isFEM) { log += "---> Recognized as FEM" + Environment.NewLine; }
                    else { log += "---> Recognized as AFEM" + Environment.NewLine; }

                    // If target object is a FEM, check if FEMs should be processed or not
                    if (isFEM && ProcessAll)
                    {
                        log += "---> Current run is processing Assembly FEM levels only:   SKIPPED" + Environment.NewLine;
                        continue;
                    }

                    // If MONO-FEM scenario (FEM object and Only This Level)
                    if (isFEM && !ProcessAll)
                    {
                        // Initialize
                        NXOpen.CAE.FemPart myFEM = (NXOpen.CAE.FemPart)targObj;

                        // Replace Reference Set
                        ReplaceReferenceSet(myFEM);

                        // Make sure Geometry Options are set correctly (so that Curve objects are propagated to the FEM level)
                        SetFEMGeometryOptions(myFEM);
                    }

                    // Check if target object has any curves at all for Bolt connections
                    if (!GetAllCurveOccurrences(targObj.Tag).Any(x => x.Name.Contains("CURVE_")))
                    {
                        log += "---> Object does not contain any curve with CURVE_ in its name:  SKIPPED" + Environment.NewLine;
                        continue;
                    }

                    // SET TO WORKING PART
                    // -------------------
                    detailStopwatch.Restart();
                    if (theSession.Parts.BaseWork.Tag != targObj.Tag) { theSession.Parts.SetWork((NXOpen.BasePart)targObj); }

                    detailStopwatch.Stop();
                    DetailExecutionTimes[0] = detailStopwatch.Elapsed.TotalSeconds;


                    // CREATE SELECTION RECIPES
                    // ------------------------
                    // (If a FEM,) FEM should NOT be processed if:
                    // - it does not contain any mesh objects
                    //   => assumed that it represents a Bolt part family member, with just the CAD curve data
                    if (isFEM)
                    {
                        NXOpen.CAE.FemPart myFEM = (NXOpen.CAE.FemPart)targObj;
                        if (myFEM.BaseFEModel.MeshManager.GetMeshes().Length < 1)
                        {
                            log += Environment.NewLine +
                                "===> FEM does not contain any mesh objects:  assumed to be a bolt representation  (-> SKIPPED)" + Environment.NewLine;
                            i++;
                            continue;
                        }
                    }

                    // Finally, create Selection Recipes
                    detailStopwatch.Restart();

                    SelRecipesHaveCurves = CreateSelectionRecipes((NXOpen.CAE.CaePart)targObj);

                    detailStopwatch.Stop();
                    DetailExecutionTimes[1] = detailStopwatch.Elapsed.TotalSeconds;


                    // CREATE UNIVERSAL BOLT CONNECTION DEFINITIONS
                    // --------------------------------------------
                    detailStopwatch.Restart();

                    if (isFEM)
                    {
                        CreateUniversalBoltConnections(null, (NXOpen.CAE.FemPart)targObj);
                    }
                    else
                    {
                        CreateUniversalBoltConnections((NXOpen.CAE.AssyFemPart)targObj, null);
                    }

                    detailStopwatch.Stop();
                    DetailExecutionTimes[2] = detailStopwatch.Elapsed.TotalSeconds;


                    // UPDATE BOLT CONNECTIONS     ---> Moved to end
                    // -----------------------
                    objectsToUpdate.Add(targObj);


                    // Diagnostics
                    log += Environment.NewLine +
                        "DIAGNOSTICS" + Environment.NewLine +
                        "-----------" + Environment.NewLine +
                        "SET TO WORKING OBJECT             =  " + DetailExecutionTimes[0].ToString() + " seconds" + Environment.NewLine +
                        "CREATE SELECTION RECIPES          =  " + DetailExecutionTimes[1].ToString() + " seconds" + Environment.NewLine +
                        "CREATE UNIVERSAL BOLT CONNECTIONS =  " + DetailExecutionTimes[2].ToString() + " seconds" + Environment.NewLine +
                        "UPDATE BOLT CONNECTIONS           =  " + DetailExecutionTimes[3].ToString() + " seconds" + Environment.NewLine;
                        
                    i++;
                }

                myStopwatch.Stop();
                log += Environment.NewLine + 
                    "[" + myStopwatch.Elapsed.TotalSeconds.ToString() + " seconds]" + Environment.NewLine;
                ExecutionTimes.Add(myStopwatch.Elapsed.TotalSeconds);
                #endregion


                #region UPDATE EACH (A)FEM OBJECT
                // -------------------------
                // UPDATE EACH (A)FEM OBJECT
                // -------------------------
                myStopwatch.Restart();
                log += Environment.NewLine +
                    " ------------------------------------------------ " + Environment.NewLine +
                    "| Update (A)FEM objects that have pending update |" + Environment.NewLine +
                    " ------------------------------------------------ " + Environment.NewLine;

                // Reverse order of list of object to update, to make sure that the updating happens bottom-up
                objectsToUpdate.Reverse();

                int j = 1;
                tot = objectsToUpdate.Count;

                // Update each modified object
                foreach (NXObject targObj in objectsToUpdate)
                {
                    SetNXstatusMessage("Updating (A)FEM objects :   " + j.ToString() + @"/" + tot.ToString() + "  (" + Math.Round(((double)j / tot) * 100) + "%)    " +
                        "[" + targObj.Name + "]");

                    UpdateCAEObjectConnections((NXOpen.CAE.BaseFemPart)targObj);

                    j++;
                }

                myStopwatch.Stop();
                log += "[" + myStopwatch.Elapsed.TotalSeconds.ToString() + " seconds]" + Environment.NewLine;
                ExecutionTimes.Add(myStopwatch.Elapsed.TotalSeconds);
                #endregion


                // Set initial working object to working again
                // -------------------------------------------
                theSession.Parts.SetWork(currWork);

                //lw.Open();

                // Diagnostics
                log += Environment.NewLine +
                    "DIAGNOSTICS" + Environment.NewLine +
                    "-----------" + Environment.NewLine +
                    "CREATE LIST OF PREDEFINED BOLT CONNECTIONS =  " + ExecutionTimes[0].ToString() + " seconds" + Environment.NewLine +
                    "GATHER ALL (A)FEM OBJECTS TO PROCESS       =  " + ExecutionTimes[1].ToString() + " seconds" + Environment.NewLine +
                    "PROCESS EACH (A)FEM OBJECT                 =  " + ExecutionTimes[2].ToString() + " seconds" + Environment.NewLine +
                    "UPDATE EACH (A)FEM OBJECT                  =  " + ExecutionTimes[3].ToString() + " seconds" + Environment.NewLine;
            }
            catch (Exception e)
            {
                log += "!ERROR occurred: " + Environment.NewLine +
                    e.ToString() + Environment.NewLine;

                lw.WriteFullline(log);
                lw.Open();
            }
        }

        /// <summary>
        /// Process target AFEM object and optionally loop through its child components
        /// </summary>
        /// <param name="myAFEM">Target AFEM object</param>
        private static void ProcessFromAFEM(NXOpen.CAE.AssyFemPart myAFEM)
        {
            log += "(AFEM) : " + myAFEM.Name.ToString() + Environment.NewLine;

            allTargObjects.Add(myAFEM);

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
            //log += Environment.NewLine +
            //    "Processing children of AFEM : " + myComp.Name.ToString());

            try
            {
                // Loop through all Child components of AFEM object
                foreach (NXOpen.Assemblies.Component myChild in myComp.GetChildren())
                {
                    //log += "CHILD : " + myChild.Name.ToString());

                    // Get OwningPart object of Child component
                    NXOpen.BasePart myBasePart = myChild.Prototype.OwningPart;

                    Type TargType = myBasePart.GetType();

                    if (TargType != null)
                    {
                        switch (TargType.ToString())
                        {
                            case "NXOpen.CAE.AssyFemPart":
                                //log += "Recognized as AFEM" + Environment.NewLine;

                                ProcessFromAFEM((NXOpen.CAE.AssyFemPart)myBasePart);
                                //ProcessChildrenAFEM(myChild);
                                break;

                            case "NXOpen.CAE.FemPart":
                                //log += "Recognized as FEM" + Environment.NewLine;
                                ProcessFromFEM((NXOpen.CAE.FemPart)myBasePart);
                                break;

                            default:
                                //log += "Recognized as " + TargType.ToString() + " -> SKIPPED" + Environment.NewLine;
                                break;
                        }
                    }
                }
            }
            catch (Exception e)
            {
                log += "!ERROR - while processing children of AFEM : " + Environment.NewLine +
                    e.ToString() + Environment.NewLine;
            }
        }


        /// <summary>
        /// Process a FEM object
        /// </summary>
        /// <param name="myFEM"></param>
        private static void ProcessFromFEM(NXOpen.CAE.FemPart myFEM)
        {
            log += "(FEM)  : " + myFEM.Name.ToString() + Environment.NewLine; 

            allTargObjects.Add(myFEM);
        }


        /// <summary>
        /// Create predefined Selection Recipes
        /// </summary>
        /// <param name="myAFEM">Target AFEM object</param>
        private static bool CreateSelectionRecipes(NXOpen.CAE.CaePart myCAEPart)
        {
            bool CreatedCurveSelRecipe = false;

            try
            {
                log += Environment.NewLine +
                "   CREATE SELECTION RECIPES" + Environment.NewLine +
                "   ------------------------" + Environment.NewLine;

                // Set Working, if needed
                if (theSession.Parts.BaseWork.Tag != myCAEPart.Tag) { theSession.Parts.SetWork((NXOpen.BasePart)myCAEPart); }


                // Create "Get all meshes" Selection Recipe
                // ----------------------------------------
                // Get target name
                string targSelRecipeName = "Get all meshes";

                //Check if not existing yet
                if (myCAEPart.SelectionRecipes.ToArray().Select(x => x.Name.ToUpper()).Contains(targSelRecipeName.ToUpper()))
                {
                    log += "      Selection Recipe:  " + targSelRecipeName.ToUpper() + "  --> exists (skipped)" + Environment.NewLine;
                    goto otherRecipes;
                }

                //log += "   Selection Recipe:  " + targSelRecipeName.ToUpper());

                // Set target entity types
                NXOpen.CAE.CaeSetGroupFilterType[] entitytypes = new NXOpen.CAE.CaeSetGroupFilterType[1];
                entitytypes[0] = NXOpen.CAE.CaeSetGroupFilterType.CaeMesh;

                // Set points for Bounding Box
                double box_offset = 999999;
                NXOpen.Point leftPoint = myCAEPart.Points.CreatePoint(new Point3d(-box_offset, -box_offset, -box_offset));
                NXOpen.Point rightPoint = myCAEPart.Points.CreatePoint(new Point3d(box_offset, box_offset, box_offset));

                //// Create Selection Recipe (NX1899)
                //NXOpen.CAE.SelRecipeBuilder selRecipeBuilder = myAFEM.SelectionRecipes.CreateSelRecipeBuilder();
                //NXOpen.CAE.SelRecipeBoundingVolumeStrategy selRecipeBoundingVolumeStrategy = selRecipeBuilder.AddBoxBoundingVolumeStrategy(leftPoint, rightPoint, entitytypes, NXOpen.CAE.SelRecipeBuilder.InputFilterType.EntireModel, null);
                //selRecipeBoundingVolumeStrategy.BoundingVolume.Containment = NXOpen.CAE.CaeBoundingVolumePrimitiveContainment.Inside;

                //selRecipeBuilder.RecipeName = "Get all Meshes";
                //selRecipeBuilder.Commit();


                // Create Seletion Recipe (NX12)
                NXOpen.CAE.BoundingVolumeSelectionRecipe SelRec_GetAllMeshes;
                SelRec_GetAllMeshes = myCAEPart.SelectionRecipes.CreateBoxBoundingVolumeRecipe("Get all meshes", leftPoint, rightPoint, entitytypes);
                SelRec_GetAllMeshes.BoundingVolume.Containment = NXOpen.CAE.CaeBoundingVolumePrimitiveContainment.Inside;

                log += "      Selection Recipe:  " + targSelRecipeName.ToUpper() + "  --> CREATED" + Environment.NewLine;

            otherRecipes:;

                if (targCurveSearching == CurveSearchingMethod.SelectionRecipe)
                {
                    // Bolt Curve-related Selection Recipes
                    // ------------------------------------
                    // Get all available curves on current object level
                    List<NXOpen.Line> availCurves = GetAllCurveOccurrences(myCAEPart.Tag);

                    // Create each curve-related Selection Recipe, for which the corresponding curve is available
                    foreach (MODELS.BoltDefinition boltDefinition in allBoltDefinitions)
                    {
                        // Get target name
                        targSelRecipeName = boltDefinition.Name + " Curves";
                        string targCurveName = "Curve_" + boltDefinition.Name;

                        // Check if target Curve is available at this object level
                        if (!availCurves.Select(x => x.Name.ToUpper()).Contains(targCurveName.ToUpper()))
                        {
                            log += "      Target Curve (" + targCurveName + ") is not present at this object level  --> skipped (unavailable)" + Environment.NewLine;
                            continue;
                        }

                        // Check if not existing yet
                        if (myCAEPart.SelectionRecipes.ToArray().Select(x => x.Name.ToUpper()).Contains(targSelRecipeName.ToUpper()))
                        {
                            log += "      Selection Recipe:  " + targSelRecipeName.ToUpper() + "  --> skipped (exists)" + Environment.NewLine;
                            CreatedCurveSelRecipe = true;
                            continue;
                        }

                        // Create Seletion Recipe (NX12)
                        NXOpen.CAE.AttributeSelectionRecipe myAttributeSelRecipe = myCAEPart.SelectionRecipes.CreateAttributeRecipe(
                            targSelRecipeName,
                            NXOpen.CAE.CaeSetGroupFilterType.GeomCurve,
                            false,
                            (NXOpen.CAE.CaeSetGroupFilterType)(-1));

                        myAttributeSelRecipe.SetUserAttributes(true, targCurveName, false, 0, new string[0], new NXObject.AttributeInformation[0], new NXObject.AttributeInformation[0]);

                        //log += "   Selection Recipe:  " + targSelRecipeName.ToUpper() + "  --> created" + Environment.NewLine;

                        // If Selection Recipe contains 0 entities, mark for deleting
                        if (myAttributeSelRecipe.GetEntities().Length == 0)
                        {
                            theSession.UpdateManager.AddObjectsToDeleteList(new TaggedObject[] { myAttributeSelRecipe });
                            log += "      Selection Recipe:  " + targSelRecipeName.ToUpper() + "  --> skipped (0 entities)" + Environment.NewLine;
                        }
                        else
                        {
                            log += "      Selection Recipe:  " + targSelRecipeName.ToUpper() + "  --> created (" + myAttributeSelRecipe.GetEntities().Length.ToString() + " entities)" + Environment.NewLine;
                            CreatedCurveSelRecipe = true;
                        }

                        //nextBoltDef:;
                    }

                    // Delete all empty Selection Recipes
                    theSession.UpdateManager.DoUpdate(new Session.UndoMarkId());
                    log += "      Deleted all skipped Selection Recipes again" + Environment.NewLine;
                }
            }
            catch (Exception e)
            {
                log += "!ERROR occurred: " + Environment.NewLine +
                    e.ToString() + Environment.NewLine;
            }

            return CreatedCurveSelRecipe;
        }


        /// <summary>
        /// Create predefined Universal Bolt Connection definitions
        /// </summary>
        /// <param name="myAFEM">Target AFEM object</param>
        private static void CreateUniversalBoltConnections(NXOpen.CAE.AssyFemPart myAFEM, NXOpen.CAE.FemPart myFEM)
        {
            try
            {
                switch (targCurveSearching)
                {
                    case CurveSearchingMethod.SelectionRecipe:
                        log += Environment.NewLine +
                        "   CREATE UNIVERSAL BOLT CONNECTIONS      [SELECTION RECIPE BASED]" + Environment.NewLine +
                        "   ---------------------------------" + Environment.NewLine;
                        break;
                    case CurveSearchingMethod.LineOccurrence:
                        log += Environment.NewLine +
                        "   CREATE UNIVERSAL BOLT CONNECTIONS      [LINE OCCURRENCE BASED]" + Environment.NewLine +
                        "   ---------------------------------" + Environment.NewLine;
                        break;

                    default:
                        break;
                }
                

                // Check whether input is an AFEM or not
                bool isAFEM = myAFEM != null ? true : false;

                NXOpen.Tag workObjTag = theSession.Parts.BaseWork.Tag;

                NXOpen.CAE.CaePart targCAEPart = null;
                if (isAFEM) { targCAEPart = (NXOpen.CAE.CaePart)myAFEM; }
                else { targCAEPart = (NXOpen.CAE.CaePart)myFEM; }

                // Set target AFEM to working
                // --------------------------
                if (isAFEM) { if (workObjTag != myAFEM.Tag) { theSession.Parts.SetWork((NXOpen.BasePart)myAFEM); } }
                else { if (workObjTag != myFEM.Tag) { theSession.Parts.SetWork((NXOpen.BasePart)myFEM); } }


                // Initializations
                // ---------------
                List<NXOpen.CAE.Connections.IConnection> newBoltConnections = new List<NXOpen.CAE.Connections.IConnection>();

                // Get existing Universal Connections
                // ----------------------------------
                List<string> existingUnivConnNames = isAFEM 
                    ? myAFEM.BaseFEModel.ConnectionsContainer.GetAllConnections().Select(x => x.Name).ToList()
                    : myFEM.BaseFEModel.ConnectionsContainer.GetAllConnections().Select(x => x.Name).ToList();


                //foreach (string connName in existingUnivConnNames)
                //{
                //    log += "      Existing Connection:  " + connName + Environment.NewLine;
                //}
                
                // Loop through all predefined Bolt Definitions
                foreach (MODELS.BoltDefinition boltDefinition in allBoltDefinitions)
                {
                    try
                    {
                        // Check if not existing yet
                        if (existingUnivConnNames.Contains(boltDefinition.Name))
                        {
                            // Delete existing Bolt Connection, to make sure:
                            // - it can be adapted if the desired properties are different
                            // - it can be re-generated, but only if the related Selection Recipe has entities in it
                            NXOpen.CAE.Connections.IConnection connToDelete = null;

                            //log += "TRYING TO DELETE EXISTING BOLT CONNECTION" + Environment.NewLine;
                            if (isAFEM)
                            {
                                connToDelete = myAFEM.BaseFEModel.ConnectionsContainer.GetAllConnections().ToList().Single(x => x.Name == boltDefinition.Name);
                            }
                            else
                            {
                                connToDelete = myFEM.BaseFEModel.ConnectionsContainer.GetAllConnections().ToList().Single(x => x.Name == boltDefinition.Name);
                            }

                            theSession.UpdateManager.AddObjectsToDeleteList(new TaggedObject[] { connToDelete });
                            theSession.UpdateManager.DoUpdate(new Session.UndoMarkId());
                            //continue;

                            log += "      BOLT DEFINITION:  " + boltDefinition.Name.ToUpper() + "  --> exists (deleted for re-creation)" + Environment.NewLine;
                        }
                        else
                        {
                            log += "      BOLT DEFINITION:  " + boltDefinition.Name.ToUpper() + Environment.NewLine;
                        }


                        // GET ALL TARGET CURVES
                        // ---------------------
                        // Initializations
                        NXOpen.CAE.SelectionRecipe targSelRecipe = null;
                        List<NXOpen.Line> targCurvesAsLine = new List<Line>();

                        // Get curves
                        switch (targCurveSearching)
                        {
                            // SELECTION RECIPE BASED:
                            case CurveSearchingMethod.SelectionRecipe:
                                // Check if target Selection Recipe exists and contains any Curves at all
                                // ----------------------------------------------------------------------
                                try
                                {
                                    targSelRecipe = targCAEPart.SelectionRecipes.ToArray().Single(x => x.Name.ToUpper().Contains(boltDefinition.Name.ToUpper()) && !x.Name.ToLower().Contains("unique"));
                                }
                                catch (Exception e)
                                {
                                    log += "         Related Selection Recipe not found --> skipped" + Environment.NewLine;
                                    continue;
                                }

                                if (targSelRecipe.GetEntities().Length == 0)
                                {
                                    log += "         Related Selection Recipe (" + targSelRecipe.Name + ") :  " + targSelRecipe.GetEntities().Length.ToString() + " entitities " +
                                        "--> skipped" + Environment.NewLine;
                                    continue;
                                }
                                else
                                {
                                    log += "         Related Selection Recipe (" + targSelRecipe.Name + ") :  " + targSelRecipe.GetEntities().Length.ToString() + " entitities " + Environment.NewLine;
                                }
                                break;

                            // LINE OCCURRENCE BASED:
                            case CurveSearchingMethod.LineOccurrence:
                                string targCurveName = "CURVE_" + boltDefinition.Name;
                                targCurvesAsLine = GetAllCurveOccurrences(targCAEPart.Tag).Where(x => x.Name.ToUpper() == targCurveName.ToUpper()).ToList();

                                if (targCurvesAsLine.Count > 0)
                                {
                                    log += "         Target Line objects (" + targCurveName + ") :  " + targCurvesAsLine.Count.ToString() + " entitities " + Environment.NewLine;
                                }
                                else
                                {
                                    log += "         Target Line objects (" + targCurveName + ") :  NONE FOUND --> skipped" + Environment.NewLine;
                                    continue;
                                }
                                break;

                            default:
                                break;
                        }
                        

                        // Create Universal Bolt Connection definition (NX12)
                        // --------------------------------------------------
                        NXOpen.CAE.Connections.Bolt newBoltConn = isAFEM
                            ? (NXOpen.CAE.Connections.Bolt)myAFEM.BaseFEModel.ConnectionsContainer.CreateConnection(NXOpen.CAE.Connections.ConnectionType.Bolt, boltDefinition.Name)
                            : (NXOpen.CAE.Connections.Bolt)myFEM.BaseFEModel.ConnectionsContainer.CreateConnection(NXOpen.CAE.Connections.ConnectionType.Bolt, boltDefinition.Name);

                        // Set Name
                        newBoltConn.SetName(boltDefinition.Name);
                        log += "         Name            : " + boltDefinition.Name + Environment.NewLine;

                        // Set Targets (Flanges)
                        newBoltConn.AddFlangeEntities(0, targCAEPart.SelectionRecipes.ToArray().Single(x => x.Name.ToUpper() == "GET ALL MESHES").GetEntities());
                        log += "         Targets         : " + targCAEPart.SelectionRecipes.ToArray().Single(x => x.Name.ToUpper() == "GET ALL MESHES").Name + " (Selection Recipe)" + Environment.NewLine;


                        // Set Locations
                        switch (targCurveSearching)
                        {
                            case CurveSearchingMethod.SelectionRecipe:
                                newBoltConn.AddLocationSelectionRecipe(0, targSelRecipe);
                                log += "         Locations       : " + targSelRecipe.Name + " (Selection Recipe based)" + Environment.NewLine;
                                break;

                            case CurveSearchingMethod.LineOccurrence:
                                int k = 0;
                                foreach (NXOpen.Line curveAsLine in targCurvesAsLine)
                                {
                                    newBoltConn.AddLocationCoordinatesWithDirectionCoordinates(
                                        0,
                                        targCAEPart.Points.CreatePoint(curveAsLine.StartPoint),
                                        targCAEPart.Points.CreatePoint(curveAsLine.EndPoint));

                                    k++;
                                    SetNXstatusMessage(baseMsg + " | Bolt Def:  " + boltDefinition.Name + "   | Setting Locations:  " + k.ToString() + @"/" + targCurvesAsLine.Count.ToString() + " added...");
                                }
                                log += "         Locations       : " + "CURVE_" + boltDefinition.Name + " (Line Occurrence based)  --> added via Coordinates with Direction Coordinates method" + Environment.NewLine;
                                break;

                            default:
                                break;
                        }
                        
                        
                        // Set Physicals
                        newBoltConn.DiameterType = NXOpen.CAE.Connections.DiameterType.User;
                        newBoltConn.Diameter.Value = boltDefinition.ShankDiam;
                        newBoltConn.HeadDiameterType = NXOpen.CAE.Connections.HeadDiameterType.User;
                        newBoltConn.HeadDiameter.Value = boltDefinition.HeadDiam;
                        newBoltConn.MaxBoltLength.Value = boltDefinition.MaxConnLength;

                        log += "         Shank Diameter  : " + newBoltConn.Diameter.Value.ToString() + Environment.NewLine;
                        log += "         Head Diameter   : " + newBoltConn.HeadDiameter.Value.ToString() + Environment.NewLine;
                        log += "         Max Bolt Length : " + newBoltConn.MaxBoltLength.Value.ToString() + Environment.NewLine;

                        // Set Material
                        NXOpen.PhysicalMaterial targMaterial = null;
                        bool isPhysicalMaterial = targCAEPart.MaterialManager.PhysicalMaterials.ToArray().Select(x => x.Name).Contains(boltDefinition.MaterialName);

                        if (isPhysicalMaterial)
                        {
                            targMaterial = (NXOpen.PhysicalMaterial)targCAEPart.MaterialManager.PhysicalMaterials.ToArray().Single(x => x.Name.ToUpper() == boltDefinition.MaterialName.ToUpper());
                        }
                        else
                        {
                            try
                            {
                                targMaterial = targCAEPart.MaterialManager.PhysicalMaterials.LoadFromNxmatmllibrary(boltDefinition.MaterialName);
                                goto matfound;
                            }
                            catch (Exception) { }
                            try
                            {
                                targMaterial = targCAEPart.MaterialManager.PhysicalMaterials.LoadFromLegacynxlibrary(boltDefinition.MaterialName);
                                goto matfound;
                            }
                            catch (Exception) { }
                            try
                            {
                                targMaterial = targCAEPart.MaterialManager.PhysicalMaterials.LoadFromNxlibrary(boltDefinition.MaterialName);
                                goto matfound;
                            }
                            catch (Exception) { }
                            try
                            {
                                string custom_library_path = "";
                                log += "TARGENV = " + targEnv.ToString();
                                if (targEnv == TargEnv.Siemens) 
                                { 
                                    custom_library_path = @"D:\3__TEAMCENTER\2_Projects\2_OCE_TCSimRollOut\4_Automatic_Bolt_Connections__Part_Families\CUSTOMER DATA\MATERIAL LIBRARIES\oce_material_library.xml"; 
                                }
                                else
                                {
                                    custom_library_path = Environment.GetEnvironmentVariable("PLMHOST") + @"\plmshare\config\nxcustom\NX-v12\UGII\materials\oce_material_library.xml";
                                }

                                if (!File.Exists(custom_library_path)) { log += "         Could not find custom material library path! (" + custom_library_path + ")" + Environment.NewLine; }
                                targMaterial = targCAEPart.MaterialManager.PhysicalMaterials.LoadFromMatmlLibrary(custom_library_path, boltDefinition.MaterialName);
                                goto matfound;
                            }
                            catch (Exception) { }

                            log += "         ! MATERIAL COULD NOT BE FOUND :  " + boltDefinition.MaterialName + Environment.NewLine;
                            continue;
                        }


                    matfound:;
                        newBoltConn.Material = targMaterial;
                        log += "         Material        : " + targMaterial.Name + Environment.NewLine;

                        log += "         CREATE  =  success" + Environment.NewLine;


                        // Add newly created Bolt Connection to list for realization step at end
                        newBoltConnections.Add(newBoltConn);
                    }
                    catch (Exception er)
                    {
                        log += "!ERROR occurred while creating Bolt Connection: " + Environment.NewLine +
                        er.ToString() + Environment.NewLine;
                    }
                }

                // Realize all new Universal Bolt Connection definitions
                // -----------------------------------------------------
                log += Environment.NewLine +
                    "   REALIZE UNIVERSAL BOLT CONNECTIONS" + Environment.NewLine +
                    "   ----------------------------------" + Environment.NewLine;
                SetNXstatusMessage(baseMsg + "   | Realizing all bolt connections' 1D elements...");

                NXOpen.CAE.Connections.Element boltConnElement = isAFEM
                    ? myAFEM.BaseFEModel.ConnectionElementCollection.Create(NXOpen.CAE.Connections.ElementType.E1DSpider, "Element - BOLT DEFINITIONS", newBoltConnections.ToArray())
                    : myFEM.BaseFEModel.ConnectionElementCollection.Create(NXOpen.CAE.Connections.ElementType.E1DSpider, "Element - BOLT DEFINITIONS", newBoltConnections.ToArray());

                boltConnElement.GenerateElements();
                log += "      Elements generated" + Environment.NewLine;


                log += "      REALIZATION =  success" + Environment.NewLine;
            }
            catch (Exception e)
            {
                log += "!ERROR occurred: " + Environment.NewLine +
                    e.ToString() + Environment.NewLine;
            }
        }
        

        /// <summary>
        /// Update (A)FEM object's Bolt Connections to make sure new Universal Bolt Connections are properly finalized
        /// </summary>
        /// <param name="myCAEpart"></param>
        private static void UpdateCAEObjectConnections(NXOpen.CAE.BaseFemPart myCAEpart)
        {
            log += Environment.NewLine +
                "   UPDATE NEW BOLT CONNECTIONS" + Environment.NewLine + Environment.NewLine;

            // Set target (A)FEM to working
            //NXOpen.CAE.BaseFemPart myCAEpart = (NXOpen.CAE.BaseFemPart)targObj;
            if (theSession.Parts.BaseWork.Tag != myCAEpart.Tag) { theSession.Parts.SetWork(myCAEpart); log += "MADE WORKING" + Environment.NewLine; }

            // Force "update" status for each Universal Bolt Connection
            foreach (NXOpen.CAE.Connections.IConnection myConn in myCAEpart.BaseFEModel.ConnectionsContainer.GetAllConnections()
                .Where(x => x.GetType().ToString() == "NXOpen.CAE.Connections.Bolt"))
            {
                try
                {
                    // Check if it is a Universal Bolt Connection
                    NXOpen.CAE.Connections.Bolt myBoltConn = (NXOpen.CAE.Connections.Bolt)myConn;

                    // Force an "update" of the Universal Bolt Connection
                    myBoltConn.MaxBoltLength.Value++;
                    myBoltConn.MaxBoltLength.Value--;

                    log += "      Update forced of Bolt Connection:  " + myBoltConn.Name.ToUpper() + Environment.NewLine;
                }
                catch (Exception)
                {
                    // Not a Universal Bolt Connection
                }
            }

            // Update AFEM to realize all Universal Bolt Connections
            myCAEpart.BaseFEModel.UpdateFemodel();
            log += "      UPDATED:  " + myCAEpart.Name.ToUpper() + Environment.NewLine;
        }


        /// <summary>
        /// Gets all Occurrences of type NXOpen.Line (representing a Curve, for example) contained in a target object
        /// </summary>
        /// <param name="targObjTag">Target object's Tag</param>
        /// <returns>All Occurrences of type NXOpen.Line</returns>
        private static List<NXOpen.Line> GetAllCurveOccurrences(NXOpen.Tag targObjTag)
        {
            // Initializations
            List<NXOpen.Line> allCurveOccurrences = new List<Line>();
            NXOpen.Tag nextTag = NXOpen.Tag.Null;
            NXOpen.NXObject obj = null;

            // Cycle through all objects of target object
            do
            {
                // Get next object's Tag
                nextTag = theUfSession.Obj.CycleAll(targObjTag, nextTag);
                if (nextTag == NXOpen.Tag.Null) { break; }

                // Get next object
                obj = (NXOpen.NXObject)NXOpen.Utilities.NXObjectManager.Get(nextTag);

                // Check if object is a line object
                //if (obj.IsOccurrence && obj.GetType().ToString() == "NXOpen.Line")
                if (obj.GetType().ToString() == "NXOpen.Line")
                {
                    //WriteToLog("LINE OCC FOUND:  " + obj.Name + "  (Prototype name = " + obj.Prototype.Name + ")");
                    allCurveOccurrences.Add((NXOpen.Line)obj);
                }
            } while (true);

            //// Print out available curves in log
            //log += "      Available curves:" + Environment.NewLine;
            //foreach (NXOpen.Line line in allCurveOccurrences)
            //{
            //    log += "         " + line.Name;
            //}

            // Return result
            return allCurveOccurrences;
        }


        /// <summary>
        /// Replaces the References Set of the target FEM object to the "CAE" Reference Set
        /// </summary>
        /// <param name="myFEM">Target FEM object</param>
        private static void ReplaceReferenceSet(NXOpen.CAE.FemPart myFEM)
        {
            log += Environment.NewLine +
                "   REPLACE REFERENCE SET" + Environment.NewLine +
                "   ---------------------" + Environment.NewLine;

            try
            {
                // Initializations
                List<NXOpen.Assemblies.Component> targComponents = new List<NXOpen.Assemblies.Component>();
                //string targReferenceSet = "Entire Part";  --> Moved to Global Variables
                //string targReferenceSet = "CAE";

                // Get all underlying Components to change the Reference Set for
                targComponents = GetAllComponents(myFEM.MasterCadPart.ComponentAssembly.RootComponent, targComponents);

                log += "      Target components:" + Environment.NewLine;
                foreach (NXOpen.Assemblies.Component component in targComponents)
                {
                    log += "         " + component.Name.ToUpper() + Environment.NewLine;
                }

                // Change the Reference Set of each target component to the "CAE" Reference Set
                NXOpen.ErrorList errorList = myFEM.ComponentAssembly.ReplaceReferenceSetInOwners(targReferenceSet, targComponents.ToArray());
                errorList.Dispose();

                log += "      Changed Reference Set to:   " + targReferenceSet + Environment.NewLine;
            }
            catch (Exception e)
            {
                log += "!ERROR occurred: " + e.ToString() + Environment.NewLine;
            }
        }


        /// <summary>
        /// Loop through all underlying components and collect them in a list
        /// </summary>
        /// <param name="targComponent">Object to start looking from</param>
        /// <param name="targComponents">List to collect all component objects</param>
        /// <returns></returns>
        private static List<NXOpen.Assemblies.Component> GetAllComponents(NXOpen.Assemblies.Component targComponent, List<NXOpen.Assemblies.Component> targComponents)
        {
            targComponents.Add(targComponent);

            // Loop through child Components
            foreach (NXOpen.Assemblies.Component childComp in targComponent.GetChildren())
            {
                targComponents = GetAllComponents(childComp, targComponents);
            }

            return targComponents;
        }


        /// <summary>
        /// Updates the Geometry Options of a FEM object, so that it includes Lines, Arcs & Circles, Splines, Conics and Sketch Curves
        /// </summary>
        /// <param name="targFEM">Target FEM object</param>
        private static void SetFEMGeometryOptions(NXOpen.CAE.FemPart targFEM)
        {
            log += Environment.NewLine +
                "   SET GEOMETRY OPTIONS" + Environment.NewLine +
                "   --------------------" + Environment.NewLine;

            // Create FemSynchronizeOptions
            NXOpen.CAE.FemSynchronizeOptions targFemSynchronizeOptions = targFEM.NewFemSynchronizeOptions();

            // Configure FemSynchronizeOptions
            targFemSynchronizeOptions.SynchronizePointsFlag = false;
            targFemSynchronizeOptions.SynchronizeCoordinateSystemFlag = false;
            targFemSynchronizeOptions.SynchronizeLinesFlag = true;
            targFemSynchronizeOptions.SynchronizeArcsFlag = true;
            targFemSynchronizeOptions.SynchronizeSplinesFlag = true;
            targFemSynchronizeOptions.SynchronizeConicsFlag = true;
            targFemSynchronizeOptions.SynchronizeSketchCurvesFlag = true;

            // Assign FemSynchronizeOptions in Geometry Data settings
            List<Body> targBodies = new List<Body>();
            targFEM.SetGeometryData(NXOpen.CAE.FemPart.UseBodiesOption.AllBodies, targBodies.ToArray(), targFemSynchronizeOptions);

            log += "      Set to: LINES, ARCS & CIRCLES, SPLINES, CONICS, SKETCH CURVES" + Environment.NewLine;
        }


        /// <summary>
        /// Execute VALIDATOR tool, after Creator tool has successfully created all pre-defined Universal Bolt Connections
        /// </summary>
        private static void RunValidator()
        {



            return;




            // If desired, run Validator after successful Creator execution
            if (File.Exists(PathToValidatorExe))
            {
                if (RunValidatorAfter)
                {
                    //List<System.String> inputArgs = new List<System.String>();
                    //inputArgs.Add("test");
                    //inputArgs.Add("test2");
                    List<Object> inputArgs = new List<object>();
                    //inputArgs.Add(true);
                    inputArgs.Add("-path=test");
                    inputArgs.Add("-path=test");
                    theUI.NXMessageBox.Show("Input arguments", NXMessageBox.DialogType.Information, inputArgs.ToString());

                    theSession.Execute(PathToValidatorExe, "Program", "Main", inputArgs.ToArray());
                    //theSession.Execute(PathToValidatorExe, "Program", "SetNXstatusMessage", new string[] { "VALIDATOR RUN FROM CREATOR"});
                }
            }
            else
            {
                theUI.NXMessageBox.Show("Running Validator went wrong:", NXMessageBox.DialogType.Warning, "Could not find Validator executable:" + Environment.NewLine +
                    Environment.NewLine +
                    "Target path = " + PathToValidatorExe + Environment.NewLine +
                    "(Target environment:  " + targEnv.ToString() + ")" + Environment.NewLine +
                    Environment.NewLine +
                    "Are you sure this path exists?" + Environment.NewLine +
                    "If yes, ask Siemens to put in the correct target path for the Validator executable.");
            }
        }
        #endregion
    }
}
