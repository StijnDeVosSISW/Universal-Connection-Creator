using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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
        private NXOpen.BlockStyler.Group group;// Block type: Group
        private NXOpen.BlockStyler.FileSelection nativeFileBrowser0;// Block type: NativeFileBrowser
        private NXOpen.BlockStyler.Button button_IMPORT;// Block type: Button
        private NXOpen.BlockStyler.Button button_CREATE;// Block type: Button

        private List<NXOpen.BlockStyler.Node> allNodes = new List<Node>();
        private List<MODELS.BoltDefinition> allBoltDefinitions = new List<MODELS.BoltDefinition>();
        private enum MenuID
        {
            AddNode,
            DeleteNode
        };

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
                theDlxFileName = "UniversalConnectionCreator.dlx";
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
                group = (NXOpen.BlockStyler.Group)theDialog.TopBlock.FindBlock("group");
                nativeFileBrowser0 = (NXOpen.BlockStyler.FileSelection)theDialog.TopBlock.FindBlock("nativeFileBrowser0");
                button_IMPORT = (NXOpen.BlockStyler.Button)theDialog.TopBlock.FindBlock("button_IMPORT");
                button_CREATE = (NXOpen.BlockStyler.Button)theDialog.TopBlock.FindBlock("button_CREATE");

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

                // Initialize Tree Control (List of predefined Universal Bolt Connections)
                int default_width = 150;
                tree_control0.InsertColumn(0, "Name", default_width);
                tree_control0.InsertColumn(1, "Shank Diameter [mm]", default_width);
                tree_control0.InsertColumn(2, "Head Diameter [mm]", default_width);
                tree_control0.InsertColumn(3, "Maximum Connection Length [mm]", default_width);
                tree_control0.InsertColumn(4, "Material", default_width);

                allNodes.Add(tree_control0.CreateNode("test"));
                allNodes.Add(tree_control0.CreateNode("test2"));
                allNodes.Add(tree_control0.CreateNode("test3"));

                tree_control0.InsertNode(allNodes[0], null, null, Tree.NodeInsertOption.First);
                tree_control0.InsertNode(allNodes[1], null, null, Tree.NodeInsertOption.Last);
                tree_control0.InsertNode(allNodes[2], null, null, Tree.NodeInsertOption.Last);

                allNodes[0].SetColumnDisplayText(0, "M10X90");
                allNodes[0].SetColumnDisplayText(1, "10");
                allNodes[0].SetColumnDisplayText(2, "12");
                allNodes[0].SetColumnDisplayText(3, "90");
                allNodes[0].SetColumnDisplayText(4, "Aluminum_1942");

                allNodes[1].SetColumnDisplayText(0, "M10X80");
                allNodes[1].SetColumnDisplayText(1, "10");
                allNodes[1].SetColumnDisplayText(2, "12");
                allNodes[1].SetColumnDisplayText(3, "80");
                allNodes[1].SetColumnDisplayText(4, "Aluminum_1942");

                allNodes[2].SetColumnDisplayText(0, "M12X50");
                allNodes[2].SetColumnDisplayText(1, "12");
                allNodes[2].SetColumnDisplayText(2, "15");
                allNodes[2].SetColumnDisplayText(3, "50");
                allNodes[2].SetColumnDisplayText(4, "Aluminum_1942");

            }
            catch (Exception ex)
            {
                //---- Enter your exception handling code here -----
                theUI.NXMessageBox.Show("Block Styler", NXMessageBox.DialogType.Error, ex.ToString());
            }
        }

        //------------------------------------------------------------------------------
        //Callback Name: apply_cb
        //------------------------------------------------------------------------------
        public int apply_cb()
        {
            int errorCode = 0;
            try
            {
                //---- Enter your callback code here -----
            }
            catch (Exception ex)
            {
                //---- Enter your exception handling code here -----
                errorCode = 1;
                theUI.NXMessageBox.Show("Block Styler", NXMessageBox.DialogType.Error, ex.ToString());
            }
            return errorCode;
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
                else if (block == button_CREATE)
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
        //Callback Name: ok_cb
        //------------------------------------------------------------------------------
        public int ok_cb()
        {
            int errorCode = 0;
            try
            {
                errorCode = apply_cb();
                //---- Enter your callback code here -----
            }
            catch (Exception ex)
            {
                //---- Enter your exception handling code here -----
                errorCode = 1;
                theUI.NXMessageBox.Show("Block Styler", NXMessageBox.DialogType.Error, ex.ToString());
            }
            return errorCode;
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
