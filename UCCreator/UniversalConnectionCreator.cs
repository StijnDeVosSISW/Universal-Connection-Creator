using System;
using System.Collections.Generic;
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
        private string theDlxFileName;
        private NXOpen.BlockStyler.BlockDialog theDialog;
        private NXOpen.BlockStyler.Group group0;// Block type: Group
        private NXOpen.BlockStyler.Label label0;// Block type: Label
        private NXOpen.BlockStyler.Separator separator0;// Block type: Separator
        private NXOpen.BlockStyler.Tree tree_control0;// Block type: Tree Control
        private NXOpen.BlockStyler.Button button_ADD;// Block type: Button
        private NXOpen.BlockStyler.Button button_Delete;// Block type: Button
        private NXOpen.BlockStyler.Group group;// Block type: Group
        private NXOpen.BlockStyler.StringBlock TB_ADD_Name;// Block type: String
        private NXOpen.BlockStyler.IntegerBlock integer_ADD_ShankDiam;// Block type: Integer
        private NXOpen.BlockStyler.IntegerBlock integer_ADD_HeadDiam;// Block type: Integer
        private NXOpen.BlockStyler.IntegerBlock integer_ADD_MaxConnLength;// Block type: Integer
        private NXOpen.BlockStyler.StringBlock TB_ADD_Material;// Block type: String

        private List<NXOpen.BlockStyler.Node> allNodes = new List<Node>();
        private List<MODELS.BoltDefinition> allBoltDefinitions = new List<MODELS.BoltDefinition>();

        //------------------------------------------------------------------------------
        //Constructor for NX Styler class
        //------------------------------------------------------------------------------
        public UniversalConnectionCreator()
        {
            try
            {
                theSession = Session.GetSession();
                theUI = UI.GetUI();
                theDlxFileName = "UniversalConnectionCreator.dlx";
                theDialog = theUI.CreateDialog(theDlxFileName);
                theDialog.AddApplyHandler(new NXOpen.BlockStyler.BlockDialog.Apply(apply_cb));
                theDialog.AddOkHandler(new NXOpen.BlockStyler.BlockDialog.Ok(ok_cb));
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
                button_ADD = (NXOpen.BlockStyler.Button)theDialog.TopBlock.FindBlock("button_ADD");
                button_Delete = (NXOpen.BlockStyler.Button)theDialog.TopBlock.FindBlock("button_Delete");
                group = (NXOpen.BlockStyler.Group)theDialog.TopBlock.FindBlock("group");
                TB_ADD_Name = (NXOpen.BlockStyler.StringBlock)theDialog.TopBlock.FindBlock("TB_ADD_Name");
                integer_ADD_ShankDiam = (NXOpen.BlockStyler.IntegerBlock)theDialog.TopBlock.FindBlock("integer_ADD_ShankDiam");
                integer_ADD_HeadDiam = (NXOpen.BlockStyler.IntegerBlock)theDialog.TopBlock.FindBlock("integer_ADD_HeadDiam");
                integer_ADD_MaxConnLength = (NXOpen.BlockStyler.IntegerBlock)theDialog.TopBlock.FindBlock("integer_ADD_MaxConnLength");
                TB_ADD_Material = (NXOpen.BlockStyler.StringBlock)theDialog.TopBlock.FindBlock("TB_ADD_Material");

                tree_control0.Width = 500;

                //------------------------------------------------------------------------------
                //Registration of Treelist specific callbacks
                //------------------------------------------------------------------------------
                //tree_control0.SetOnExpandHandler(new NXOpen.BlockStyler.Tree.OnExpandCallback(OnExpandCallback));

                //tree_control0.SetOnInsertColumnHandler(new NXOpen.BlockStyler.Tree.OnInsertColumnCallback(OnInsertColumnCallback));

                //tree_control0.SetOnInsertNodeHandler(new NXOpen.BlockStyler.Tree.OnInsertNodeCallback(OnInsertNodecallback));

                //tree_control0.SetOnDeleteNodeHandler(new NXOpen.BlockStyler.Tree.OnDeleteNodeCallback(OnDeleteNodecallback));

                //tree_control0.SetOnPreSelectHandler(new NXOpen.BlockStyler.Tree.OnPreSelectCallback(OnPreSelectcallback));

                //tree_control0.SetOnSelectHandler(new NXOpen.BlockStyler.Tree.OnSelectCallback(OnSelectcallback));

                //tree_control0.SetOnStateChangeHandler(new NXOpen.BlockStyler.Tree.OnStateChangeCallback(OnStateChangecallback));

                //tree_control0.SetToolTipTextHandler(new NXOpen.BlockStyler.Tree.ToolTipTextCallback(ToolTipTextcallback));

                //tree_control0.SetColumnSortHandler(new NXOpen.BlockStyler.Tree.ColumnSortCallback(ColumnSortcallback));

                //tree_control0.SetStateIconNameHandler(new NXOpen.BlockStyler.Tree.StateIconNameCallback(StateIconNameCallback));

                //tree_control0.SetOnBeginLabelEditHandler(new NXOpen.BlockStyler.Tree.OnBeginLabelEditCallback(OnBeginLabelEditCallback));

                //tree_control0.SetOnEndLabelEditHandler(new NXOpen.BlockStyler.Tree.OnEndLabelEditCallback(OnEndLabelEditCallback));

                //tree_control0.SetOnEditOptionSelectedHandler(new NXOpen.BlockStyler.Tree.OnEditOptionSelectedCallback(OnEditOptionSelectedCallback));

                //tree_control0.SetAskEditControlHandler(new NXOpen.BlockStyler.Tree.AskEditControlCallback(AskEditControlCallback));

                //tree_control0.SetOnMenuHandler(new NXOpen.BlockStyler.Tree.OnMenuCallback(OnMenuCallback));;

                //tree_control0.SetOnMenuSelectionHandler(new NXOpen.BlockStyler.Tree.OnMenuSelectionCallback(OnMenuSelectionCallback));;

                //tree_control0.SetIsDropAllowedHandler(new NXOpen.BlockStyler.Tree.IsDropAllowedCallback(IsDropAllowedCallback));;

                //tree_control0.SetIsDragAllowedHandler(new NXOpen.BlockStyler.Tree.IsDragAllowedCallback(IsDragAllowedCallback));;

                //tree_control0.SetOnDropHandler(new NXOpen.BlockStyler.Tree.OnDropCallback(OnDropCallback));;

                //tree_control0.SetOnDropMenuHandler(new NXOpen.BlockStyler.Tree.OnDropMenuCallback(OnDropMenuCallback));

                //tree_control0.SetOnDefaultActionHandler(new NXOpen.BlockStyler.Tree.OnDefaultActionCallback(OnDefaultActionCallback));

                //------------------------------------------------------------------------------
                //------------------------------------------------------------------------------
                //Registration of StringBlock specific callbacks
                //------------------------------------------------------------------------------
                //TB_ADD_Name.SetKeystrokeCallback(new NXOpen.BlockStyler.StringBlock.KeystrokeCallback(KeystrokeCallback));

                //TB_ADD_Material.SetKeystrokeCallback(new NXOpen.BlockStyler.StringBlock.KeystrokeCallback(KeystrokeCallback));

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
                // Intialize GUI
                

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

                NXOpen.BlockStyler.TreeListMenu treeMenu = tree_control0.CreateMenu();
                treeMenu.AddMenuItem(0, "Add");
                treeMenu.AddMenuItem(1, "Delete");
                tree_control0.SetMenu(treeMenu);
                tree_control0.Redraw(true);
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
                else if (block == button_ADD)
                {
                    //---------Enter your code here-----------
                }
                else if (block == button_Delete)
                {
                    //---------Enter your code here-----------
                }
                else if (block == TB_ADD_Name)
                {
                    //---------Enter your code here-----------
                }
                else if (block == integer_ADD_ShankDiam)
                {
                    //---------Enter your code here-----------
                }
                else if (block == integer_ADD_HeadDiam)
                {
                    //---------Enter your code here-----------
                }
                else if (block == integer_ADD_MaxConnLength)
                {
                    //---------Enter your code here-----------
                }
                else if (block == TB_ADD_Material)
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

        //public void OnSelectcallback(NXOpen.BlockStyler.Tree tree, NXOpen.BlockStyler.Node node, int columnID, bool Selected)
        //{
        //}

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

        //public Tree.BeginLabelEditState OnBeginLabelEditCallback(NXOpen.BlockStyler.Tree tree, NXOpen.BlockStyler.Node node, int columnID)
        //{
        //}

        //public Tree.EndLabelEditState OnEndLabelEditCallback(NXOpen.BlockStyler.Tree tree, NXOpen.BlockStyler.Node node, int columnID, string editedText)
        //{
        //}

        //public Tree.EditControlOption OnEditOptionSelectedCallback(NXOpen.BlockStyler.Tree tree, NXOpen.BlockStyler.Node node, int columnID, int selectedOptionID, string selectedOptionText, Tree.ControlType type)
        //{
        //}

        //public Tree.ControlType AskEditControlCallback(NXOpen.BlockStyler.Tree tree, NXOpen.BlockStyler.Node node, int columnID)
        //{
        //}

        //public void OnMenuCallback(NXOpen.BlockStyler.Tree tree, NXOpen.BlockStyler.Node node, int columnID)
        //{
        //}

        //public void OnMenuSelectionCallback(NXOpen.BlockStyler.Tree tree, NXOpen.BlockStyler.Node node, int menuItemID)
        //{
        //}

        //public Node.DropType IsDropAllowedCallback(NXOpen.BlockStyler.Tree tree, NXOpen.BlockStyler.Node node, int columnID, NXOpen.BlockStyler.Node targetNode, int targetColumnID)
        //{
        //}

        //public Node.DragType IsDragAllowedCallback(NXOpen.BlockStyler.Tree tree, NXOpen.BlockStyler.Node node, int columnID)
        //{
        //}

        //public bool OnDropCallback(NXOpen.BlockStyler.Tree tree, NXOpen.BlockStyler.Node[] node, int columnID, NXOpen.BlockStyler.Node targetNode, int targetColumnID, Node.DropType dropType, int dropMenuItemId)
        //{
        //}

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
    }
}
