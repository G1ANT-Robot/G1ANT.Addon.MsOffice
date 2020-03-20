using G1ANT.Addon.MSOffice.Api.Access;
using G1ANT.Addon.MSOffice.Controllers;
using G1ANT.Addon.MSOffice.Controllers.Access;
using G1ANT.Addon.MSOffice.Models.Access;
using G1ANT.Addon.MSOffice.Models.Access.Dao;
using G1ANT.Language;
using System;
using System.Linq;
using System.Windows.Forms;

namespace G1ANT.Addon.MSOffice.Panels
{
    [Panel(Name = "Access forms and controls tree", DockingSide = DockingSide.Right, InitialAppear = false, Width = 400, 
        Description = "Panel with Access forms and controls tree")]
    public partial class AccessControlsTreePanel : RobotPanel
    {
        private AccessUIControlsTreeController controller;

        public AccessControlsTreePanel()
            : this(null)
        { }

        public AccessControlsTreePanel(AccessUIControlsTreeController controller) // IoC
        {
            InitializeComponent();
            this.controller = controller ?? new AccessUIControlsTreeController(
                controlsTree,
                comboBox1,
                new RunningObjectTableService(),
                new TooltipService()
            );
        }

        public override void Initialize(IMainForm mainForm)
        {
            base.Initialize(mainForm);
            controller.Initialize(mainForm);
        }

        public override void RefreshContent() => controller.InitRootElements();

        private void controlsTree_BeforeExpand(object sender, TreeViewCancelEventArgs e) => controller.TryLoadChildNodes(e.Node);

        private void controlsTree_DoubleClick(object sender, EventArgs e) => controller.InsertPathIntoScript(controlsTree.SelectedNode);

        private void insertWPathButton_Click(object sender, EventArgs e) => controller.InsertPathIntoScript(controlsTree.SelectedNode);

        private void refreshButton_Click(object sender, EventArgs e) => controller.InitRootElements(true);

        private void highlightToolStripMenuItem_Click(object sender, EventArgs e) => controller.ShowMarkerForm(controlsTree.SelectedNode);

        private void controlsTree_ElementMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                controlsTree.SelectedNode = e.Node;
                contextMenuStrip.Show(MousePosition);
            }
        }

        private void toolStripButton1_Click(object sender, EventArgs e) => controller.ShowMarkerForm(controlsTree.SelectedNode);

        private void copyNodeDetailsToolStripMenuItem_Click(object sender, EventArgs e) => controller.CopyNodeDetails(controlsTree.SelectedNode);

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e) => controller.SelectedApplicationChanged(comboBox1.SelectedItem as RotApplicationModel);

        private void contextMenuStrip_Opening(object sender, System.ComponentModel.CancelEventArgs e)
        {
            var clickedNode = controlsTree.SelectedNode;
            if (clickedNode == null)
            {
                e.Cancel = true;
                return;
            }
            var clickedModel = clickedNode.Tag;
            
            //todo: move to controller and use constants
            loadFormToolStripMenuItem.Available = clickedModel is AccessObjectModel && clickedNode.Parent?.Text == "Forms";

            openToolStripMenuItem.Available = clickedModel is AccessObjectModel && !loadFormToolStripMenuItem.Available;

            copynameToolStripMenuItem.Available = clickedNode.Parent != null && !(clickedModel is RotApplicationModel);

            highlightToolStripMenuItem.Available = clickedModel is AccessControlModel;
            copyNodeDetailsToolStripMenuItem.Available = clickedModel is AccessControlModel;

            viewDataToolStripMenuItem.Available = clickedModel is AccessTableDefModel;

            e.Cancel = !contextMenuStrip.Items.OfType<ToolStripMenuItem>().Any(item => item.Available);
        }

        private void acNormalToolStripMenuItem_Click(object sender, EventArgs e)
        {
            controller.TryOpenFormFromSelectedNode(false);
            // todo: replace current tree node with loaded form data
        }

        private void acDesignToolStripMenuItem_Click(object sender, EventArgs e)
        {
            controller.TryOpenFormFromSelectedNode(true);
            // todo: replace current tree node with loaded form data
        }

        private void copynameToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(controlsTree.SelectedNode?.Text);
        }

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var model = (AccessObjectModel)controlsTree.SelectedNode.Tag;
            controller.OpenAccessObject(model);
        }

        private void viewDataToolStripMenuItem_Click(object sender, EventArgs e) => controller.ViewDataFromTable(controlsTree.SelectedNode.Text);
    }
}
