using G1ANT.Addon.MSOffice.Controllers;
using G1ANT.Language;
using System;
using System.Windows.Forms;

namespace G1ANT.Addon.MSOffice.Panels
{
    [Panel(Name = "Access forms and controls tree", DockingSide = DockingSide.Right, InitialAppear = false, Width = 400, 
        Description = "Panel with Access forms and controls tree")]
    public partial class AccessControlsTreePanel : RobotPanel
    {
        private AccessUIControlsTreeController controller;

        public AccessControlsTreePanel() : this(new AccessUIControlsTreeController())
        { }

        public AccessControlsTreePanel(AccessUIControlsTreeController controller) // IoC
        {
            this.controller = controller;
            InitializeComponent();
        }


        public override void Initialize(IMainForm mainForm)
        {
            base.Initialize(mainForm);
            controller.Initialize(mainForm);
        }

        public override void RefreshContent() => controller.InitRootElements(controlsTree);

        private void controlsTree_BeforeExpand(object sender, TreeViewCancelEventArgs e) => controller.LoadChildNodes(e.Node);

        private void controlsTree_DoubleClick(object sender, EventArgs e) => controller.InsertPathIntoScript(controlsTree.SelectedNode);

        private void insertWPathButton_Click(object sender, EventArgs e) => controller.InsertPathIntoScript(controlsTree.SelectedNode);

        private void refreshButton_Click(object sender, EventArgs e) => controller.InitRootElements(controlsTree);

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
    }
}
