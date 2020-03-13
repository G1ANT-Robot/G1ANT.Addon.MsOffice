using G1ANT.Addon.MSOffice.Api.Access;
using G1ANT.Addon.MSOffice.Models.Access;
using G1ANT.Language;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace G1ANT.Addon.MSOffice.Controllers
{
    public class AccessUIControlsTreeController
    {
        private IMainForm mainForm;
        public bool initialized = false;
        private readonly IRunningObjectTableService runningObjectTableService;
        private readonly TreeView controlsTree;

        public AccessUIControlsTreeController(IRunningObjectTableService runningObjectTableService, TreeView controlsTree)
        {
            this.runningObjectTableService = runningObjectTableService;
            this.controlsTree = controlsTree;
        }

        public void Initialize(IMainForm mainForm) => this.mainForm = mainForm;


        public void InitRootElements(ComboBox applications)
        {
            //if (initialized)
            //    return;
            //initialized = true;


            var selectedItemText = applications.SelectedItem?.ToString();

            var applicationInstances = runningObjectTableService.GetApplicationInstances("msaccess");


            applications.Items.Clear();
            applications.Items.AddRange(applicationInstances.ToArray());


            if (applicationInstances.Any())
            {
                var itemToSelect = applications.Items.Cast<RotApplicationModel>().FirstOrDefault(a => a.ToString() == selectedItemText) ?? applications.Items[0];

                applications.SelectedItem = itemToSelect;
            }
        }

        internal void SelectedApplicationChanged(RotApplicationModel rotApplicationModel)
        {
            controlsTree.BeginUpdate();
            controlsTree.Nodes.Clear();

            controlsTree.Nodes.AddRange(
                new TreeNode[]
                {
                    new TreeNode("Forms") { Tag = rotApplicationModel, Nodes = { new TreeNode("") } },
                    new TreeNode("Macros") { Tag = rotApplicationModel, Nodes = { new TreeNode("") } },
                    new TreeNode("Reports") { Tag = rotApplicationModel, Nodes = { new TreeNode("") } },
                    new TreeNode("Queries") { Tag = rotApplicationModel, Nodes = { new TreeNode("") } },
                }
            );

            controlsTree.EndUpdate();
        }

        //private string FormatLongLine(string line)
        //{
        //    const int maxLineLength = 100;
        //    if (line.Length <= maxLineLength)
        //        return line;

        //    var sb = new StringBuilder(line.Length);
        //    var isFirstLine = true;
        //    do
        //    {
        //        var linePart = line.Substring(0, Math.Min(line.Length, maxLineLength));
        //        line = line.Substring(linePart.Length);
        //        sb.AppendLine((isFirstLine ? "" : "\t") + linePart);
        //        isFirstLine = false;
        //    } while (line != "");

        //    return sb.ToString();
        //}

        private static readonly string[] ControlTooltipProperties = new string[]
        {
            "Name",
            "ControlType",
            "FontName",
            "FontSize",
            "Caption",
            "Visible",
            "Width",
            "Height",
            "Top",
            "Left",
            "BackColor",
            "Enabled",
            "OnClick",
            "Default",
            "Cancel",
            "TabIndex",
            "TabStop",
            "RowSource",
            "RowSourceType",
            "BoundColumn",
            "ColumnCount",
            "ColumnWidths",
            "ColumnHeads",
            "ListRows",
            "ListWidth",
            "ListCount",
            "ListIndex"
        };

        private static readonly string[] FormTooltipProperties = new string[]
        {
            "Name",
            "Hwnd",
            "Modal",
            "Width",
            "WindowWidth",
            "WindowHeight",
            "InsideHeight",
            "InsideWidth",
            "WindowTop",
            "WindowLeft",
            "ScrollBars",
            "ControlBox",
            "CloseButton",
            "MinButton",
            "MaxButton",
            "MinMaxButtons",
            "Moveable",
            "GridX",
            "GridY",
            "ShowGrid",
            "LogicalPageWidth",
            "Visible",
        };

        //private static readonly string[] UnloadedFormTooltipProperties = new string[] { "Name", "Attributes", "DateCreated", "DateModified", "FullName", "Type" };

        private string GetTooltip(AccessControlModel controlModel)
        {
            var result = new StringBuilder($"Type: {controlModel.Type}\r\n");

            //foreach (var propertyName in ControlTooltipProperties)
            //{
            //    if (controlModel.TryGetPropertyValue(propertyName, out string propertyValue))
            //        result.AppendLine($"{propertyName}: {propertyValue}");
            //}

            return result.ToString();
        }

        private string GetTooltip(AccessFormModel formModel)
        {
            var result = new StringBuilder();

            foreach (var propertyName in FormTooltipProperties)
            {
                if (formModel.TryGetPropertyValue(propertyName, out string propertyValue))
                    result.AppendLine($"{propertyName}: {propertyValue}");
            }

            return result.ToString();
        }

        private string GetTooltip(AccessObjectModel formModel)
        {
            var result = new StringBuilder();

            result.AppendLine($"Name: {formModel.Name}");
            result.AppendLine($"FullName: {formModel.FullName}");
            result.AppendLine($"Type: {formModel.Type}");
            result.AppendLine($"IsLoaded: {formModel.IsLoaded}");
            result.AppendLine($"IsWeb: {formModel.IsWeb}");
            result.AppendLine($"Attributes: {formModel.Attributes}");
            result.AppendLine($"DateCreated: {formModel.DateCreated}");
            result.AppendLine($"DateModified: {formModel.DateModified}");

            for (var i = 0; i < formModel.Form.Properties.Count; ++i)
            {
                var property = formModel.Form.Properties[i];
                result.AppendLine($"{property.Name}: {property.Value}");
            }

            return result.ToString();
        }

        public void CopyNodeDetails(TreeNode treeNode)
        {
            Clipboard.SetText(treeNode.ToolTipText);
        }

        private static string GetNameForNode(AccessControlModel model)
        {
            return $"\"{model.Name}\" {model.Type} {model.Value}";
        }

        private static string GetNameForNode(AccessFormModel model)
        {
            return $"\"{model.Name}\" {(model.Name != model.Caption ? model.Caption : "")} {(model.Name != model.FormName ? model.FormName : "")}";
        }

        private string GetNameForNode(AccessObjectModel formModel)
        {
            return $"\"{formModel.Name}\" {(formModel.Name != formModel.FullName ? formModel.FullName : "")}, loaded: {formModel.IsLoaded}";
        }


        public void LoadChildNodes(TreeNode treeNode)
        {
            //if (treeNode.Parent == null)
            //    return; // don't clear jvms and their windows as they are already rendered

            controlsTree.BeginUpdate();

            if (treeNode.Tag is AccessControlModel accessControlModel)
                LoadControlNodes(treeNode, accessControlModel);
            else if (treeNode.Tag is AccessFormModel accessFormModel)
                LoadControlNodes(treeNode, accessFormModel);
            else if (treeNode.Tag is RotApplicationModel rotApplicationModel)
                LoadFormNodes(treeNode, rotApplicationModel);

            controlsTree.EndUpdate();
        }

        private void LoadFormNodes(TreeNode treeNode, RotApplicationModel rotApplicationModel)
        {
            if (treeNode.Nodes.Count == 1 && treeNode.Nodes[0].Text == "")
            {
                treeNode.Nodes.Clear();

                foreach (Microsoft.Office.Interop.Access.AccessObject form in rotApplicationModel.Application.CurrentProject.AllForms)
                {
                    if (form.IsLoaded)
                    {
                        var formModel = new AccessFormModel(rotApplicationModel.Application.Forms[form.Name], false, false, false);

                        var childNode = new TreeNode(GetNameForNode(formModel))
                        {
                            Tag = formModel,
                            ToolTipText = GetTooltip(formModel)
                        };
                        if (formModel.Form.Controls.Count > 0)
                            childNode.Nodes.Add("");
                        treeNode.Nodes.Add(childNode);
                    }
                    else
                    {
                        var formModel = new AccessObjectModel(form);
                        var childNode = new TreeNode(GetNameForNode(formModel))
                        {
                            Tag = formModel,
                            ToolTipText = GetTooltip(formModel)
                        };
                        treeNode.Nodes.Add(childNode);
                    }
                }
            }
        }

        private void LoadControlNodes(TreeNode treeNode, AccessFormModel accessFormModel)
        {
            if (treeNode.Nodes.Count == 1 && treeNode.Nodes[0].Text == "")
            {
                treeNode.Nodes.Clear();

                accessFormModel.LoadControls(false);

                foreach (var childModel in accessFormModel.Controls)
                {
                    var childNode = new TreeNode(GetNameForNode(childModel))
                    {
                        Tag = childModel,
                        ToolTipText = GetTooltip(childModel)
                    };

                    if (childModel.GetChildrenCount() > 0)
                        childNode.Nodes.Add("");
                    treeNode.Nodes.Add(childNode);
                }

            }
        }


        private void LoadControlNodes(TreeNode treeNode, AccessControlModel accessControlModel)
        {
            if (treeNode.Nodes.Count == 1 && treeNode.Nodes[0].Text == "")
            {
                treeNode.Nodes.Clear();

                accessControlModel.LoadChildren(false);

                foreach (var childModel in accessControlModel.Children)
                {
                    var childNode = new TreeNode(GetNameForNode(childModel))
                    {
                        Tag = childModel,
                        ToolTipText = GetTooltip(childModel)
                    };

                    if (childModel.GetChildrenCount() > 0)
                        childNode.Nodes.Add("");
                    treeNode.Nodes.Add(childNode);
                }

            }
        }

        private AccessControlModel GetAccessControlModelFromNode(TreeNode node)
        {
            if (node?.Tag is AccessControlModel accessControlModel)
                return accessControlModel;

            return null;
        }

        public void InsertPathIntoScript(TreeNode node)
        {
            if (node != null && node.Tag is AccessControlModel accessControlModel)
            {
                var path = "";
                while (node != null)
                {
                    var model = GetAccessControlModelFromNode(node);
                    path = $"{model.Name}/{path}";
                    node = node.Parent;
                }
                path = "/" + path;

                if (mainForm == null)
                    MessageBox.Show(path);
                else
                    mainForm.InsertTextIntoCurrentEditor($"{SpecialChars.Text}{path}{SpecialChars.Text}");
            }
        }

        public void ShowMarkerForm(TreeNode treeNode)
        {
            if (treeNode != null)
            {
                var node = (AccessControlModel)treeNode.Tag;
                node.Blink();
            }
        }

    }
}
