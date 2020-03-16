using G1ANT.Addon.MSOffice.Api.Access;
using G1ANT.Addon.MSOffice.Controllers.Access;
using G1ANT.Addon.MSOffice.Models.Access;
using G1ANT.Addon.MSOffice.Models.Access.Dao;
using G1ANT.Language;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
//using Access = Microsoft.Office.Interop.Access;

namespace G1ANT.Addon.MSOffice.Controllers
{
    public class AccessUIControlsTreeController
    {
        private IMainForm mainForm;
        public bool initialized = false;
        private readonly IRunningObjectTableService runningObjectTableService;
        private readonly ITooltipService tooltipService;
        private readonly TreeView controlsTree;
        private readonly ComboBox applications;
        public List<object> expandedTreeNodeModels = new List<object>();
        private object selectedTreeNodeModel;

        public AccessUIControlsTreeController(
            TreeView controlsTree, ComboBox applications,
            IRunningObjectTableService runningObjectTableService,
            ITooltipService tooltipService
            )
        {
            this.runningObjectTableService = runningObjectTableService;
            this.tooltipService = tooltipService;
            this.controlsTree = controlsTree;
            this.applications = applications;
        }

        public void Initialize(IMainForm mainForm) => this.mainForm = mainForm;


        public void InitRootElements(bool force = false)
        {
            if (initialized && !force)
                return;
            initialized = true;

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

        const string FormsLabel = "Forms";
        const string MacrosLabel = "Macros";
        const string ReportsLabel = "Reports";
        const string QueriesLabel = "Queries";
        const string PropertiesLabel = "Properties";
        const string InternalName = "internal";

        internal void SelectedApplicationChanged(RotApplicationModel rotApplicationModel)
        {
            //rotApplicationModel.Application.da

            if (rotApplicationModel.Application == null)
            {
                controlsTree.Nodes.Clear();
                return;
            }

            selectedTreeNodeModel = controlsTree.SelectedNode?.Tag;
            CollectExpandedTreeNodeModels(controlsTree.Nodes);

            controlsTree.BeginUpdate();
            controlsTree.Nodes.Clear();

            controlsTree.Nodes.AddRange(
                new TreeNode[]
                {
                    new TreeNode(FormsLabel) { Tag = rotApplicationModel, Nodes = { "" } },
                    new TreeNode(MacrosLabel) { Tag = rotApplicationModel, Nodes = { "" } },
                    new TreeNode(ReportsLabel) { Tag = rotApplicationModel, Nodes = { "" } },
                    new TreeNode(QueriesLabel) { Tag = rotApplicationModel, Nodes = { "" } },
                }
            );

            ApplyExpandedTreeNodes(controlsTree.Nodes);
            controlsTree.EndUpdate();
        }



        internal void LoadForm(AccessObjectModel formToLoad, bool openInDesigner)
        {
            var applicationModel = GetCurrentApplication();
            new Thread(() =>
            {
                try { OpenForm(formToLoad, openInDesigner, applicationModel); }
                catch (COMException ex) { RobotMessageBox.Show(ex.Message); }
            }).Start();
        }

        private RotApplicationModel GetCurrentApplication()
        {
            return (RotApplicationModel)applications.SelectedItem;
        }

        private static void OpenForm(AccessObjectModel formToLoad, bool openInDesigner, RotApplicationModel applicationModel)
        {
            try
            {
                var formName = formToLoad.FullName ?? formToLoad.Name;
                applicationModel.Application.DoCmd.OpenForm(
                    formName,
                    openInDesigner ? Microsoft.Office.Interop.Access.AcFormView.acDesign : Microsoft.Office.Interop.Access.AcFormView.acNormal
                );

                var form = applicationModel.Application.Forms[formName];
                form.SetFocus();
                RobotWin32.BringWindowToFront((IntPtr)form.Hwnd);
            }
            catch (Exception ex)
            {
                RobotMessageBox.Show(ex.Message);
            }
        }


        internal void OpenReport(AccessObjectModel report)
        {
            try
            {
                var app = GetCurrentApplication().Application;
                app.DoCmd.OpenReport(report.FullName ?? report.Name);
                RobotWin32.BringWindowToFront((IntPtr)app.hWndAccessApp());
            }
            catch (Exception ex)
            {
                RobotMessageBox.Show(ex.Message);
            }
        }

        internal void ExecuteQuery(AccessObjectModel model)
        {
            try
            {
                var app = GetCurrentApplication().Application;
                app.DoCmd.OpenQuery(
                    model.FullName ?? model.Name,
                    Microsoft.Office.Interop.Access.AcView.acViewNormal,
                    Microsoft.Office.Interop.Access.AcOpenDataMode.acReadOnly
                );
                RobotWin32.BringWindowToFront((IntPtr)app.hWndAccessApp());
            }
            catch (Exception ex)
            {
                RobotMessageBox.Show(ex.Message);
            }
        }

        public void CopyNodeDetails(TreeNode treeNode)
        {
            Clipboard.SetText(treeNode.ToolTipText);
        }

        private static string GetNameForNode(AccessControlModel model)
        {
            return $"{model.Name} {model.Type} {model.Value}";
        }

        private static string GetNameForNode(AccessFormModel model)
        {
            return $"{model.Name} {(model.Name != model.Caption ? model.Caption : "")} {(model.Name != model.FormName ? model.FormName : "")}";
        }

        private string GetNameForNode(AccessObjectModel model)
        {
            return $"{model.Name} {(model.Name != model.FullName ? model.FullName : "")}";// {(model.IsLoaded ? "" : "(not loaded)")}";
        }

        private string GetNameForNode(AccessQueryModel model)
        {
            return $"{model.Name} {model.Type}";
        }


        private void CollectExpandedTreeNodeModels(TreeNodeCollection nodes)
        {
            var expandedNodes = nodes.Cast<TreeNode>()
                .Where(tn => tn.IsExpanded)
                .ToList();
            expandedNodes.ForEach(en =>
            {
                expandedTreeNodeModels.Add(en.Tag);
                CollectExpandedTreeNodeModels(en.Nodes);
            });
        }


        private bool AreModelsSame(object source, object dest)
        {
            if (source == null || dest == null)
                return false;

            if (source is IComparable sc)
                return sc.CompareTo(dest) == 0;

            return false;
        }


        private void ApplyExpandedTreeNodes(TreeNodeCollection nodes)
        {
            foreach (TreeNode node in nodes)
            {
                var nodeModel = node.Tag;

                if (selectedTreeNodeModel != null && AreModelsSame(nodeModel, selectedTreeNodeModel))
                    controlsTree.SelectedNode = node;

                var expandedTreeNodeModel = expandedTreeNodeModels.FirstOrDefault(etn => AreModelsSame(etn, nodeModel));
                if (expandedTreeNodeModel != null)
                {
                    expandedTreeNodeModels.Remove(expandedTreeNodeModel);

                    LoadChildNodes(node);
                    node.Expand();

                    ApplyExpandedTreeNodes(node.Nodes);
                    expandedTreeNodeModels.Remove(expandedTreeNodeModel);
                }
            }
        }

        public void LoadChildNodes(TreeNode treeNode)
        {
            controlsTree.BeginUpdate();

            if (treeNode.Tag is AccessControlModel accessControlModel)
                LoadControlNodes(treeNode, accessControlModel);
            else if (treeNode.Tag is AccessFormModel accessFormModel)
                LoadControlNodes(treeNode, accessFormModel);
            else if (treeNode.Tag is RotApplicationModel rotApplicationModel)
            {
                switch (treeNode.Text)
                {
                    case FormsLabel:
                        LoadFormNodes(treeNode, rotApplicationModel);
                        break;
                    case MacrosLabel:
                        LoadAccessObjectNodes(treeNode, new AccessObjectMacroCollectionModel(rotApplicationModel));
                        break;
                    case QueriesLabel:
                        LoadAccessObjectNodes(treeNode, new AccessObjectQueryCollectionModel(rotApplicationModel));
                        //LoadQueryNodes(treeNode, rotApplicationModel);
                        break;
                    case ReportsLabel:
                        LoadAccessObjectNodes(treeNode, new AccessObjectReportCollectionModel(rotApplicationModel));
                        break;
                }
            }

            ApplyExpandedTreeNodes(treeNode.Nodes);
            controlsTree.EndUpdate();
        }


        private void LoadAccessObjectNodes(TreeNode treeNode, AccessObjectCollectionModel objects)
        {
            if (IsEmptyNode(treeNode))
            {
                treeNode.Nodes.Clear();

                foreach (var @object in objects)
                {
                    var childNode = new TreeNode(GetNameForNode(@object))
                    {
                        Tag = @object,
                        ToolTipText = tooltipService.GetTooltip(@object)
                    };
                    treeNode.Nodes.Add(childNode);
                }
            }
        }


        //private void LoadQueryNodes(TreeNode treeNode, RotApplicationModel rotApplicationModel)
        //{
        //    if (IsEmptyNode(treeNode))
        //    {
        //        treeNode.Nodes.Clear();

        //        //var queries = new AccessQueryCollectionModel(rotApplicationModel);
        //        var queries = new AccessObjectQueryCollectionModel(rotApplicationModel);
        //        foreach (var query in queries)
        //        {
        //            var childNode = new TreeNode(GetNameForNode(query))
        //            {
        //                Tag = query,
        //                ToolTipText = tooltipService.GetTooltip(query)
        //            };
        //            treeNode.Nodes.Add(childNode);
        //        }
        //    }
        //}




        private TreeNode CreatePropertyNode(object model)
        {
            var node = new TreeNode(PropertiesLabel)
            {
                Tag = model,
                Name = InternalName,
                Nodes = { "" }
            };

            return node;
        }

        private bool IsEmptyNode(TreeNode node)
        {
            return node.Nodes.Count == 1 && node.Nodes[0].Text == "";
        }


        private void LoadFormNodes(TreeNode treeNode, RotApplicationModel rotApplicationModel)
        {
            if (IsEmptyNode(treeNode))
            {
                treeNode.Nodes.Clear();

                var formAccessObjects = new AccessObjectFormCollectionModel(rotApplicationModel);
                foreach (var accessObject in formAccessObjects)
                {
                    if (accessObject.IsLoaded)
                    {
                        var formModel = new AccessFormModel(rotApplicationModel.Application.Forms[accessObject.Name], false, false, false);

                        var childNode = new TreeNode(GetNameForNode(formModel))
                        {
                            Tag = formModel,
                            ToolTipText = tooltipService.GetTooltip(formModel)
                        };
                        if (formModel.Form.Controls.Count > 0)
                            childNode.Nodes.Add("");
                        treeNode.Nodes.Add(childNode);
                    }
                    else
                    {
                        var childNode = new TreeNode(GetNameForNode(accessObject))
                        {
                            Tag = accessObject,
                            ToolTipText = tooltipService.GetTooltip(accessObject)
                        };
                        treeNode.Nodes.Add(childNode);
                    }
                }
            }
        }

        private void LoadControlNodes(TreeNode treeNode, AccessFormModel accessFormModel)
        {
            if (IsEmptyNode(treeNode))
            {
                treeNode.Nodes.Clear();

                if (treeNode.Name == InternalName && treeNode.Text == PropertiesLabel)
                {
                    treeNode.Nodes.AddRange(CreatePropertyNodes(accessFormModel.Form.Properties));
                    return;
                }

                treeNode.Nodes.Add(CreatePropertyNode(accessFormModel));

                accessFormModel.LoadControls(false);

                foreach (var childModel in accessFormModel.Controls)
                {
                    var childNode = new TreeNode(GetNameForNode(childModel))
                    {
                        Tag = childModel,
                        ToolTipText = tooltipService.GetTooltip(childModel),
                        Nodes = { "" }
                    };

                    treeNode.Nodes.Add(childNode);
                }

            }
        }

        private TreeNode[] CreatePropertyNodes(Microsoft.Office.Interop.Access.Properties properties)
        {
            return new AccessPropertiesModel(properties)
                .Select(p => new TreeNode($"{p.Key}: {p.Value}"))
                .ToArray();
        }

        private void LoadControlNodes(TreeNode treeNode, AccessControlModel accessControlModel)
        {
            if (IsEmptyNode(treeNode))
            {
                treeNode.Nodes.Clear();

                if (treeNode.Name == InternalName && treeNode.Text == PropertiesLabel)
                {
                    treeNode.Nodes.AddRange(CreatePropertyNodes(accessControlModel.Control.Properties));
                    return;
                }

                treeNode.Nodes.Add(CreatePropertyNode(accessControlModel));

                if (accessControlModel.GetChildrenCount() > 0)
                {
                    accessControlModel.LoadChildren(false);

                    foreach (var childModel in accessControlModel.Children)
                    {
                        var childNode = new TreeNode(GetNameForNode(childModel))
                        {
                            Tag = childModel,
                            ToolTipText = tooltipService.GetTooltip(childModel),
                            Nodes = { "" }
                        };

                        treeNode.Nodes.Add(childNode);
                    }
                }
            }
        }


        private string GetNameFromNodeModel(TreeNode node)
        {
            if (node.Tag is AccessControlModel accessControlModel)
                return accessControlModel.Name;

            if (node.Tag is AccessFormModel accessFormModel)
                return accessFormModel.Name;

            return null;
        }

        public void InsertPathIntoScript(TreeNode node)
        {
            if (node != null && node.Tag is AccessControlModel accessControlModel)
            {
                var path = "";
                while (node != null)
                {
                    var name = GetNameFromNodeModel(node);
                    if (name == null)
                        break;
                    path = $"{name}/{path}";
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
            if (treeNode?.Tag is AccessControlModel model)
            {
                model.Blink();
            }
            else if (treeNode?.Tag == null && treeNode?.Parent?.Tag is AccessControlModel parentModel)
            {
                parentModel.Blink();
            }
        }

    }
}
