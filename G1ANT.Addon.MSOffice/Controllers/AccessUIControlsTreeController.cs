using G1ANT.Addon.MSOffice.Api.Access;
using G1ANT.Addon.MSOffice.Controllers.Access;
using G1ANT.Addon.MSOffice.Models.Access;
using G1ANT.Addon.MSOffice.Models.Access.Dao;
using G1ANT.Addon.MSOffice.Models.Access.Data;
using G1ANT.Addon.MSOffice.Models.Access.VBE;
using G1ANT.Language;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using MSAccess = Microsoft.Office.Interop.Access;

namespace G1ANT.Addon.MSOffice.Controllers
{
    public class AccessUIControlsTreeController
    {
        private const string ApplicationLabel = "Application";
        private const string FormsLabel = "Forms";
        private const string MacrosLabel = "Macros";
        private const string ReportsLabel = "Reports";
        private const string ResourcesLabel = "Resources";
        private const string ModulesLabel = "Modules";

        private const string DatabaseLabel = "Database";
        private const string DatabaseDiagramsLabel = "Database Diagrams";
        private const string FunctionsLabel = "Functions";
        private const string QueriesLabel = "Queries";
        private const string StoredProceduresLabel = "Stored Prodecures";
        private const string TablesLabel = "Tables";
        private const string ViewsLabel = "Views";

        private const string PropertiesLabel = "Properties";
        private const string DynamicPropertiesLabel = "Dynamic Properties";

        private const string InternalName = "internal";

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

            //controlsTree.DrawMode = TreeViewDrawMode.OwnerDrawText;
            //controlsTree.DrawNode += controlsTree_DrawNode;
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

        //private void controlsTree_DrawNode(object sender, DrawTreeNodeEventArgs e)
        //{
        //    if (e.Node.Text == "")
        //        return;

        //    if (e.Node.Tag is AccessControlModel accessControlModel)
        //    {
        //        var boldText = accessControlModel.Caption ?? "";
        //        var normalText = $"{accessControlModel.Name} {accessControlModel.Type}";

        //        using (Font font = new Font(controlsTree.Font, FontStyle.Bold))
        //        {
        //            using (Brush brush = new SolidBrush(controlsTree.ForeColor))
        //            {
        //                e.Graphics.DrawString(boldText, font, brush, e.Bounds.Left, e.Bounds.Top);

        //                var s = e.Graphics.MeasureString(boldText, controlsTree.Font);
        //                e.Graphics.DrawString(normalText, controlsTree.Font, brush, e.Bounds.Left + (int)s.Width + 10, e.Bounds.Top);
        //            }
        //        }
        //    }
        //    else
        //    {
        //        e.DrawDefault = true;
        //    }
        //}

        internal void SelectedApplicationChanged(RotApplicationModel rotApplicationModel)
        {
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
                    new TreeNode(ApplicationLabel) { Tag = rotApplicationModel, Nodes = { "" } },
                    new TreeNode(FormsLabel) { Tag = rotApplicationModel, Nodes = { "" } },
                    new TreeNode(MacrosLabel) { Tag = rotApplicationModel, Nodes = { "" } },
                    new TreeNode(ReportsLabel) { Tag = rotApplicationModel, Nodes = { "" } },
                    new TreeNode(ResourcesLabel) { Tag = rotApplicationModel, Nodes = { "" } },
                    new TreeNode(ModulesLabel) { Tag = rotApplicationModel, Nodes = { "" } },
                    new TreeNode(DatabaseLabel) {
                        Nodes = {
                            new TreeNode(DatabaseDiagramsLabel) { Tag = rotApplicationModel, Nodes = { "" } },
                            new TreeNode(FunctionsLabel) { Tag = rotApplicationModel, Nodes = { "" } },
                            new TreeNode(QueriesLabel) { Tag = rotApplicationModel, Nodes = { "" } },
                            new TreeNode(StoredProceduresLabel) { Tag = rotApplicationModel, Nodes = { "" } },
                            new TreeNode(TablesLabel) { Tag = rotApplicationModel, Nodes = { "" } },
                            new TreeNode(ViewsLabel) { Tag = rotApplicationModel, Nodes = { "" } },
                        },
                    }
                }
            );

            ApplyExpandedTreeNodes(controlsTree.Nodes);
            controlsTree.EndUpdate();
        }



        internal void TryOpenForm(AccessObjectModel formToLoad, bool openInDesigner)
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
                    openInDesigner ? MSAccess.AcFormView.acDesign : MSAccess.AcFormView.acNormal
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


        internal void OpenAccessObject(AccessObjectModel report)
        {
            try
            {
                var app = GetCurrentApplication().Application;
                var name = report.FullName ?? report.Name;

                switch (report.Type)
                {
                    case MSAccess.AcObjectType.acReport:
                        app.DoCmd.OpenReport(name);
                        break;
                    case MSAccess.AcObjectType.acTable:
                        app.DoCmd.OpenTable(name, MSAccess.AcView.acViewNormal, MSAccess.AcOpenDataMode.acReadOnly);
                        break;
                    case MSAccess.AcObjectType.acServerView:
                        app.DoCmd.OpenView(name, MSAccess.AcView.acViewNormal, MSAccess.AcOpenDataMode.acReadOnly);
                        break;
                    case MSAccess.AcObjectType.acStoredProcedure:
                        app.DoCmd.OpenStoredProcedure(name, MSAccess.AcView.acViewNormal, MSAccess.AcOpenDataMode.acReadOnly);
                        break;
                    case MSAccess.AcObjectType.acQuery:
                        app.DoCmd.OpenQuery(name, MSAccess.AcView.acViewNormal, MSAccess.AcOpenDataMode.acReadOnly);
                        break;
                    case MSAccess.AcObjectType.acFunction:
                        app.DoCmd.OpenFunction(name, MSAccess.AcView.acViewNormal, MSAccess.AcOpenDataMode.acReadOnly);
                        break;
                    case MSAccess.AcObjectType.acDiagram:
                        app.DoCmd.OpenDiagram(name);
                        break;
                    case MSAccess.AcObjectType.acMacro:
                        app.DoCmd.RunMacro(name, 1, true);
                        break;
                    default:
                        throw new NotImplementedException($"Opener for {report.TypeName} not implemented.");
                }

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
            return $"{model.Caption} {model.Name} {model.Type} {model.Value}";
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
            else if (treeNode.Tag is ModuleModel moduleModel)
                LoadModulePropertyNodes(treeNode, moduleModel);
            else if (treeNode.Tag is RotApplicationModel rotApplicationModel)
            {
                switch (treeNode.Text)
                {
                    case ApplicationLabel:
                        LoadApplicationNodes(treeNode, rotApplicationModel);
                        break;
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

                    case ResourcesLabel:
                        LoadResourceNodes(treeNode);
                        break;
                    case ModulesLabel:
                        LoadModuleNodes(treeNode);
                        break;

                    case DatabaseDiagramsLabel:
                        LoadAccessObjectNodes(treeNode, new AccessObjectDatabaseDiagramCollectionModel(rotApplicationModel));
                        break;
                    case FunctionsLabel:
                        LoadAccessObjectNodes(treeNode, new AccessObjectFunctionCollectionModel(rotApplicationModel));
                        break;
                    case StoredProceduresLabel:
                        LoadAccessObjectNodes(treeNode, new AccessObjectStoredProcedureCollectionModel(rotApplicationModel));
                        break;
                    case TablesLabel:
                        LoadAccessObjectNodes(treeNode, new AccessObjectTableCollectionModel(rotApplicationModel));
                        break;
                    case ViewsLabel:
                        LoadAccessObjectNodes(treeNode, new AccessObjectViewCollectionModel(rotApplicationModel));
                        break;
                }
            }

            ApplyExpandedTreeNodes(treeNode.Nodes);
            controlsTree.EndUpdate();
        }

        private void LoadApplicationNodes(TreeNode parentNode, RotApplicationModel rotApplicationModel)
        {
            if (IsEmptyNode(parentNode))
            {
                parentNode.Nodes.Clear();
                var app = rotApplicationModel.Application;

                var vbe = new VbeModel(app.VBE);
                var vbeProjectsNode = new TreeNode("Projects", vbe.Projects.Select(p => new TreeNode(p.ToString())).ToArray());
                var vbeWindowsNode = new TreeNode("Windows", vbe.Windows.Select(w => new TreeNode(w.ToString())).ToArray());
                var vbeAddinsNode = new TreeNode("Addins", vbe.Addins.Select(a => new TreeNode(a.ToString())).ToArray());

                var vbeNode = new TreeNode(
                    $"VBE Version: {vbe.Version}",
                    new TreeNode[] {
                        new TreeNode($"MainWindow: {vbe.MainWindow}"),
                        vbeWindowsNode,
                        new TreeNode($"Active project {vbe.ActiveVBProject}"),
                        vbeProjectsNode,
                        vbeAddinsNode
                    }
                );

                parentNode.Nodes.AddRange(
                    new TreeNode[]
                    {
                        new TreeNode($"Name: {app.Name}"),
                        new TreeNode($"Version: {app.Version}"),
                        new TreeNode($"BaseConnectionString: {app.CurrentProject.BaseConnectionString}"),
                        new TreeNode($"Current Object Name: {app.CurrentObjectName}"),
                        new TreeNode($"Current Object Type: {app.CurrentObjectType}"),
                        vbeNode,
                    }
                );

                var tempVars = new TempVarCollectionModel(app.TempVars);
                if (tempVars.Any())
                    parentNode.Nodes.Add(new TreeNode($"Temp Vars", tempVars.Select(t => new TreeNode(t.ToString())).ToArray()));
            }
        }

        private void LoadModulePropertyNodes(TreeNode parentNode, ModuleModel model)
        {
            if (IsEmptyNode(parentNode))
            {
                parentNode.Nodes.Clear();

                parentNode.Nodes.AddRange(
                    new TreeNode[]
                    {
                        new TreeNode($"Name: {model.Name}"),
                        new TreeNode($"Type: {model.TypeName}"),
                        new TreeNode($"CountOfDeclarationLines: {model.CountOfDeclarationLines}"),
                        new TreeNode($"CountOfLines: {model.CountOfLines}"),
                        new TreeNode($"Code: {model.Code}"), // ???
                    }
                );
            }
        }

        private void LoadModuleNodes(TreeNode parentNode)
        {
            if (IsEmptyNode(parentNode))
            {
                parentNode.Nodes.Clear();
                var modules = GetCurrentApplication().Application.Modules;// CurrentProject.AllModules;

                parentNode.Nodes.AddRange(
                    modules.Cast<MSAccess.Module>()
                        .Select(m => new ModuleModel(m))
                        .Select(mm => new TreeNode(mm.ToString()) { Tag = mm, Nodes = { CreatePropertyNode(mm) } })
                        .ToArray()
                );
            }
        }

        private void LoadResourceNodes(TreeNode parentNode)
        {
            if (IsEmptyNode(parentNode))
            {
                parentNode.Nodes.Clear();
                var resources = GetCurrentApplication().Application.CurrentProject.Resources;

                parentNode.Nodes.AddRange(
                    resources.Cast<MSAccess.SharedResource>()
                        .Select(r => new ResourceModel(r))
                        .Select(rm => new TreeNode(rm.ToString()) { Tag = rm })
                        .ToArray()
                );
            }
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
                foreach (var accessObject in formAccessObjects.OrderByDescending(f => f.IsLoaded).ThenBy(f => f.Name))
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
                    treeNode.Nodes.Add(new TreeNode(DynamicPropertiesLabel, CreateDynamicPropertyNodes(accessFormModel.Form.Properties)));
                    treeNode.Nodes.AddRange(CreatePropertyNodes(accessFormModel.Form));
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

        private void LoadControlNodes(TreeNode treeNode, AccessControlModel accessControlModel)
        {
            if (IsEmptyNode(treeNode))
            {
                treeNode.Nodes.Clear();

                if (treeNode.Name == InternalName && treeNode.Text == PropertiesLabel)
                {
                    treeNode.Nodes.Add(new TreeNode(DynamicPropertiesLabel, CreateDynamicPropertyNodes(accessControlModel.Control.Properties)));
                    treeNode.Nodes.AddRange(CreatePropertyNodes(accessControlModel.Control));
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

        private TreeNode[] CreateDynamicPropertyNodes(MSAccess.Properties properties)
        {
            return new AccessDynamicPropertiesModel(properties)
                .OrderBy(p => p.Key)
                .Select(p => new TreeNode($"{p.Key}: {p.Value}"))
                .ToArray();
        }

        private TreeNode[] CreatePropertyNodes(object accessObject)
        {
            var objectProperties = TypeDescriptor.GetProperties(accessObject);
            return objectProperties
                .Cast<PropertyDescriptor>()
                .OrderBy(p => p.Name)
                .Select(p => new TreeNode($"{p.Name}: {p.GetValue(accessObject)}"))
                .ToArray();
        }

        private string GetNameFromNodeModel(TreeNode node)
        {
            if (node.Tag is INameModel nameModel)
                return nameModel.Name;

            return null;
        }

        public void InsertPathIntoScript(TreeNode node)
        {
            var path = "";
            if (node == null)
                return;

            if (node.Tag is AccessControlModel accessControlModel)
            {
                while (node != null)
                {
                    var name = GetNameFromNodeModel(node);
                    if (name == null)
                        break;
                    path = $"{name}/{path}";
                    node = node.Parent;
                }
                path = "/" + path;
            }
            else if (node.Tag is INameModel nameModel)
            {
                path = nameModel.Name;
            }

            if (mainForm == null)
                MessageBox.Show(path);
            else
                mainForm.InsertTextIntoCurrentEditor($"{SpecialChars.Text}{path}{SpecialChars.Text}");
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
