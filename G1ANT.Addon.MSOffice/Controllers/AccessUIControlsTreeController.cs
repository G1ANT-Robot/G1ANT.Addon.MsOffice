using G1ANT.Addon.MSOffice.Api.Access;
using G1ANT.Addon.MSOffice.Controllers.Access;
using G1ANT.Addon.MSOffice.Forms;
using G1ANT.Addon.MSOffice.Models.Access;
using G1ANT.Addon.MSOffice.Models.Access.Dao;
using G1ANT.Addon.MSOffice.Models.Access.Data;
using G1ANT.Addon.MSOffice.Models.Access.VBE;
using G1ANT.Language;
using Microsoft.Office.Interop.Access.Dao;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data.OleDb;
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
        public List<TreeNode> expandedTreeNodes = new List<TreeNode>();
        private TreeNode selectedTreeNode;

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

            selectedTreeNode = controlsTree.SelectedNode;
            CollectExpandedTreeNodeModels(controlsTree.Nodes);

            controlsTree.BeginUpdate();
            controlsTree.Nodes.Clear();

            controlsTree.Nodes.AddRange(
                new TreeNode[]
                {
                    new LazyTreeNode(ApplicationLabel, () => GetApplicationNodes(rotApplicationModel)) { Tag = rotApplicationModel },
                    new LazyTreeNode(FormsLabel, () => GetFormNodes(rotApplicationModel)) { Tag = rotApplicationModel },
                    new LazyTreeNode(MacrosLabel, () => GetAccessObjectNodes(new AccessObjectMacroCollectionModel(rotApplicationModel))) { Tag = rotApplicationModel },
                    new LazyTreeNode(ReportsLabel, () => GetAccessObjectNodes(new AccessObjectReportCollectionModel(rotApplicationModel))) { Tag = rotApplicationModel },
                    new LazyTreeNode(ResourcesLabel, () => GetResourceNodes()) { Tag = rotApplicationModel },
                    new LazyTreeNode(ModulesLabel, () => GetModuleNodes() ) { Tag = rotApplicationModel },
                    new TreeNode(DatabaseLabel) {
                        Nodes = {
                            new LazyTreeNode(
                                DatabaseDiagramsLabel,
                                () => GetAccessObjectNodes(new AccessObjectDatabaseDiagramCollectionModel(rotApplicationModel))
                            ) { Tag = rotApplicationModel },
                            new LazyTreeNode(
                                FunctionsLabel,
                                () => GetAccessObjectNodes(new AccessObjectFunctionCollectionModel(rotApplicationModel))) { Tag = rotApplicationModel },
                            new LazyTreeNode(
                                QueriesLabel,
                                () => GetQueryNodes(rotApplicationModel)) { Tag = rotApplicationModel },
                            new LazyTreeNode(
                                StoredProceduresLabel,
                                () => GetAccessObjectNodes(new AccessObjectStoredProcedureCollectionModel(rotApplicationModel))
                            ) { Tag = rotApplicationModel },
                            new LazyTreeNode(
                                TablesLabel,
                                () => GetTableDefNodes(rotApplicationModel.Application.CurrentDb())
                            ) { Tag = rotApplicationModel },
                            new LazyTreeNode(
                                ViewsLabel,
                                () => GetAccessObjectNodes(new AccessObjectViewCollectionModel(rotApplicationModel))
                            ) { Tag = rotApplicationModel },
                        },
                    }
                }
            );

            ApplyExpandedTreeNodes(controlsTree.Nodes);
            controlsTree.EndUpdate();
        }

        internal void ViewDataFromTable(string tableName)
        {
            try
            {
                var app = GetCurrentApplication().Application;
                var form = new DataTableForm();
                using (var connection = new OleDbConnection(app.ADOConnectString))
                {
                    form.LoadData(connection, tableName);
                    form.ShowDialog();
                }
            }
            catch (Exception ex)
            {
                RobotMessageBox.Show($"Exception while loading data from {tableName}: {ex.Message}");
            }
        }

        internal void TryOpenFormFromSelectedNode(bool openInDesigner)
        {
            var selectedNode = controlsTree.SelectedNode;
            var model = (AccessObjectModel)selectedNode.Tag;

            var applicationModel = GetCurrentApplication();
            new Thread(() =>
            {
                try
                {
                    OpenForm(model, openInDesigner, applicationModel);
                    var newNode = GetLoadedFormNode(applicationModel.Application, selectedNode.Text);
                    controlsTree.FindForm().Invoke((MethodInvoker) delegate { ReplaceNode(selectedNode, newNode); });

                }
                catch (COMException ex) { RobotMessageBox.Show(ex.Message); }
            }).Start();
        }

        private void ReplaceNode(TreeNode oldNode, TreeNode newNode)
        {
            var isSelected = oldNode.IsSelected;
            oldNode.Parent.Nodes[oldNode.Index] = newNode;
            oldNode.Remove();

            if (isSelected)
                controlsTree.SelectedNode = newNode;
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


        private void CollectExpandedTreeNodeModels(TreeNodeCollection nodes)
        {
            var expandedNodes = nodes.Cast<TreeNode>()
                .Where(tn => tn.IsExpanded)
                .ToList();

            expandedTreeNodes.AddRange(expandedNodes);
            expandedNodes.ForEach(en => CollectExpandedTreeNodeModels(en.Nodes));
        }


        private bool AreNodesSame(TreeNode sourceNode, TreeNode destNode)
        {
            return sourceNode.Text == destNode.Text && AreModelsSame(sourceNode.Tag, destNode.Tag);
        }


        private bool AreModelsSame(object source, object dest)
        {
            if (source == null && dest == null)
                return true;

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
                if (selectedTreeNode != null && AreNodesSame(node, selectedTreeNode))
                {
                    controlsTree.SelectedNode = node;
                    selectedTreeNode = null;
                }

                var expandedTreeNode = expandedTreeNodes.FirstOrDefault(etn => AreNodesSame(etn, node));
                if (expandedTreeNode != null)
                {
                    expandedTreeNodes.Remove(expandedTreeNode);

                    TryLoadChildNodes(node);
                    node.Expand();

                    if (!IsEmptyNode(node))
                        ApplyExpandedTreeNodes(node.Nodes);
                }
            }
        }

        public void TryLoadChildNodes(TreeNode treeNode)
        {
            try
            {
                LoadChildNodes(treeNode);
            }
            catch (Exception ex)
            {
                treeNode.Nodes.Add($"Exception while loading node contents: {ex.Message}");
            }
        }

        public void LoadChildNodes(TreeNode treeNode)
        {
            controlsTree.BeginUpdate();
            var oldCursor = Cursor.Current;
            Cursor.Current = Cursors.WaitCursor;

            if (treeNode is LazyTreeNode lazyTreeNode)
                lazyTreeNode.LoadLazyChildren();

            ApplyExpandedTreeNodes(treeNode.Nodes);
            controlsTree.EndUpdate();
            Cursor.Current = oldCursor;
        }


        private TreeNode[] GetObjectPropertiesAsTreeNodes(object @object)
        {
            return @object
                .GetType()
                .GetProperties()
                .Select(p => new TreeNode($"{p.Name}: {p.GetValue(@object)}"))
                .ToArray();
        }

        private IEnumerable<TreeNode> GetQueryDetailsPropertyNodes(AccessQueryModel accessQueryModel)
        {
            var details = accessQueryModel.Details.Value;

            return new TreeNode[]
            {
                new LazyTreeNode(
                    "Fields",
                    () => details.Fields.Select(f => new LazyTreeNode(f.Name, () => GetObjectPropertiesAsTreeNodes(f)))
                ),
                new LazyTreeNode(
                    "Parameters",
                    () => details.Parameters.Select(p => new TreeNode($"{p.Name}: {p.Value}, type: {p.Type}"))
                ),
                new LazyTreeNode(
                    "Properties",
                    () => details.Properties.Select(p => new TreeNode($"{p.Name}: {p.Value}, type: {p.PropertyType}"))
                ),
                new TreeNode($"SQL: {details.SQL}"),
                //new TreeNode($"Name: {details.Name}"),
                new TreeNode($"DateCreated: {details.DateCreated}"),
                new TreeNode($"LastUpdated: {details.LastUpdated}"),
                new TreeNode($"Connect: {details.Connect}"),
                new TreeNode($"MaxRecords: {details.MaxRecords}"),
                new TreeNode($"RecordsAffected: {details.RecordsAffected}"),
                new TreeNode($"ReturnsRecords: {details.ReturnsRecords}"),
                new TreeNode($"Type: {details.Type}"),
                new TreeNode($"Updatable: {details.Updatable}")
            };
        }

        private IEnumerable<TreeNode> GetTableDefNodes(Database database)
        {
            var tableDefs = new AccessTableDefCollectionModel(database.TableDefs);

            return tableDefs.Select(td => new TreeNode(
                td.ToString(),
                new TreeNode[] {
                    new LazyTreeNode(
                        "Fields",
                        () => td.Fields.Value.Select(f => new LazyTreeNode(f.Name, () => GetObjectPropertiesAsTreeNodes(f)))
                    ),
                    new LazyTreeNode(
                        "Properties",
                        () => td.Properties.Value.Select(p => new LazyTreeNode(p.Name, () => GetObjectPropertiesAsTreeNodes(p)))
                    ),
                    new LazyTreeNode(
                        "Indexes",
                        () => td.Indexes.Value.Select(i => new LazyTreeNode(i.Name, () => GetObjectPropertiesAsTreeNodes(i)))
                    ),
                    new TreeNode($"DateCreated: {td.DateCreated}"),
                    new TreeNode($"LastUpdated: {td.LastUpdated}"),
                    new TreeNode($"Connect: {td.Connect}"),
                    new TreeNode($"RecordCount: {td.RecordCount}"),
                    new TreeNode($"SourceTableName: {td.SourceTableName}"),
                    new TreeNode($"Updatable: {td.Updatable}"),
                }
            )
            { Tag = td });
        }

        private List<TreeNode> GetApplicationNodes(RotApplicationModel rotApplicationModel)
        {
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

            var result = new List<TreeNode>()
            {
                new TreeNode($"Name: {app.Name}"),
                new TreeNode($"Version: {app.Version}"),
                new TreeNode($"ADOConnectString: {app.ADOConnectString}"),
                new TreeNode($"BaseConnectionString: {app.CurrentProject.BaseConnectionString}"),
                new TreeNode($"Current Object Name: {app.CurrentObjectName}"),
                new TreeNode($"Current Object Type: {app.CurrentObjectType}"),
                vbeNode,
            };


            var tempVars = new TempVarCollectionModel(app.TempVars);
            if (tempVars.Any())
                result.Add(new TreeNode($"Temp Vars", tempVars.Select(t => new TreeNode(t.ToString())).ToArray()));

            return result;
        }

        private IEnumerable<TreeNode> GetModulePropertyNodes(ModuleModel model)
        {
            return new TreeNode[]
            {
                new TreeNode($"Name: {model.Name}"),
                new TreeNode($"Type: {model.TypeName}"),
                new TreeNode($"CountOfDeclarationLines: {model.CountOfDeclarationLines}"),
                new TreeNode($"CountOfLines: {model.CountOfLines}"),
                new TreeNode($"Code: {model.Code}"), // ???
            };
        }

        private IEnumerable<TreeNode> GetModuleNodes()
        {
            var modules = GetCurrentApplication().Application.Modules;

            return modules.Cast<MSAccess.Module>()
                .Select(m => new ModuleModel(m))
                .Select(mm => new LazyTreeNode(mm.ToString(), () => GetModulePropertyNodes(mm)) { Tag = mm })
                .ToArray();
        }


        private IEnumerable<TreeNode> GetResourceNodes()
        {
            var resources = GetCurrentApplication().Application.CurrentProject.Resources;

            return resources.Cast<MSAccess.SharedResource>()
                .Select(r => new ResourceModel(r))
                .Select(rm => new TreeNode(rm.ToString()) { Tag = rm });
        }


        private IEnumerable<TreeNode> GetAccessObjectNodes(AccessObjectCollectionModel objects)
        {
            return objects.Select(o => new TreeNode(o.ToString()) { Tag = o, ToolTipText = tooltipService.GetTooltip(o) });
        }


        private IEnumerable<TreeNode> GetQueryNodes(RotApplicationModel rotApplicationModel)
        {
            var queries = new AccessQueryCollectionModel(rotApplicationModel);
            return queries.Select(q => new LazyTreeNode(
                q.ToString(),
                () => GetQueryDetailsPropertyNodes(q)
            ) { Tag = q, ToolTipText = tooltipService.GetTooltip(q), });
        }

 
        private bool IsEmptyNode(TreeNode node)
        {
            return node.Nodes.Count == 1 && node.Nodes[0].Text == "";
        }

        private TreeNode GetLoadedFormNode(Microsoft.Office.Interop.Access.Application application, string formName)
        {
            var formModel = new AccessFormModel(application.Forms[formName], false, false, false);

            return new LazyTreeNode(formModel.ToString(), () => GetControlNodes(formModel))
            {
                Tag = formModel,
                ToolTipText = tooltipService.GetTooltip(formModel)
            };
        }

        private IEnumerable<TreeNode> GetFormNodes(RotApplicationModel rotApplicationModel)
        {
            var formAccessObjects = new AccessObjectFormCollectionModel(rotApplicationModel);

            foreach (var accessObject in formAccessObjects.OrderByDescending(f => f.IsLoaded).ThenBy(f => f.Name))
            {
                TreeNode childNode;

                if (accessObject.IsLoaded)
                {
                    yield return GetLoadedFormNode(rotApplicationModel.Application, accessObject.Name);
                }
                else
                {
                    yield return childNode = new TreeNode(accessObject.ToString()) {
                        Tag = accessObject,
                        ToolTipText = tooltipService.GetTooltip(accessObject)
                    };
                }
            }
        }


        private List<TreeNode> GetControlNodes(AccessFormModel accessFormModel)
        {
            var result = new List<TreeNode>();

            result.Add(
                new LazyTreeNode(
                    PropertiesLabel,
                    () => new List<TreeNode>() {
                        new LazyTreeNode(DynamicPropertiesLabel, () => CreateDynamicPropertyNodes(accessFormModel.Form.Properties))
                    }.Concat(CreatePropertyNodes(accessFormModel.Form))
                )
            );

            accessFormModel.LoadControls(false);

            foreach (var childModel in accessFormModel.Controls)
            {
                var childNode = new LazyTreeNode(childModel.ToString(), () => GetControlNodes(childModel))
                {
                    Tag = childModel,
                    ToolTipText = tooltipService.GetTooltip(childModel),
                };

                result.Add(childNode);
            }

            return result;
        }

        private List<TreeNode> GetControlNodes(AccessControlModel accessControlModel)
        {
            var result = new List<TreeNode>();

            result.Add(
                new LazyTreeNode(
                    PropertiesLabel,
                    () => new List<TreeNode>() {
                        new LazyTreeNode(DynamicPropertiesLabel, () => CreateDynamicPropertyNodes(accessControlModel.Control.Properties))
                    }.Concat(CreatePropertyNodes(accessControlModel.Control))
                )
            );

            if (accessControlModel.GetChildrenCount() > 0)
            {
                accessControlModel.LoadChildren(false);

                result.AddRange(accessControlModel.Children.Select(
                    cm => new LazyTreeNode(
                        cm.ToString(),
                        () => GetControlNodes(cm)
                    ) { Tag = cm, ToolTipText = tooltipService.GetTooltip(cm) }
                ));
            }
            return result;
        }

        private IEnumerable<TreeNode> CreateDynamicPropertyNodes(MSAccess.Properties properties)
        {
            return new AccessDynamicPropertiesModel(properties)
                .OrderBy(p => p.Key)
                .Select(p => new TreeNode($"{p.Key}: {p.Value}"));
        }

        private IEnumerable<TreeNode> CreatePropertyNodes(object accessObject)
        {
            var objectProperties = TypeDescriptor.GetProperties(accessObject);
            return objectProperties
                .Cast<PropertyDescriptor>()
                .OrderBy(p => p.Name)
                .Select(p => new TreeNode($"{p.Name}: {p.GetValue(accessObject)}"));
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
