using G1ANT.Addon.MSOffice.Api.Access;
using G1ANT.Addon.MSOffice.Forms;
using G1ANT.Addon.MSOffice.Models.Access;
using G1ANT.Addon.MSOffice.Models.Access.AccessObjects;
using G1ANT.Addon.MSOffice.Models.Access.Dao;
using G1ANT.Addon.MSOffice.Models.Access.Dao.Containers;
using G1ANT.Addon.MSOffice.Models.Access.Dao.Documents;
using G1ANT.Addon.MSOffice.Models.Access.Dao.Fields;
using G1ANT.Addon.MSOffice.Models.Access.Dao.Parameters;
using G1ANT.Addon.MSOffice.Models.Access.Dao.Properties;
using G1ANT.Addon.MSOffice.Models.Access.Dao.QueryDefs;
using G1ANT.Addon.MSOffice.Models.Access.Data;
using G1ANT.Addon.MSOffice.Models.Access.Forms.Recordsets;
using G1ANT.Addon.MSOffice.Models.Access.Modules;
using G1ANT.Addon.MSOffice.Models.Access.Printers;
using G1ANT.Addon.MSOffice.Models.Access.Resources;
using G1ANT.Addon.MSOffice.Models.Access.VBE;
using G1ANT.Language;
using Microsoft.Office.Interop.Access.Dao;
using System;
using System.Collections;
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
        private const string FormLabel = "Form";
        private const string MacrosLabel = "Macros";
        private const string ReportsLabel = "Reports";
        private const string ReportLabel = "Report";
        private const string ResourcesLabel = "Resources";
        private const string ModulesLabel = "Modules";

        private const string DatabaseLabel = "Database";
        private const string DatabaseDiagramsLabel = "Database Diagrams";
        private const string FunctionsLabel = "Functions";
        private const string QueryDefsLabel = "Query Definitions";
        private const string StoredProceduresLabel = "Stored Prodecures";
        private const string TablesLabel = "Tables";
        private const string FieldsLabel = "Fields";
        private const string IndexesLabel = "Indexes";
        private const string ViewsLabel = "Views";
        private const string ContainersLabel = "Containers";
        private const string DocumentsLabel = "Documents";
        private const string RecordsetsLabel = "Recordsets";
        private const string RecordsetLabel = "Recordset";

        private const string PrintersLabel = "Printers";

        private const string PropertiesLabel = "Properties";
        private const string ParametersLabel = "Parameters";
        private const string DynamicPropertiesLabel = "Dynamic Properties";

        private const string InternalName = "internal";

        private IMainForm mainForm;
        public bool initialized = false;
        private readonly IRunningObjectTableService runningObjectTableService;
        private readonly TreeView controlsTree;
        private readonly ComboBox applications;
        public List<TreeNode> expandedTreeNodes = new List<TreeNode>();
        private TreeNode selectedTreeNode;

        public AccessUIControlsTreeController(
            TreeView controlsTree, ComboBox applications,
            IRunningObjectTableService runningObjectTableService
        )
        {
            this.runningObjectTableService = runningObjectTableService;
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

        internal void SelectedApplicationChanged(RotApplicationModel model)
        {
            if (model.Application == null)
            {
                controlsTree.Nodes.Clear();
                return;
            }

            selectedTreeNode = controlsTree.SelectedNode;
            CollectExpandedTreeNodeModels(controlsTree.Nodes);

            controlsTree.BeginUpdate();
            controlsTree.Nodes.Clear();

            var application = model.Application;
            var currentDb = new Lazy<Database>(() => application.CurrentDb());

            controlsTree.Nodes.AddRange(
                new TreeNode[]
                {
                    new LazyTreeNode(ApplicationLabel, model).AddRange(() => GetApplicationNodes(model)),
                    new LazyTreeNode(FormsLabel, model).AddRange(() => GetFormNodes(model)),
                    new LazyTreeNode(MacrosLabel, model).AddRange(() => GetAccessObjectNodes(new AccessObjectMacroCollectionModel(model))),
                    new LazyTreeNode(ReportsLabel, model).AddRange(() => GetReportNodes(new AccessReportCollectionModel(application.Reports))),
                    new LazyTreeNode(ResourcesLabel, model).AddRange(() => GetResourceNodes(new AccessResourceCollectionModel(application.CurrentProject.Resources))),
                    new LazyTreeNode(ModulesLabel, model).AddRange(() => GetModuleNodes(new AccessModuleCollectionModel(application.Modules))),
                    new TreeNode(DatabaseLabel) {
                        Nodes = {
                            new LazyTreeNode(DatabaseDiagramsLabel, model).AddRange(() => GetAccessObjectNodes(new AccessObjectDatabaseDiagramCollectionModel(model))),
                            new LazyTreeNode(FunctionsLabel, model).AddRange(() => GetAccessObjectNodes(new AccessObjectFunctionCollectionModel(model))),
                            new LazyTreeNode(QueryDefsLabel, model).AddRange(() => GetQueryNodes(new AccessQueryDefCollectionModel(currentDb.Value))),
                            new LazyTreeNode(StoredProceduresLabel, model).AddRange(() => GetAccessObjectNodes(new AccessObjectStoredProcedureCollectionModel(model))),
                            new LazyTreeNode(TablesLabel, model).AddRange(() => GetTableDefNodes(new AccessTableDefCollectionModel(currentDb.Value.TableDefs))),
                            new LazyTreeNode(ViewsLabel, model).AddRange(() => GetAccessObjectNodes(new AccessObjectViewCollectionModel(model))),
                            new LazyTreeNode(ContainersLabel, model).AddRange(() => GetContainerNodes(new AccessContainerCollectionModel(currentDb.Value.Containers))),
                            new LazyTreeNode(RecordsetsLabel, model).AddRange(() => GetRecordsetNodes(new AccessRecordsetCollectionModel(currentDb.Value.Recordsets)))
                        },
                    },
                    new LazyTreeNode(PrintersLabel, model).AddRange(() => GetPrinterNodes(new AccessPrinterCollectionModel(application.Printers))),
                }
            );

            ApplyExpandedTreeNodes(controlsTree.Nodes);
            controlsTree.EndUpdate();
        }

        private IEnumerable<TreeNode> GetRecordsetNodes(AccessRecordsetCollectionModel model)
        {
            return model.Select(r => GetRecordsetNode(r));
        }

        private TreeNode GetRecordsetNode(AccessRecordsetModel model)
        {
            var result = new LazyTreeNode(model).Add(GetDaoFieldsParentNode(model.Fields));

            if (model.Connection != null)
                result.Add(new LazyTreeNode("Connection", model.Connection).Add(GetConnectionNode(model.Connection)));

            result.AddRange(() => GetObjectPropertiesAsTreeNodes(model));

            return result;
        }

        private LazyTreeNode GetDaoFieldsParentNode(Lazy<AccessDaoFieldCollectionModel> model)
        {
            return new LazyTreeNode(FieldsLabel, model).AddRange(
                () => model.Value.Select(
                    field => new LazyTreeNode(field).AddRange(() => GetObjectPropertiesAsTreeNodes(field))
                )
            );
        }

        private TreeNode GetConnectionNode(AccessConnectionModel model)
        {
            return new LazyTreeNode(model)
                .Add(() => new LazyTreeNode(QueryDefsLabel, model).AddRange(() => GetQueryNodes(model.QueryDefs.Value)))
                .Add(() => new LazyTreeNode(RecordsetsLabel, model).AddRange(() => GetRecordsetNodes(model.Recordsets.Value)))
                .Add(() => new LazyTreeNode(DatabaseLabel, model).Add(GetDatabaseNode(model.Database.Value)))
                .AddRange(() => GetObjectPropertiesAsTreeNodes(model));
        }

        private TreeNode GetDatabaseNode(AccessDatabaseModel model)
        {
            return new LazyTreeNode(model)
                .Add(() => new LazyTreeNode(QueryDefsLabel).AddRange(() => GetQueryNodes(model.QueryDefs.Value)))
                .Add(() => new LazyTreeNode(RecordsetsLabel).AddRange(() => GetRecordsetNodes(model.Recordsets.Value)))
                .Add(() => GetDaoPropertyParentNode(model.Properties))
                .Add(() => new LazyTreeNode("Connection").Add(GetConnectionNode(model.Connection.Value)))
                .Add(() => new LazyTreeNode(ContainersLabel).AddRange(() => GetContainerNodes(model.Containers.Value)))
                .AddRange(() => GetObjectPropertiesAsTreeNodes(model));
        }

        private IEnumerable<TreeNode> GetDaoPropertyNodes(Lazy<AccessDaoPropertyCollectionModel> model) => model.Value.Select(p => new LazyTreeNode(p).EmptyChildren());
        private TreeNode GetDaoPropertyParentNode(Lazy<AccessDaoPropertyCollectionModel> model)
        {
            return new LazyTreeNode(PropertiesLabel).AddRange(() => GetDaoPropertyNodes(model));
        }

        private IEnumerable<TreeNode> GetDynamicPropertyNodes(Lazy<AccessDynamicPropertyCollectionModel> model) => model.Value.Select(p => new LazyTreeNode(p).EmptyChildren());
        private TreeNode GetDynamicPropertyParentNode(Lazy<AccessDynamicPropertyCollectionModel> model, string label = PropertiesLabel)
        {
            return new LazyTreeNode(label).AddRange(() => GetDynamicPropertyNodes(model));
        }


        private IEnumerable<TreeNode> GetReportNodes(AccessReportCollectionModel model)
        {
            return model.Select(r => GetReportNode(r));
        }

        private TreeNode GetReportNode(AccessReportModel model)
        {
            return new LazyTreeNode(model)
                .Add(GetDynamicPropertyParentNode(model.Properties))
                .AddRange(() => GetObjectPropertiesAsTreeNodes(model));
        }

        private IEnumerable<TreeNode> GetContainerNodes(AccessContainerCollectionModel model)
        {
            return model.Select(c => new LazyTreeNode(c).AddRange(() => GetContainerDetailNodes(c)));
        }

        private IEnumerable<TreeNode> GetContainerDetailNodes(AccessContainerModel model)
        {
            return new TreeNode[] {
                GetDaoPropertyParentNode(model.Properties),
                new LazyTreeNode(DocumentsLabel).AddRange(() => GetContainerDocumentsNodes(model.Documents.Value)),
            }.Concat(GetObjectPropertiesAsTreeNodes(model));
        }

        private IEnumerable<TreeNode> GetContainerDocumentsNodes(AccessDocumentCollectionModel documents)
        {
            foreach (var document in documents)
            {
                yield return new LazyTreeNode(document)
                    .Add(GetDaoPropertyParentNode(document.Properties))
                    .AddRange(() => GetObjectPropertiesAsTreeNodes(document));
            }

        }

        private IEnumerable<TreeNode> GetPrinterNodes(AccessPrinterCollectionModel model)
        {
            return model.Select(
                p => new LazyTreeNode(p).AddRange(() => GetObjectPropertiesAsTreeNodes(p))
            );
        }

        internal void ViewDataFromTable(string tableName)
        {
            try
            {
                var app = GetCurrentApplication();
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

            var application = GetCurrentApplication();
            new Thread(() =>
            {
                try
                {
                    OpenForm(model, openInDesigner, application);
                    var newNode = GetLoadedFormNode(application, selectedNode.Text);
                    controlsTree.FindForm().Invoke((MethodInvoker)delegate { ReplaceNode(selectedNode, newNode); });

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

        private MSAccess.Application GetCurrentApplication()
        {
            return ((RotApplicationModel)applications.SelectedItem).Application;
        }

        private void OpenForm(AccessObjectModel formToLoad, bool openInDesigner, MSAccess.Application application)
        {
            try
            {
                var formName = formToLoad.FullName ?? formToLoad.Name;
                application.DoCmd.OpenForm(
                    formName,
                    openInDesigner ? MSAccess.AcFormView.acDesign : MSAccess.AcFormView.acNormal
                );

                var form = application.Forms[formName];
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
                var app = GetCurrentApplication();
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

            try
            {
                if (treeNode is LazyTreeNode lazyTreeNode)
                    lazyTreeNode.LoadLazyChildren();

                ApplyExpandedTreeNodes(treeNode.Nodes);
            }
            catch { }

            controlsTree.EndUpdate();
            Cursor.Current = oldCursor;
        }


        private TreeNode[] GetObjectPropertiesAsTreeNodes(object @object)
        {
            return @object
                .GetType()
                .GetProperties()
                .Select(p => new { p.Name, Value = p.GetValue(@object) })
                .Where(p => !(p.Value is IEnumerable) || p.Value is string)
                .Select(p => new TreeNode($"{p.Name}: {p.Value}"))
                .ToArray();
        }

        private IEnumerable<TreeNode> GetDaoParameterNodes(Lazy<AccessDaoParameterCollectionModel> model)
        {
            return model.Value.Select(
               p => new LazyTreeNode(p).AddRange(() => GetDaoPropertyNodes(p.Properties))
            );
        }

        private IEnumerable<TreeNode> GetQueryDetailsNodes(AccessQueryDefModel model)
        {
            var details = model.Details.Value;

            return new TreeNode[]
            {
                GetDaoFieldsParentNode(details.Fields),
                new LazyTreeNode(ParametersLabel, model).AddRange(() => GetDaoParameterNodes(details.Parameters)),
                GetDaoPropertyParentNode(details.Properties),
            }.Concat(GetObjectPropertiesAsTreeNodes(model));
        }

        private IEnumerable<TreeNode> GetTableDefNodes(AccessTableDefCollectionModel model)
        {
            return model.Select(
                td => new LazyTreeNode(td)
                    .Add(() => GetDaoFieldsParentNode(td.Fields))
                    .Add(() => GetDaoPropertyParentNode(td.Properties))
                    .Add(() => GetIndexParentNode(td))
                    .AddRange(() => GetObjectPropertiesAsTreeNodes(td))
            );
        }

        private LazyTreeNode GetIndexParentNode(AccessTableDefModel td)
        {
            return new LazyTreeNode(IndexesLabel).AddRange(
                () => td.Indexes.Value.Select(i => GetIndexNodes(i))
            );
        }

        private LazyTreeNode GetIndexNodes(AccessTableDefIndexModel index)
        {
            return new LazyTreeNode(index)
                .Add(() => GetDaoFieldsParentNode(index.Fields))
                .Add(() => GetDaoPropertyParentNode(index.Properties))
                .AddRange(() => GetObjectPropertiesAsTreeNodes(index));
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

        private IEnumerable<TreeNode> GetModulePropertyNodes(AccessModuleModel model)
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

        private IEnumerable<TreeNode> GetModuleNodes(AccessModuleCollectionModel model)
        {
            return model.Select(
                mm => new LazyTreeNode(mm).AddRange(() => GetModulePropertyNodes(mm))
            );
        }


        private IEnumerable<TreeNode> GetResourceNodes(AccessResourceCollectionModel model)
        {
            return model.Select(rm => new TreeNode(rm.ToString()) { Tag = rm });
        }

        private string GetDetailedString(object @object)
        {
            return @object is IDetailedNameModel dnm ? dnm.ToDetailedString() : @object?.ToString();
        }

        private IEnumerable<TreeNode> GetAccessObjectNodes(AccessObjectCollectionModel objects)
        {
            return objects.Select(o => new TreeNode(o.ToString()) { Tag = o, ToolTipText = GetDetailedString(o) });
        }


        private IEnumerable<TreeNode> GetQueryNodes(AccessQueryDefCollectionModel model)
        {
            return model.Select(
                query => new LazyTreeNode(query).AddRange(() => GetQueryDetailsNodes(query))
            );
        }


        private bool IsEmptyNode(TreeNode node)
        {
            return node.Nodes.Count == 1 && node.Nodes[0].Text == "";
        }

        private TreeNode GetLoadedFormNode(MSAccess.Application application, string formName)
        {
            var formModel = new AccessFormModel(application.Forms[formName], false, false, false);
            return GetLoadedFormNode(formModel);
        }


        private TreeNode GetLoadedFormNode(AccessFormModel model) => new LazyTreeNode(model).AddRange(() => GetFormDetailNodes(model));
        

        private IEnumerable<TreeNode> GetFormNodes(RotApplicationModel rotApplicationModel)
        {
            var formAccessObjects = new AccessObjectFormCollectionModel(rotApplicationModel);

            foreach (var accessObject in formAccessObjects.OrderByDescending(f => f.IsLoaded).ThenBy(f => f.Name))
            {
                if (accessObject.IsLoaded)
                {
                    yield return GetLoadedFormNode(rotApplicationModel.Application, accessObject.Name);
                }
                else
                {
                    yield return new LazyTreeNode(accessObject).EmptyChildren();
                }
            }
        }


        private List<TreeNode> GetFormDetailNodes(AccessFormModel model)
        {
            var result = new List<TreeNode>();

            if (model.HasRecordset())
            {
                result.Add(new LazyTreeNode(RecordsetLabel).Add(GetRecordsetNode(model.Recordset.Value)));
            }

            var dynamicProperties = new Lazy<AccessDynamicPropertyCollectionModel>(() => new AccessDynamicPropertyCollectionModel(model.Form.Properties));

            result.Add(
                new LazyTreeNode(PropertiesLabel)
                    .Add(GetDynamicPropertyParentNode(dynamicProperties, DynamicPropertiesLabel))
                    .AddRange(() => CreatePropertyNodes(model.Form))
            );

            model.LoadControls(false);

            result.AddRange(
                model.Controls.Select(
                    cn => new LazyTreeNode(cn).AddRange(() => GetControlNodes(cn))
                )
            );

            return result;
        }

        private List<TreeNode> GetControlNodes(AccessControlModel model)
        {
            var result = new List<TreeNode>();

            if (model.HasForm())
            {
                result.Add(new LazyTreeNode(FormLabel).Add(
                    GetLoadedFormNode(model.GetForm())
                ));
            }

            if (model.HasReport())
            {
                result.Add(new LazyTreeNode(ReportLabel).Add(
                    () => GetReportNode(model.Report.Value)
                ));
            }

            var dynamicProperties = new Lazy<AccessDynamicPropertyCollectionModel>(() => new AccessDynamicPropertyCollectionModel(model.Control.Properties));
            result.Add(
                new LazyTreeNode(PropertiesLabel)
                    .Add(GetDynamicPropertyParentNode(dynamicProperties))
                    .AddRange(() => CreatePropertyNodes(model.Control))
            );


            if (model.GetChildrenCount() > 0)
            {
                model.LoadChildren(false);

                result.AddRange(model.Children.Select(
                    cm => new LazyTreeNode(cm).AddRange(() => GetControlNodes(cm))
                ));
            }
            return result;
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
