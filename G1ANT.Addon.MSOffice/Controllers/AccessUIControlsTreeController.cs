using G1ANT.Addon.MSOffice.Api.Access;
using G1ANT.Addon.MSOffice.Models.Access;
using G1ANT.Language;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Windows.Forms;

namespace G1ANT.Addon.MSOffice.Controllers
{
    public class AccessUIControlsTreeController
    {
        private IMainForm mainForm;

        public AccessUIControlsTreeController()
        { }



        //struct RunningObject
        //{
        //    public string Name;
        //    public object Application;
        //}

        //[DllImport("ole32.dll")]
        //static extern int CreateBindCtx(uint reserved, out IBindCtx ppbc);

        //// Returns the contents of the Running Object Table (ROT), where
        //// open Microsoft applications and their documents are registered.
        //List<RunningObject> GetRunningObjects()
        //{
        //    var result = new List<RunningObject>();

        //    CreateBindCtx(0, out IBindCtx bindContext);
            
        //    bindContext.GetRunningObjectTable(out IRunningObjectTable runningObjectTable);
        //    runningObjectTable.EnumRunning(out IEnumMoniker monikerEnumerator);

        //    monikerEnumerator.Reset();

        //    var monikers = new IMoniker[1];
        //    IntPtr numFetched = IntPtr.Zero;
        //    while (monikerEnumerator.Next(1, monikers, numFetched) == 0)
        //    {
        //        RunningObject running;
        //        monikers[0].GetDisplayName(bindContext, null, out running.Name);
        //        runningObjectTable.GetObject(monikers[0], out running.Application);
        //        result.Add(running);
        //    }
        //    return result;
        //}



        public void Initialize(IMainForm mainForm) => this.mainForm = mainForm;


        public bool initialized = false;
        public void InitRootElements(ComboBox applications, TreeView controlsTree)
        {
            if (initialized)
                return;
            initialized = true;


            const uint OBJID_NATIVEOM = 0xFFFFFFF0;

            var accessProcesses = new List<Process>(Process.GetProcessesByName("MSACCESS").Concat(Process.GetProcessesByName("MSACCESS.EXE")));

            foreach (var accessProcess in accessProcesses)
            {
                var mainHandle = accessProcess.MainWindowHandle;
                if (mainHandle.ToInt32() > 0)
                {
                    var IID_IDispatch = new Guid("{00020400-0000-0000-C000-000000000046}");

                    int res = OleAccWrapper.AccessibleObjectFromWindow(mainHandle, OBJID_NATIVEOM, ref IID_IDispatch, out Microsoft.Office.Interop.Access.Application app);
                    if (res >= 0)
                    {
                        //Debug.Assert(app.hWndAccessApp() == mainHandle);
                        //Console.WriteLine(app.Name);

                        applications.Items.Add($"{app.Name} {app.CurrentProject.Name}");
                        // todo: check how to release app COM object when done using it
                    }
                    else
                        throw new Exception(); //todo: collect exception and throw AggregateException
                }
            }

            if (applications.Items.Cast<object>().Any())
                applications.SelectedIndex = 0;

            //try
            //{
            //    //var runningObjects = GetRunningObjects();
            //    //foreach (var runningObject in runningObjects)
            //    //{
            //    //    if (runningObject.Application is Microsoft.Office.Interop.Access.Application)
            //    //    {
            //    //        var name = runningObject.Name;
            //    //    }
            //    //}



            //    var access = (Microsoft.Office.Interop.Access.Application)Marshal.GetActiveObject("Access.Application");
            //}
            //catch (Exception ex)
            //{
            //    if (ex.Message.Contains("MK_E_UNAVAILABLE"))
            //        return;
            //    throw;
            //}


            controlsTree.BeginUpdate();
            controlsTree.Nodes.Clear();

            //var jvms = nodeService.GetJvmNodes();
            //foreach (var jvm in jvms)
            //{
            //    var name = $"{jvm.Name} {jvm.JvmId}";
            //    var rootNode = controlsTree.Nodes.Add(name);
            //    rootNode.Tag = jvm;

            //    var windows = nodeService.GetChildNodes(jvm);
            //    rootNode.Nodes.AddRange(windows.Select(w => CreateTreeNode(w)).ToArray());

            //    rootNode.Expand();
            //}

            controlsTree.EndUpdate();
        }

        private TreeNode CreateTreeNode(AccessControlModel controlModel)
        {
            var treeNode = new TreeNode(GetNameForNode(controlModel))
            {
                Tag = controlModel,
                ToolTipText = GetTooltip(controlModel)
            };

            if (controlModel.GetChildrenCount() > 0)
                treeNode.Nodes.Add("");

            return treeNode;
        }

        private string FormatLongLine(string line)
        {
            const int maxLineLength = 100;
            if (line.Length <= maxLineLength)
                return line;

            var sb = new StringBuilder(line.Length);
            var isFirstLine = true;
            do
            {
                var linePart = line.Substring(0, Math.Min(line.Length, maxLineLength));
                line = line.Substring(linePart.Length);
                sb.AppendLine((isFirstLine ? "" : "\t") + linePart);
                isFirstLine = false;
            } while (line != "");

            return sb.ToString();
        }

        private string GetTooltip(AccessControlModel controlModel)
        {
            return "";
            //var nodeProperties = controlModel.GetType().GetProperties()
            //    .Where(p => p.Name != nameof(controlModel.Node))
            //    .Select(p => new { Name = p.Name, Value = p.GetValue(controlModel) })
            //    .Select(v => new { Name = v.Name, Value = v.Value is IEnumerable<string> ? string.Join(", ", v.Value as IEnumerable<string>) : v.Value });

            //return string.Join("\r\n", nodeProperties.Where(np => !string.IsNullOrEmpty(np.Value?.ToString())).Select(np => $"{np.Name}: {FormatLongLine(np.Value.ToString())}"));
        }

        public void CopyNodeDetails(TreeNode treeNode)
        {
            var node = (AccessControlModel)treeNode.Tag;
            var tooltip = GetTooltip(node);
            Clipboard.SetText(tooltip);
        }

        private static string GetNameForNode(AccessControlModel controlModel)
        {
            var name = controlModel.Name;

            //if (controlModel.Id > 0)
            //    name += $" {controlModel.Id}";

            //if (!string.IsNullOrEmpty(controlModel.Name))
            //{
            //    name += ": ";
            //    name += $"\"{controlModel.Name}\"";
            //}

            return name;
        }

        public void LoadChildNodes(TreeNode treeNode)
        {
            if (treeNode.Parent == null)
                return; // don't clear jvms and their windows as they are already rendered

            var node = (AccessControlModel)treeNode.Tag;
            treeNode.Nodes.Clear();

            node.LoadChildren(false);

            var children = node.Children;
            foreach (var child in children)
            {
                treeNode.Nodes.Add(CreateTreeNode(child));
            }
        }

        public void InsertPathIntoScript(TreeNode node)
        {
            //if (node != null)
            //{
            //    var controlModel = (AccessControlModel)node.Tag;
            //    var path = pathService.GetXPathTo(controlModel);

            //    if (mainForm == null)
            //        MessageBox.Show(path);
            //    else
            //        mainForm.InsertTextIntoCurrentEditor($"{SpecialChars.Text}{path}{SpecialChars.Text}");
            //}
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
