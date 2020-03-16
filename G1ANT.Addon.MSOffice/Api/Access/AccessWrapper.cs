/**
*    Copyright(C) G1ANT Ltd, All rights reserved
*    Solution G1ANT.Addon, Project G1ANT.Addon.MSOffice
*    www.g1ant.com
*
*    Licensed under the G1ANT license.
*    See License.txt file in the project root for full license information.
*
*/

using G1ANT.Addon.MSOffice.Access;
using G1ANT.Addon.MSOffice.Api.Access;
using G1ANT.Addon.MSOffice.Models.Access;
using G1ANT.Language;
using Microsoft.Office.Interop.Access;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace G1ANT.Addon.MSOffice
{
    public partial class AccessWrapper
    {
        private string path;
        private Application application = null;
        private readonly IAccessFormControlsTreeWalker accessFormControlsTreeWalker;
        private readonly IRunningObjectTableService runningObjectTableService;

        public int Id { get; }

        internal AccessWrapper(IAccessFormControlsTreeWalker accessFormControlsTreeWalker, IRunningObjectTableService runningObjectTableService)
        {
            Id = AccessManager.GetFreeId();
            this.accessFormControlsTreeWalker = accessFormControlsTreeWalker;
            this.runningObjectTableService = runningObjectTableService;
        }


        internal void ExecuteOnClick(string path)
        {
            var controlPath = new ControlPathModel(path);
            var control = accessFormControlsTreeWalker.GetAccessControlByPath(application, controlPath);

            var onClickCode = control.TryGetPropertyValue<string>("OnClick");
            if (onClickCode == "[Event Procedure]")
            {
                var formName = controlPath.FormName;

                var form = application.Forms[formName];

                //form.SetFocus();
                //control.Control.Form.SetFocus();
                control.Control.SetFocus();

                RobotWin32.SetForegroundWindow((IntPtr)form.Hwnd);
                System.Windows.Forms.SendKeys.SendWait("{ENTER}");
            }
            else
            {
                control.Control.Application.DoCmd.RunMacro(onClickCode);
            }
        }



        public void Open(string path, string password = "", bool openExclusive = false, bool shouldShowApplication = true)
        {
            if (string.IsNullOrEmpty(path))
                throw new ArgumentNullException(nameof(path));
            this.path = path;

            application = application ?? new Application();

            if (shouldShowApplication)
                Show();
            else
                Hide();

            application.OpenCurrentDatabase(path, openExclusive);
        }

        public void Show()
        {
            application.Visible = true;
        }

        public AccessControlModel GetAccessControlByPath(string path)
        {
            return accessFormControlsTreeWalker.GetAccessControlByPath(application, path);
        }

        public AccessControlModel GetActiveControl(bool getProperties = true, bool getChildrenRecursively = true)
        {
            return new AccessControlModel(application.Screen.ActiveControl, getProperties, getChildrenRecursively);
        }

        public AccessFormModel GetActiveForm()
        {
            return new AccessFormModel(application.Screen.ActiveForm, true, true, false);
        }

        public void Hide()
        {
            application.Visible = false;
        }

        public ICollection<AccessObjectModel> GetAllProjectForms()
        {
            var result = application.CurrentProject.AllForms
                .Cast<AccessObject>()
                .Select(f => new AccessObjectModel(f));

            return result.ToList();
        }

        public ICollection<AccessFormModel> GetAllForms()
        {
            var result = application.Forms
                .Cast<Form>()
                .Select(f => new AccessFormModel(f, true, false, false));

            return result.ToList();
        }

        public AccessFormModel GetForm(string formName)
        {
            var form = application.Forms[formName];
            var result = new AccessFormModel(form, true, true, true);

            return result;
        }


        public void CloseDatabase()
        {
            application.DoCmd.CloseDatabase();
        }

        public void Close()
        {
            application.DoCmd.Close();
        }

        public void Quit(bool saveChanges)
        {
            application.DoCmd.Quit(saveChanges ? AcQuitOption.acQuitSaveAll : AcQuitOption.acQuitSaveNone);
            AccessManager.Remove(this);
        }

        public void KillOrphanedAccessProcesses()
        {
            var processIds = runningObjectTableService.GetOrphanedApplicationProcessIds("msaccess");
            foreach (var processId in processIds)
            {
                var process = Process.GetProcessById(processId);
                process.Kill();
            }
        }

            
        public void Save(string objectType, string objectName)
        {
            var acObjectType = (AcObjectType)Enum.Parse(typeof(AcObjectType), objectType);
            application.DoCmd.Save(acObjectType, objectName);
        }

        public void RunMacro(string macroName)
        {
            application.DoCmd.RunMacro(macroName, 1, true);
            //var me = application.MacroError;
        }

        //public void RunCommand(string command)
        //{
        //    application.RunCommand(AcCommand. command)
        //}

        public dynamic RunProcedure(string procedure)
        {
            var result = application.Run(procedure);

            return result;
        }


        public IReadOnlyCollection<AccessMacroModel> GetMacros()
        {
            var macros = application.CurrentProject.AllMacros;
            var result = new List<AccessMacroModel>();

            foreach (var macro in macros)
            {
                var model = new AccessMacroModel(macro);
                result.Add(model);
            }

            return result;
        }


        //public static int ConvertTwipsToPixels(int twips, MeasureDirection direction)
        //{
        //    return twips * GetPPI(direction) / 1440;
        //}

        //public enum MeasureDirection
        //{
        //    Horizontal = 88,
        //    Vertical = 90
        //}

        //public static int GetPPI(MeasureDirection direction)
        //{
        //    IntPtr dc = GetDC(IntPtr.Zero);

        //    var ppi = GetDeviceCaps(dc, (int)direction); //DEVICECAP LOGPIXELSY

        //    ReleaseDC(IntPtr.Zero, dc);
        //    return ppi;
        //}

        //[DllImport("user32.dll")]
        //static extern IntPtr GetDC(IntPtr hWnd);

        //[DllImport("user32.dll")]
        //static extern bool ReleaseDC(IntPtr hWnd, IntPtr hDC);

        //[DllImport("gdi32.dll")]
        //static extern int GetDeviceCaps(IntPtr hdc, int devCap);

        //public void Click(AccessControlModel control)
        //{
        //    var xTwips = 0;
        //    var yTwips = 0;

        //    var topMostControl = control;
        //    var c = control;
        //    while (c != null)
        //    {
        //        xTwips += c.TryGetPropertyValue<int>("Left");
        //        yTwips += c.TryGetPropertyValue<int>("Top");

        //        xTwips += c.TryGetPropertyValue<int>("LeftMargin");
        //        yTwips += c.TryGetPropertyValue<int>("TopMargin");
        //        xTwips += c.TryGetPropertyValue<int>("LeftPadding");
        //        yTwips += c.TryGetPropertyValue<int>("TopPadding");


        //        topMostControl = c;
        //        c = c.GetParent();

        //        if (c != null)
        //        {
        //            xTwips += c.TryGetPropertyValue<int>("RightMargin");
        //            yTwips += c.TryGetPropertyValue<int>("BottomMargin");
        //            xTwips += c.TryGetPropertyValue<int>("RightPadding");
        //            yTwips += c.TryGetPropertyValue<int>("BottomPadding");
        //        }
        //    }

        //    var form = new AccessFormModel(control.Control.Application.Screen.ActiveForm, false, false, false);
        //    //var form = topMostControl.GetForm();
        //    xTwips += form.X;
        //    yTwips += form.Y;


        //    var x = ConvertTwipsToPixels(xTwips, MeasureDirection.Horizontal);
        //    var y = ConvertTwipsToPixels(yTwips, MeasureDirection.Vertical);

        //    var args = MouseStr.ToMouseEventsArgs(x, y, 0, 0, MouseStr.Action.Left.ToString());
        //    args.ForEach(arg => MouseWin32.MouseEvent(arg.dwFlags, arg.dx, arg.dy, arg.dwData));
        //}


        public void Test()
        {
            var afs = application.CurrentProject.Resources;
            var reports = application.CurrentProject.AllReports.Cast<AccessObject>().ToList();

        }




        //public object RunMacro(string macroName, string args = null)
        //{
        //    List<object> arguments = new List<object> { macroName };
        //    object result = null;
        //    if (!string.IsNullOrEmpty(args))
        //    {
        //        arguments.AddRange(args.Split(','));
        //    }
        //    result = application.GetType().InvokeMember("Run", BindingFlags.InvokeMethod, null, this.application, arguments.ToArray());
        //    return result;
        //}
        //public void InsertText(string text, bool replaceAllText)
        //{
        //    if (!replaceAllText)
        //    {
        //        document.Content.InsertAfter(text);
        //    }
        //    else
        //    {
        //        document.Content.Select();
        //        document.Content.Text = text;
        //    }
        //}
        //public string GetText()
        //{
        //    return document.Content.Text;
        //}
        //public void InsertParagraph()
        //{
        //    document.Content.InsertParagraph();
        //}

        //public void ReplaceWord(string from, string to, bool Match, bool WholeWord)
        //{
        //    document.Content.Find.Execute(from, Match, WholeWord, false, false, false, true, false, 1, to, 2, false, false, false, false);

        //}

        //public void Save(string path)
        //{
        //    if (string.IsNullOrEmpty(path))
        //    {
        //        document.SaveAs();
        //    }
        //    else
        //    {
        //        if (string.IsNullOrEmpty(Path.GetDirectoryName(path)))
        //            this.path = application.Options.DefaultFilePath[Word.WdDefaultFilePath.wdDocumentsPath] + "\\" + path;
        //        else
        //            this.path = path;
        //        document.SaveAs(this.path);
        //    }
        //}

        //public void Export(string path, string type)
        //{
        //    if (string.IsNullOrEmpty(type))
        //    {
        //        type = path.Split('.').LastOrDefault();
        //    }
        //    try
        //    {
        //        string outPath = string.IsNullOrEmpty(path) ? this.path : path;
        //        Word.WdExportFormat format;

        //        switch (type.ToLower())
        //        {
        //            case "pdf":
        //                format = Word.WdExportFormat.wdExportFormatPDF;
        //                break;
        //            case "xps":
        //                format = Word.WdExportFormat.wdExportFormatXPS;
        //                break;
        //            default:
        //                throw new ApplicationException("Unknown format type");
        //        }

        //        document.ExportAsFixedFormat(path, format);
        //    }
        //    catch (Exception)
        //    {

        //        throw;
        //    }
        //}

        //private void Application_WindowDeactivate(Word.Document Doc, Word.Window Wn)
        //{
        //    Close();
        //}

        //public void Close()
        //{
        //    try
        //    {
        //        //application.WindowDeactivate -= Application_WindowDeactivate;
        //        WordManager.Remove(this);

        //        application.Quit(
        //            Word.WdSaveOptions.wdDoNotSaveChanges,
        //            Word.WdOriginalFormat.wdOriginalDocumentFormat,
        //            false);
        //    }
        //    catch (Exception)
        //    {

        //        throw;
        //    }
        //}
    }
}
