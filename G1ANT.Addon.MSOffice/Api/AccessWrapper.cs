/**
*    Copyright(C) G1ANT Ltd, All rights reserved
*    Solution G1ANT.Addon, Project G1ANT.Addon.MSOffice
*    www.g1ant.com
*
*    Licensed under the G1ANT license.
*    See License.txt file in the project root for full license information.
*
*/
using G1ANT.Addon.MSOffice.Models.Access;
using Microsoft.Office.Interop.Access;
using System;
using System.Collections.Generic;
using System.Linq;


namespace G1ANT.Addon.MSOffice
{
    public partial class AccessWrapper
    {
        private string path;
        private Application application = null;

        internal AccessWrapper()
        {
            Id = AccessManager.GetFreeId();
        }

        public int Id { get; private set; }

        public void Open(string path, string password = "", bool openExclusive = false)
        {
            if (string.IsNullOrEmpty(path))
                throw new ArgumentNullException(nameof(path));
            this.path = path;

            application = new Application();
            

            application.OpenCurrentDatabase(path, openExclusive);

            //Word.Options opt = application.Options;
            //string defaultPath = opt.DefaultFilePath[Word.WdDefaultFilePath.wdDocumentsPath];
            //application.DisplayAlerts = Word.WdAlertLevel.wdAlertsAll;
            //application.Visible = true;

            //if (string.IsNullOrEmpty(path))
            //{
            //    document = application.Documents.Add(!string.IsNullOrEmpty(path) ? (object)path : Missing.Value);
            //    document.Activate();
            //}
            //else
            //{
            //    if (string.IsNullOrEmpty(Path.GetDirectoryName(path)))
            //        path = defaultPath + "\\" + path;
            //    document = application.Documents.Open(path);
            //    document.Activate();
            //}
            //this.path = path;
            
        }

        public ICollection<AccessObjectModel> GetAllProjectForms()
        {
            var result = application.CurrentProject.AllForms
                .Cast<AccessObject>()
                .Select(f => new AccessObjectModel(f));

            return result.ToList();
        }

        //public ICollection<AccessFormModel> GetAllOpenForms() => GetAllForms().Where(f => f.IsLoaded).ToList();

        public ICollection<AccessFormModel> GetAllForms()
        {
            var result = application.Forms
                .Cast<Form>()
                .Select(f => new AccessFormModel(f, false));

            return result.ToList();
        }


        public AccessFormModel GetForm(string formName)
        {
            var form = application.Forms[formName];
            var result = new AccessFormModel(form, true);

            return result;
        }


        public void Test()
        {
            var afs = application.CurrentProject.Resources;

            var reports = application.CurrentProject.AllReports.Cast<AccessObject>().ToList();

        //    Access.Forms forms = application.Forms;
        //    var count = forms.Count;
        //    Access.Form f = forms[0];
        //    _Form3 f2 = forms[0];
        }

        //public void Show()
        //{
        //    document.Activate();
        //    document.Application.ShowMe();
        //    Language.RobotWin32.BringWindowToFront((IntPtr)document.Application.ActiveWindow.Hwnd);
        //}

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
