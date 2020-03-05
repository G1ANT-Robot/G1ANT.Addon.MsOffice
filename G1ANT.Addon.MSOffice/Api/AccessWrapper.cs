/**
*    Copyright(C) G1ANT Ltd, All rights reserved
*    Solution G1ANT.Addon, Project G1ANT.Addon.MSOffice
*    www.g1ant.com
*
*    Licensed under the G1ANT license.
*    See License.txt file in the project root for full license information.
*
*/
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using Access = Microsoft.Office.Interop.Access;

namespace G1ANT.Addon.MSOffice
{
    public class AccessWrapper
    {
        private string path;
        private Access.Application application = null;
        //private readonly Access.Document document = null;

        internal AccessWrapper()
        {
            Id = AccessManager.GetFreeId();
        }

        public int Id { get; private set; }

        public void Open(string path, string password = "")
        {
            if (string.IsNullOrEmpty(path))
                throw new ArgumentNullException(nameof(path));
            this.path = path;

            application = new Access.Application();

            application.OpenCurrentDatabase(path, true);

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
