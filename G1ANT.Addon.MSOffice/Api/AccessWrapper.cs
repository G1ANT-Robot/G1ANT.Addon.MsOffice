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


        public AccessControlModel GetAccessControlByPath(string path)
        {
            var pathElements = path.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);

            var formName = pathElements[0];
            var form = application.Forms[formName];

            Control controlFound = null;

            var controls = form.Controls.OfType<Control>()/*.Select(c => new AccessControlModel(c, true, false))*/.ToList();

            pathElements = pathElements.Skip(1).ToArray();

            for (var i = 0; i < pathElements.Length; i++)
            {
                controlFound = GetMatchingControl(controls, pathElements[i]);

                if (i == pathElements.Length - 1)
                    break; // don't load children of last element of path
                controls = controlFound.Controls.OfType<Control>().ToList();
            }

            return new AccessControlModel(controlFound, true, false);
        }


        private Control GetMatchingControl(ICollection<Control> controls, string pathElement)
        {
            if (string.IsNullOrEmpty(pathElement))
                throw new ArgumentNullException(nameof(pathElement));

            var pathParts = pathElement.Split('=');
            var propertyName = pathParts.Length > 1 ? pathParts[0] : "Name";
            var propertyValue = pathParts.Last();
            var childIndex = -1;

            if (propertyValue.Contains("[") && propertyValue.Contains("]"))
            {
                var childIndexValue = propertyValue.Substring(propertyValue.IndexOf("[") + 1);
                childIndexValue = childIndexValue.Substring(0, childIndexValue.IndexOf("]"));
                childIndex = int.Parse(childIndexValue);
                propertyValue = propertyValue.Substring(0, propertyValue.IndexOf("["));
            }

            IEnumerable<Control> controlsFound = controls;
            if (pathParts.Contains("=") || childIndex < 0)
                controlsFound = controlsFound.Where(c => c.Properties[propertyName]?.Value?.ToString() == propertyValue);

            if (childIndex >= 0)
                controlsFound = controlsFound.Skip(childIndex);

            var controlFound = controlsFound.FirstOrDefault();

            if (controlFound == null)
                throw new Exception($"Element {pathElement} not found (index is {childIndex})");

            return controlFound;
        }


        public void CloseDatabase()
        {
            application.DoCmd.CloseDatabase();
        }

        public void Quit(bool saveChanges)
        {
            application.DoCmd.Quit(saveChanges ? AcQuitOption.acQuitSaveAll : AcQuitOption.acQuitSaveNone);
            AccessManager.Remove(this);
        }

        public void Save(string objectType, string objectName)
        {
            var acObjectType = (AcObjectType)Enum.Parse(typeof(AcObjectType), objectType);
            application.DoCmd.Save(acObjectType, objectName);
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
