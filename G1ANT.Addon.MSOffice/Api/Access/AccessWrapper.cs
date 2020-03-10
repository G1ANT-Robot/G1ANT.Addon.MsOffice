﻿/**
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
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;

namespace G1ANT.Addon.MSOffice
{
    public partial class AccessWrapper
    {
        private string path;
        private Application application = null;
        private readonly IAccessFormControlsTreeWalker accessFormControlsTreeWalker;

        internal AccessWrapper(IAccessFormControlsTreeWalker accessFormControlsTreeWalker)
        {
            Id = AccessManager.GetFreeId();
            this.accessFormControlsTreeWalker = accessFormControlsTreeWalker;
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

        public void RunMacro(string macroName)
        {
            application.DoCmd.RunMacro(macroName, 1, true);
            //var me = application.MacroError;
        }


        //public void RunCommand(string command)
        //{
        //    application.RunCommand(AcCommand. command)
        //}

        public dynamic Run(string procedure)
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


        ///// <summary>
        ///// Converts an integer value in twips to the corresponding integer value in pixels on the x-axis.
        ///// </summary>
        ///// <param name="source">The Graphics context to use</param>
        ///// <param name="inTwips">The number of twips to be converted</param>
        ///// <returns>The number of pixels in that many twips</returns>
        //public static int ConvertTwipsToXPixels(Graphics source, int twips)
        //{
        //    return (int)(twips * source.DpiX / 1440.0);
        //}

        ///// <summary>
        ///// Converts an integer value in twips to the corresponding integer value in pixels on the y-axis.
        ///// </summary>
        ///// <param name="source">The Graphics context to use</param>
        ///// <param name="inTwips">The number of twips to be converted</param>
        ///// <returns>The number of pixels in that many twips</returns>
        //public static int ConvertTwipsToYPixels(Graphics source, int twips)
        //{
        //    return (int)(twips * source.DpiY / 1440.0);
        //}

        public static int ConvertTwipsToPixels(int twips, MeasureDirection direction)
        {
            return (int)(twips * GetPPI(direction) / 1440.0);
        }

        public enum MeasureDirection
        {
            Horizontal,
            Vertical
        }

        public static int GetPPI(MeasureDirection direction)
        {
            int ppi;
            IntPtr dc = GetDC(IntPtr.Zero);

            if (direction == MeasureDirection.Horizontal)
                ppi = GetDeviceCaps(dc, 88); //DEVICECAP LOGPIXELSX
            else
                ppi = GetDeviceCaps(dc, 90); //DEVICECAP LOGPIXELSY

            ReleaseDC(IntPtr.Zero, dc);
            return ppi;
        }

        [DllImport("user32.dll")]
        static extern IntPtr GetDC(IntPtr hWnd);

        [DllImport("user32.dll")]
        static extern bool ReleaseDC(IntPtr hWnd, IntPtr hDC);

        [DllImport("gdi32.dll")]
        static extern int GetDeviceCaps(IntPtr hdc, int devCap);

        public void Click(AccessControlModel control)
        {
            var xTwips = control.Control.TryGetPropertyValue<int>("Left");
            var yTwips = control.Control.TryGetPropertyValue<int>("Top");

            var x = ConvertTwipsToPixels(xTwips, MeasureDirection.Horizontal) + 60;
            var y = ConvertTwipsToPixels(xTwips, MeasureDirection.Vertical) + 2;

            var args = MouseStr.ToMouseEventsArgs(x, y, 0,0, MouseStr.Action.Left.ToString());
            args.ForEach(arg => MouseWin32.MouseEvent(arg.dwFlags, arg.dx, arg.dy, arg.dwData));
        }


        public void Test()
        {
            //var control = accessFormControlsTreeWalker.GetAccessControlByPath(application, "/Start/TabCtl52/Caption=Configuration login/");
            var control = accessFormControlsTreeWalker.GetAccessControlByPath(application, "/Start/TabCtl55/Production/Command147/");

            

            //var c = control.Control;

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
