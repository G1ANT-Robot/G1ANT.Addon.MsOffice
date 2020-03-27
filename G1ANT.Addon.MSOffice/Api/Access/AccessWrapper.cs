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
using G1ANT.Addon.MSOffice.Models.Access.AccessObjects;
using G1ANT.Addon.MSOffice.Models.Access.Dao;
using G1ANT.Addon.MSOffice.Models.Access.Dao.QueryDefs;
using G1ANT.Addon.MSOffice.Models.Access.Data;
using G1ANT.Addon.MSOffice.Models.Access.Printers;
using G1ANT.Language;
using Microsoft.Office.Interop.Access;
using Microsoft.Office.Interop.Access.Dao;
using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;

namespace G1ANT.Addon.MSOffice
{
    internal partial class AccessWrapper
    {
        private string path;
        private Application application = null;
        private readonly IAccessFormControlsTreeWalker accessFormControlsTreeWalker;

        private readonly IRunningObjectTableService runningObjectTableService;

        internal int Id { get; }

        internal void CreateAccessProject(string path, string connectionString = null)
        {
            application.CreateAccessProject(path, connectionString);
        }

        /// <summary></summary>
        /// <param name="typeOfObjectToExport">Possible values: Form, Function, Query, Report ServerView, StoredProcedure, Table</param>
        /// <param name="objectName"></param>
        /// <param name="destinationPath"></param>
        /// <param name="destinationPathForSchema">The path for the exported schema information. If this argument is omitted, schema information is not exported to a separate XML file.</param>
        /// <param name="destinationPathForPresentation">The path for the exported presentation information. If this argument is omitted, presentation information is not exported.</param>
        /// <param name="destinationDirectoryForImages">The directory for exported images. If this argument is omitted, images are not exported.</param>
        /// <param name="useUTF16ForEncoding">true for UTF16, false for UTF8</param>
        /// <param name="otherFlags">Values: EmbedSchema, ExcludePrimaryKeyAndIndexes, RunFromServer, LiveReportSource, PersistReportML, ExportAllTableAndFieldProperties</param>
        /// <param name="whereCondition">Specifies a subset of records to be exported.</param>
        /// <param name="additionalData">Specifies additional tables to export. This argument is ignored if the OtherFlags argument is set to acLiveReportSource.</param>
        internal void ExportXML(
            string typeOfObjectToExport,
            string objectName,
            string destinationPath,
            string destinationPathForSchema = null,
            string destinationPathForPresentation = null,
            string destinationDirectoryForImages = null,
            bool useUTF16ForEncoding = false,
            string otherFlags = null,
            string whereCondition = null,
            object additionalData = null
        )
        {
            application.ExportXML(
                ToEnum<AcExportXMLObjectType>(typeOfObjectToExport, "acExport"),
                objectName,
                destinationPath,
                destinationPathForSchema,
                destinationPathForPresentation,
                destinationDirectoryForImages,
                useUTF16ForEncoding ? AcExportXMLEncoding.acUTF16 : AcExportXMLEncoding.acUTF8,
                otherFlags != null ? ToEnum<AcExportXMLOtherFlags>(otherFlags) : (AcExportXMLOtherFlags)0,
                whereCondition,
                additionalData
            );
        }

        /// <summary></summary>
        /// <param name="path"></param>
        /// <param name="importOptions">StructureOnly, StructureAndData, AppendData</param>
        internal void ImportXml(string path, string importOptions)
        {
            application.ImportXML(path, ToEnum<AcImportXMLOption>(importOptions));
        }

        internal List<string> GetQueryNames()
        {
            return new AccessQueryDefCollectionModel(application.CurrentDb())
                .Select(q => q.Name)
                .ToList();
        }

        internal List<string> GetFunctionNames()
        {
            return new AccessObjectFunctionCollectionModel(application.CurrentData.AllFunctions)
                .Select(t => t.Name)
                .ToList();
        }

        internal List<string> GetDatabaseDiagramNames()
        {
            return new AccessObjectFunctionCollectionModel(application.CurrentData.AllDatabaseDiagrams)
                .Select(t => t.Name)
                .ToList();
        }

        internal List<string> GetStoredProcedureNames()
        {
            return new AccessObjectFunctionCollectionModel(application.CurrentData.AllStoredProcedures)
                .Select(t => t.Name)
                .ToList();
        }

        internal List<string> GetViewNames()
        {
            return new AccessObjectFunctionCollectionModel(application.CurrentData.AllViews)
                .Select(t => t.Name)
                .ToList();
        }

        internal List<string> GetTableNames()
        {
            return new AccessTableDefCollectionModel(application.CurrentDb().TableDefs)
                .Select(t => t.Name)
                .ToList();
        }

        internal List<string> GetReportNames()
        {
            return new AccessReportCollectionModel(application.Reports)
                .Select(r => r.Name)
                .ToList();
        }

        internal AccessReportModel GetReportDetails(string name) => new AccessReportModel(application, name);
        internal AccessObjectModel GetDatabaseDiagramDetails(string name) => new AccessObjectModel(application.CurrentData.AllStoredProcedures[name]);
        internal AccessObjectModel GetStoredProcedureDetails(string name) => new AccessObjectModel(application.CurrentData.AllDatabaseDiagrams[name]);
        internal AccessObjectModel GetViewDetails(string name) => new AccessObjectModel(application.CurrentData.AllViews[name]);
        internal AccessObjectModel GetFunctionDetails(string name) => new AccessObjectModel(application.CurrentData.AllFunctions[name]);
        internal AccessTableDefModel GetTableDetails(string name) => new AccessTableDefModel(application.CurrentDb().TableDefs[name]);
        internal AccessQueryDefDetailsModel GetQueryDetails(string name) => new AccessQueryDefDetailsModel(application.CurrentDb().QueryDefs[name]);


        internal AccessWrapper(
            IAccessFormControlsTreeWalker accessFormControlsTreeWalker,
            IRunningObjectTableService runningObjectTableService
        )
        {
            Id = AccessManager.GetFreeId();
            this.accessFormControlsTreeWalker = accessFormControlsTreeWalker;
            this.runningObjectTableService = runningObjectTableService;
        }

        internal AccessTableDefModel GetTableDetails(object sourceObjectName)
        {
            var currentDb = application.CurrentDb();
            return new AccessTableDefModel(currentDb.TableDefs[sourceObjectName]);
        }


        internal List<List<object>> GetTableContents(string sourceObjectName)
        {
            return ExecuteSql($"select * from {sourceObjectName}");
        }

        internal List<List<object>> ExecuteSql(string sql, string connectionString = null)
        {
            connectionString = connectionString ?? application.ADOConnectString;

            using (var connection = new OleDbConnection(connectionString))
            using (var command = new OleDbCommand(sql, connection))
            {
                connection.Open();

                var reader = command.ExecuteReader();
                var columnNames = Enumerable.Range(0, reader.FieldCount).Select(i => (object)reader.GetName(i)).ToList();

                var result = new List<List<object>>
                {
                    columnNames
                };

                while (reader.Read())
                {
                    var row = new List<object>(reader.FieldCount);
                    for (var i = 0; i < reader.FieldCount; ++i)
                        row.Add(reader.GetValue(i));

                    result.Add(row);
                }

                connection.Close();

                return result;
            }
        }

        internal void RunSql(string sql, bool useTransaction = false) => application.DoCmd.RunSQL(sql, useTransaction);

        internal string ApplyPrefix(string prefix, string name)
        {
            return name.StartsWith(prefix, StringComparison.CurrentCultureIgnoreCase) ? name : (prefix + name);
        }

        private T ToEnum<T>(string fieldName, string prefixToApply = "ac") => (T)Enum.Parse(typeof(T), ApplyPrefix(prefixToApply, fieldName), true);

        private AcView ToAcView(string viewType) => ToEnum<AcView>(viewType, "acView");
        private AcFormView ToAcFormView(string viewFormType) => ToEnum<AcFormView>(viewFormType);
        private AcWindowMode ToAcWindowMode(string windowMode) => ToEnum<AcWindowMode>(windowMode);
        private AcPrintRange ToAcPrintRange(string printRange) => ToEnum<AcPrintRange>(printRange);
        private AcPrintQuality ToAcPrintQuality(string printRange) => ToEnum<AcPrintQuality>(printRange);

        private AcOpenDataMode ToAcOpenDataMode(bool createNew, bool openReadonly)
        {
            return createNew ? AcOpenDataMode.acAdd : openReadonly ? AcOpenDataMode.acReadOnly : AcOpenDataMode.acEdit;
        }

        private AcFormOpenDataMode ToAcFormOpenDataMode(bool createNew, bool openReadonly, bool openPropertySettings)
        {
            if (createNew)
                return AcFormOpenDataMode.acFormAdd;
            if (openReadonly)
                return AcFormOpenDataMode.acFormReadOnly;
            if (openPropertySettings)
                return AcFormOpenDataMode.acFormPropertySettings;

            return AcFormOpenDataMode.acFormEdit;
        }


        internal void OpenForm(
            string name,
            string viewFormType = "Normal",
            bool createNew = false,
            bool openReadonly = true,
            bool openPropertySettings = false,
            string windowMode = "acWindowNormal",
            string filterName = null,
            string whereCondition = null,
            string openArgs = null
        )
        {
            application.DoCmd.OpenForm(
                name,
                ToAcFormView(viewFormType),
                filterName,
                whereCondition,
                ToAcFormOpenDataMode(createNew, openReadonly, openPropertySettings),
                ToAcWindowMode(windowMode),
                openArgs
            );
        }


        internal void OpenTable(string name, string viewType = "Normal", bool createNew = false, bool openReadonly = true)
        {
            application.DoCmd.OpenTable(name, ToAcView(viewType), ToAcOpenDataMode(createNew, openReadonly));
        }

        internal void OpenStoredProcedure(string name, string viewType = "Normal", bool createNew = false, bool openReadonly = true)
        {
            application.DoCmd.OpenStoredProcedure(name, ToAcView(viewType), ToAcOpenDataMode(createNew, openReadonly));
        }

        internal void OpenReport(string name, string viewType = "Normal", bool createNew = false, bool openReadonly = true)
        {
            application.DoCmd.OpenReport(name, ToAcView(viewType), ToAcOpenDataMode(createNew, openReadonly));
        }

        internal void OpenQuery(string name, string viewType = "Normal", bool createNew = false, bool openReadonly = true)
        {
            application.DoCmd.OpenQuery(name, ToAcView(viewType), ToAcOpenDataMode(createNew, openReadonly));
        }

        internal void OpenFunction(string name, string viewType = "Normal", bool createNew = false, bool openReadonly = true)
        {
            application.DoCmd.OpenFunction(name, ToAcView(viewType), ToAcOpenDataMode(createNew, openReadonly));
        }

        internal void OpenView(string name, string viewType = "Normal", bool createNew = false, bool openReadonly = true)
        {
            application.DoCmd.OpenView(name, ToAcView(viewType), ToAcOpenDataMode(createNew, openReadonly));
        }

        internal void OpenDiagram(string name) => application.DoCmd.OpenDiagram(name);


        internal void Close(AcObjectType objectType, string objectName, bool? save)
        {
            application.DoCmd.Close(
                objectType,
                objectName,
                save.HasValue
                    ? save.Value ? AcCloseSave.acSaveYes : AcCloseSave.acSaveNo
                    : AcCloseSave.acSavePrompt
            );
        }


        /// <summary>
        /// Executes handler assigned to OnClick property; if there's no handler action is repeated for OnEnter property.
        /// </summary>
        /// <param name="path"></param>
        internal void ExecuteDefaultClickEvent(string path) => ExecuteEvents(path, "OnClick", "OnEnter");

        internal void ExecuteEvents(string path, params string[] eventNames)
        {
            var controlPath = new ControlPathModel(path);
            var control = accessFormControlsTreeWalker.GetAccessControlByPath(application, controlPath);

            var formName = controlPath.FormName;
            var form = new AccessFormModel(application.Forms[formName], false, false, false);

            foreach (var eventName in eventNames)
            {
                if (ExecuteHandlerCode(form, control, eventName))
                    return;
            }

            throw new Exception("No action to execute found");
        }

        private bool ExecuteHandlerCode(AccessFormModel form, AccessControlModel control, string actionName)
        {
            var handlerCode = control.TryGetPropertyValue<string>(actionName);
            if (!string.IsNullOrEmpty(handlerCode))
            {
                ExecuteCode(form, control, handlerCode);
                return true;
            }

            return false;
        }

        private void ExecuteCode(AccessFormModel formModel, AccessControlModel control, string code)
        {
            if (code == "[Event Procedure]")
            {
                //form.SetFocus();
                control.SetFocus();

                RobotWin32.SetForegroundWindow((IntPtr)formModel.Hwnd);
                System.Windows.Forms.SendKeys.SendWait("{ENTER}");
            }
            else if (!string.IsNullOrEmpty(code))
            {
                application.DoCmd.RunMacro(code);
            }
            else
                throw new Exception("Code to execute is empty");
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="processId">0 to join to newest instance</param>
        internal void JoinToExistingInstance(int processId = 0)
        {
            application = processId > 0
                ? runningObjectTableService.GetApplicationInstance(processId)
                : runningObjectTableService.GetNewestApplicationInstance();
        }

        internal void New()
        {
            if (application != null)
                throw new ApplicationException("Please close current instance or create new instance");
            application = new Application();
        }

        internal void Open(string path, string password = "", bool openExclusive = false, bool shouldShowApplication = true)
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

        internal void OpenProject(string path, bool openExclusive = false, bool shouldShowApplication = true)
        {
            if (string.IsNullOrEmpty(path))
                throw new ArgumentNullException(nameof(path));
            this.path = path;

            application = application ?? new Application();

            if (shouldShowApplication)
                Show();
            else
                Hide();

            application.OpenAccessProject(path, openExclusive);
        }

        internal void Show() => application.Visible = true;
        internal void Hide() => application.Visible = false;


        internal AccessControlModel GetControlByPath(string path)
        {
            return accessFormControlsTreeWalker.GetAccessControlByPath(application, path);
        }

        internal AccessControlModel GetActiveControl(bool getProperties = true, bool getChildrenRecursively = true)
        {
            return new AccessControlModel(application.Screen.ActiveControl, getProperties, getChildrenRecursively);
        }

        internal AccessFormModel GetActiveForm()
        {
            return new AccessFormModel(application.Screen.ActiveForm, true, true, false);
        }


        internal ICollection<AccessObjectModel> GetAllProjectForms()
        {
            var result = application.CurrentProject.AllForms
                .Cast<AccessObject>()
                .Select(f => new AccessObjectModel(f));

            return result.ToList();
        }

        internal ICollection<AccessFormModel> GetAllForms()
        {
            var result = application.Forms
                .Cast<Form>()
                .Select(f => new AccessFormModel(f, true, false, false));

            return result.ToList();
        }

        /// <summary>Set selected record i.e. in subform</summary>
        /// <param name="objectType">ActiveDataObject, DataTable, DataQuery, DataForm, DataReport, DataServerView, DataStoredProcedure, DataFunction</param>
        /// <param name="objectName"></param>
        /// <param name="operation">Previous, Next, First, Last, GoTo, NewRec</param>
        /// <param name="offset"></param>
        internal void GoToRecord(string objectType = "ActiveDataObject", string objectName = null, string operation = "Next", int offset = -1)
        {
            var acDataObjectType = objectType == "ActiveDataObject" ? AcDataObjectType.acActiveDataObject : ToEnum<AcDataObjectType>(objectType);
            if (acDataObjectType != AcDataObjectType.acActiveDataObject && string.IsNullOrEmpty(objectName))
                throw new Exception($"{nameof(objectType)} differs from ActiveDataObject and {nameof(objectName)} is empty");

            application.DoCmd.GoToRecord(acDataObjectType, objectName, ToEnum<AcRecord>(operation), offset);
        }

        internal AccessFormModel GetForm(string formNameOrPathToSubform, bool getFormProperties = true, bool getControls = true, bool getControlsProperties = false)
        {
            if (formNameOrPathToSubform.Contains("/"))
            {
                var control = GetControlByPath(formNameOrPathToSubform);
                if (!control.HasForm())
                {
                    throw new ApplicationException("Selected control does not contain a subform");
                }
                return new AccessFormModel(control.Control.Form, getFormProperties, getControls, getControlsProperties);
            }
            var form = application.Forms[formNameOrPathToSubform];
            var result = new AccessFormModel(form, getFormProperties, getControls, getControlsProperties);

            return result;
        }


        internal void CloseDatabase() => application.DoCmd.CloseDatabase();

        internal void Quit(bool saveChanges)
        {
            application.DoCmd.Quit(saveChanges ? AcQuitOption.acQuitSaveAll : AcQuitOption.acQuitSaveNone);
            AccessManager.Remove(this);
        }

        internal void CloseCurrentDatabase() => application.CloseCurrentDatabase();


        internal void Save(string objectType, string objectName)
        {
            var acObjectType = (AcObjectType)Enum.Parse(typeof(AcObjectType), objectType);
            application.DoCmd.Save(acObjectType, objectName);
        }

        internal void RunMacro(string macroName)
        {
            application.DoCmd.RunMacro(macroName, 1, true);
            //var me = application.MacroError;
        }

        //internal void RunCommand(string command)
        //{
        //    application.RunCommand(AcCommand. command)
        //}

        internal dynamic RunProcedure(string procedure)
        {
            var result = application.Run(procedure);

            return result;
        }


        internal IReadOnlyCollection<AccessMacroModel> GetMacros()
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


        //internal static int ConvertTwipsToPixels(int twips, MeasureDirection direction)
        //{
        //    return twips * GetPPI(direction) / 1440;
        //}

        //internal enum MeasureDirection
        //{
        //    Horizontal = 88,
        //    Vertical = 90
        //}

        //internal static int GetPPI(MeasureDirection direction)
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

        //internal void Click(AccessControlModel control)
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


        //private void Application_WindowDeactivate(Word.Document Doc, Word.Window Wn)
        //{
        //    Close();
        //}

        //internal void Close()
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

        public AccessPrinterCollectionModel GetPrinters() => new AccessPrinterCollectionModel(application.Printers);
        public AccessPrinterModel GetCurrentPrinter() => new AccessPrinterModel(application.Printer);
        public void SetCurrentPrinter(string name) => application.Printer = application.Printers[name];
        public void PrintActiveObject(string printRange = "PrintAll", int pageFrom = 0, int pageTo = 0, string printQuality = "High", int copyCount = 1, bool collateCopies = true)
        {
            application.DoCmd.PrintOut(ToAcPrintRange(printRange), pageFrom, pageTo, ToAcPrintQuality(printQuality), copyCount, collateCopies);
        }

        public void SetNewPassword(string oldPassword, string newPassword) => application.CodeDb().NewPassword(oldPassword, newPassword);

        public void BeginTransaction() => application.CodeDb().BeginTrans();
        public void CommitTransaction(bool forceFlush = false) => application.CodeDb().CommitTrans(forceFlush ? (int)CommitTransOptionsEnum.dbForceOSFlush : 0);

        public string GetAdoConnectionString() => application.ADOConnectString;
        public string GetBaseConnectionString() => application.CurrentProject.BaseConnectionString;


        public AccessApplicationModel GetApplicationDetails() => new AccessApplicationModel(application);
    }
}
