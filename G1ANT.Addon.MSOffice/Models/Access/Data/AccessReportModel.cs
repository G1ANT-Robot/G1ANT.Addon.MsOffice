using G1ANT.Addon.MSOffice.Models.Access.Modules;
using Microsoft.Office.Interop.Access;
using System;
using System.Collections.Generic;
using System.Linq;

namespace G1ANT.Addon.MSOffice.Models.Access.Data
{
    internal class AccessReportModel : INameModel
    {
        internal enum CurrentViewEnum
        {
            Design = 0,
            PrintPreview = 5,
            Report = 6,
            Layout = 7
        }

        public Report Report { get; }
        public string Name { get; }

        //public AccessControlModel ActiveControl { get => new AccessControlModel(Report.ActiveControl, false, false); }
        //public AccessControlModel DefaultControl { get => new AccessControlModel(Report.DefaultControl, false, false); }
        public Lazy<List<AccessControlModel>> Controls { get; }
        public Lazy<AccessDynamicPropertyCollectionModel> Properties { get; }
        public Lazy<ModuleModel> Module { get; }
        public bool AllowDesignChanges { get => Report.AllowDesignChanges; set => Report.AllowDesignChanges = value; }
        public bool AllowLayoutView { get => Report.AllowLayoutView; set => Report.AllowLayoutView = value; }
        public bool AllowReportView { get => Report.AllowReportView; set => Report.AllowReportView = value; }
        public bool AutoCenter { get => Report.AutoCenter; set => Report.AutoCenter = value; }

        public bool AutoResize { get => Report.AutoResize; set => Report.AutoResize = value; }
        public byte BorderStyle { get => Report.BorderStyle; set => Report.BorderStyle = value; }
        public string Caption { get => Report.Caption; set => Report.Caption = value; }
        public bool CloseButton { get => Report.CloseButton; set => Report.CloseButton = value; }
        public bool ControlBox { get => Report.ControlBox; set => Report.ControlBox = value; }
        public int CurrentRecord { get { try { return Report.CurrentRecord; } catch { return -1; } } set => Report.CurrentRecord = value; }

        /// <summary>You can use the CurrentView property to determine how a report is currently displayed. Read/write Integer./// </summary>
        public CurrentViewEnum CurrentView { get => (CurrentViewEnum)Report.CurrentView; set => Report.CurrentView = (short)value; }

        public string Filter { get => ((_Report3)Report).Filter; set => ((_Report3)Report).Filter = value; }
        public bool FilterOn { get => Report.FilterOn; set => Report.FilterOn = value; }
        public bool FilterOnLoad { get => Report.FilterOnLoad; set => Report.FilterOnLoad = value; }

        public bool FitToPage { get => Report.FitToPage; set => Report.FitToPage = value; }
        public string FormName { get => Report.FormName; set => Report.FormName = value; }

        public short GridX { get => Report.GridX; set => Report.GridX = value; }
        public short GridY { get => Report.GridY; set => Report.GridY = value; }

        //public int HasData { get => Report.HasData; }

        public int Height
        {
            get { try { return Report.Height; } catch { return -1; } }
            set => Report.Height = value;
        }
        public short Width { get => Report.Width; set => Report.Width = value; }
        public int Top
        {
            get { try { return Report.Top; } catch { return -1; } }
            set => Report.Top = value;
        }
        public int Left
        {
            get { try { return Report.Left; } catch { return -1; } }
            set => Report.Left = value;
        }

        public short WindowHeight { get => Report.WindowHeight; set => Report.WindowHeight = value; }
        public short WindowWidth { get => Report.WindowWidth; set => Report.WindowWidth = value; }
        public short WindowTop { get => Report.WindowTop; }
        public short WindowLeft { get => Report.WindowLeft; }

        public bool Visible { get => Report.Visible; set => Report.Visible = value; }
        public string RecordSource { get => Report.RecordSource; set => Report.RecordSource = value; }

        public dynamic OpenArgs { get => Report.OpenArgs; set => Report.OpenArgs = value; }
        public bool Moveable { get => Report.Moveable; set => Report.Moveable = value; }

        public AccessReportModel(Report report)
        {
            Report = report;

            Name = report.Name;
            Controls = new Lazy<List<AccessControlModel>>(
                () => report.Controls.Cast<Control>().Select(c => new AccessControlModel(c, false, false)).ToList()
            );

            Properties = new Lazy<AccessDynamicPropertyCollectionModel>(() => new AccessDynamicPropertyCollectionModel(report.Properties));
            Module = new Lazy<ModuleModel>(() => new ModuleModel(report.Module));
        }

        public AccessReportModel(Application application, string name) : this(application.Reports[name])
        { }

        public void Requery()
        {
            Report.Requery();
        }

        public override string ToString() => Name;
    }
}