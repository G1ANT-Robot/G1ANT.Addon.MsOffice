using G1ANT.Addon.MSOffice.Models.Access.Application.TempVars;
using G1ANT.Addon.MSOffice.Models.Access.VBE;
using System;

namespace G1ANT.Addon.MSOffice.Models.Access.Application
{
    internal class AccessApplicationModel : INameModel
    {
        public string ADOConnectString { get; }
        public int Build { get; }
        public string Name { get; }
        public string ProductCode { get; }
        public string ActiveControl { get; }
        public string ActiveDataAccessPage { get; }
        public string ActiveDatasheet { get; }
        public string ActiveForm { get; }
        public string ActiveReport { get; }
        public Lazy<VbeModel> VBE { get; }
        public Lazy<AccessCurrentProjectModel> CurrentProject { get; }
        public Lazy<AccessTempVarsCollectionModel> TempVars { get; }

        public AccessApplicationModel(Microsoft.Office.Interop.Access.Application application)
        {
            ADOConnectString = application.ADOConnectString;
            Build = application.Build;
            Name = application.Name;
            ProductCode = application.ProductCode;
            try
            {
                ActiveControl = application.Screen.ActiveControl?.Name;
                ActiveDataAccessPage = application.Screen.ActiveDataAccessPage?.Name;
                ActiveDatasheet = application.Screen.ActiveDatasheet?.Name;
                ActiveForm = application.Screen.ActiveForm?.Name;
                ActiveReport = application.Screen.ActiveReport?.Name;

            }
            catch { }

            TempVars = new Lazy<AccessTempVarsCollectionModel>(() => new AccessTempVarsCollectionModel(application.TempVars));
            VBE = new Lazy<VbeModel>(() => new VbeModel(application.VBE));
            CurrentProject = new Lazy<AccessCurrentProjectModel>(() => new AccessCurrentProjectModel(application.CurrentProject));
        }
    }
}
