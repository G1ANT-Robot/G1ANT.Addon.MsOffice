using G1ANT.Addon.MSOffice.Models.Access.VBE;
using Microsoft.Office.Interop.Access;

namespace G1ANT.Addon.MSOffice.Models.Access
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
        public VbeModel VBE { get; }
        public AccessCurrentProjectModel CurrentProject { get; }

        public AccessApplicationModel(Application application)
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

            VBE = new VbeModel(application.VBE);
            CurrentProject = new AccessCurrentProjectModel(application?.CurrentProject);
        }
    }
}
