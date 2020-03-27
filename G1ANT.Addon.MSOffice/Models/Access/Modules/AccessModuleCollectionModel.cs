using Microsoft.Office.Interop.Access;
using System.Collections.Generic;

namespace G1ANT.Addon.MSOffice.Models.Access.Modules
{
    internal class AccessModuleCollectionModel : List<AccessModuleModel>
    {
        public AccessModuleCollectionModel(Microsoft.Office.Interop.Access.Modules modules)
        {
            foreach (Module module in modules)
            {
                try { Add(new AccessModuleModel(module)); }
                catch { }
            }
        }
    }
}
