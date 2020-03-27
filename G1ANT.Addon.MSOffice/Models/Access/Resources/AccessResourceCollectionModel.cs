using Microsoft.Office.Interop.Access;
using System.Collections.Generic;

namespace G1ANT.Addon.MSOffice.Models.Access.Resources
{
    internal class AccessResourceCollectionModel : List<AccessResourceModel>
    {
        public AccessResourceCollectionModel(SharedResources resources)
        {
            foreach (SharedResource resource in resources)
            {
                try { Add(new AccessResourceModel(resource)); }
                catch { }
            }
        }
    }
}
