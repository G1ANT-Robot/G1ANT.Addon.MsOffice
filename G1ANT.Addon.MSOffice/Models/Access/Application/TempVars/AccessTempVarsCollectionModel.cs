using Microsoft.Office.Interop.Access;
using System.Collections.Generic;
using System.Linq;

namespace G1ANT.Addon.MSOffice.Models.Access.Application.TempVars
{
    internal class AccessTempVarsCollectionModel : List<AccessTempVarsModel>
    {
        public AccessTempVarsCollectionModel(Microsoft.Office.Interop.Access.TempVars tempVars)
        {
            AddRange(tempVars.Cast<TempVar>().Select(tv => new AccessTempVarsModel(tv)));
        }

        public override string ToString() => "TempVars";
    }
}
