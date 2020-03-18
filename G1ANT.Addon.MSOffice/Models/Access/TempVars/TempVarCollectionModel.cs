using Microsoft.Office.Interop.Access;
using System.Collections.Generic;
using System.Linq;

namespace G1ANT.Addon.MSOffice.Models.Access.VBE
{
    internal class TempVarCollectionModel : List<TempVarModel>
    {
        public TempVarCollectionModel(TempVars tempVars)
        {
            AddRange(tempVars.Cast<TempVar>().Select(tv => new TempVarModel(tv)));
        }
    }
}
