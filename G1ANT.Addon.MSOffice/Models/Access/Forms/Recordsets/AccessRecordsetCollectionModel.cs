using Microsoft.Office.Interop.Access.Dao;
using System.Collections.Generic;
using System.Linq;

namespace G1ANT.Addon.MSOffice.Models.Access.Forms.Recordsets
{
    internal class AccessRecordsetCollectionModel : List<AccessRecordsetModel>
    {
        public AccessRecordsetCollectionModel(Microsoft.Office.Interop.Access.Dao.Recordsets recordsets)
        {
            AddRange(
                recordsets
                    .Cast<Recordset>()
                    .Select(r => new AccessRecordsetModel(r))
            );
        }
    }
}
