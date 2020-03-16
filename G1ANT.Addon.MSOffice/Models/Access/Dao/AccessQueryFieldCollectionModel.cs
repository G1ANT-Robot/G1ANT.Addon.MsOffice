using Microsoft.Office.Interop.Access.Dao;
using System.Collections.Generic;
using System.Linq;

namespace G1ANT.Addon.MSOffice.Models.Access.Dao
{
    public class AccessQueryFieldCollectionModel : List<AccessQueryFieldModel>
    {
        public AccessQueryFieldCollectionModel(Fields fields)
        {
            AddRange(fields.Cast<Field>().Select(f => new AccessQueryFieldModel(f)));
        }
    }
}