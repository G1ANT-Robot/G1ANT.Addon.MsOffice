/**
*    Copyright(C) G1ANT Ltd, All rights reserved
*    Solution G1ANT.Addon, Project G1ANT.Addon.MSOffice
*    www.g1ant.com
*
*    Licensed under the G1ANT license.
*    See License.txt file in the project root for full license information.
*
*/

using System.Collections.Generic;
using System.Linq;

namespace G1ANT.Addon.MSOffice.Models.Access.Dao.Parameters
{
    public class AccessDaoParameterCollectionModel : List<AccessDaoParameterModel>
    {
        public AccessDaoParameterCollectionModel(Microsoft.Office.Interop.Access.Dao.Parameters parameters)
        {
            AddRange(parameters.Cast<Microsoft.Office.Interop.Access.Dao.Parameter>().Select(p => new AccessDaoParameterModel(p)));
        }
    }
}
