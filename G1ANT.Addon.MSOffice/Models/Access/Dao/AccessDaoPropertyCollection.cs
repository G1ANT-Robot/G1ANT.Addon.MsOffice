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

namespace G1ANT.Addon.MSOffice.Models.Access.Dao
{
    internal class AccessDaoPropertyCollection : List<AccessDaoPropertyModel>
    {
        public AccessDaoPropertyCollection(Microsoft.Office.Interop.Access.Dao.Properties properties)
        {
            try
            {
                AddRange(properties.Cast<Microsoft.Office.Interop.Access.Dao.Property>().Select(p => new AccessDaoPropertyModel(p)));
            }
            catch { }
        }
    }
}
