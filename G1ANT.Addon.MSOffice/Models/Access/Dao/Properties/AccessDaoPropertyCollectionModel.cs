/**
*    Copyright(C) G1ANT Ltd, All rights reserved
*    Solution G1ANT.Addon, Project G1ANT.Addon.MSOffice
*    www.g1ant.com
*
*    Licensed under the G1ANT license.
*    See License.txt file in the project root for full license information.
*
*/

using Microsoft.Office.Interop.Access.Dao;
using System.Collections.Generic;

namespace G1ANT.Addon.MSOffice.Models.Access.Dao.Properties
{
    public class AccessDaoPropertyCollectionModel : List<AccessDaoPropertyModel>
    {
        public AccessDaoPropertyCollectionModel()
        { }

        public AccessDaoPropertyCollectionModel(Microsoft.Office.Interop.Access.Dao.Properties properties)
        {
            try
            {
                foreach (Property property in properties)
                {
                    try
                    {
                        var model = new AccessDaoPropertyModel(property);
                        Add(model);
                    }
                    catch { }
                }
            }
            catch { }
        }
    }
}
