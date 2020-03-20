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
using System.Linq;

namespace G1ANT.Addon.MSOffice.Models.Access.Dao.Containers
{
    internal class AccessContainerCollectionModel : List<AccessContainerModel>
    {
        public AccessContainerCollectionModel(Microsoft.Office.Interop.Access.Dao.Containers containers)
        {
            try
            {
                foreach (Container container in containers)
                {
                    try
                    {
                        var model = new AccessContainerModel(container);
                        Add(model);
                    }
                    catch { }
                }
            }
            catch { }
            //AddRange(containers.Cast<Container>().Select(c => new AccessContainerModel(c)));
        }
    }
}
