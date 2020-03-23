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

namespace G1ANT.Addon.MSOffice.Models.Access
{
    internal class AccessDynamicPropertyCollectionModel : List<AccessDynamicPropertyModel>
    {
        public AccessDynamicPropertyCollectionModel(IEnumerable<dynamic> properties)
        {
            foreach (var property in properties.Where(p => p.Name != "InSelection").OrderBy(p => p.Name))
            {
                Add(new AccessDynamicPropertyModel(property));
            }
        }

        public AccessDynamicPropertyCollectionModel(Microsoft.Office.Interop.Access.Properties properties) : this(properties.OfType<dynamic>())
        { }
    }
}
