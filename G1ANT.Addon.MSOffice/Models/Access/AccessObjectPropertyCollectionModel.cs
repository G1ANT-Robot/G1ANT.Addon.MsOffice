/**
*    Copyright(C) G1ANT Ltd, All rights reserved
*    Solution G1ANT.Addon, Project G1ANT.Addon.MSOffice
*    www.g1ant.com
*
*    Licensed under the G1ANT license.
*    See License.txt file in the project root for full license information.
*
*/

using Microsoft.Office.Interop.Access;
using System.Collections.Generic;

namespace G1ANT.Addon.MSOffice.Models.Access
{
    public class AccessObjectPropertyCollectionModel : List<AccessObjectPropertyModel>
    {
        public string Name { get; set; }
        public string Value { get; set; }

        public AccessObjectPropertyCollectionModel(AccessObjectProperties properties)
        {
            try
            {
                foreach (AccessObjectProperty property in properties)
                {
                    try
                    {
                        var model = new AccessObjectPropertyModel(property);
                        Add(model);
                    }
                    catch { }
                }
            }
            catch { }
        }
    }
}
