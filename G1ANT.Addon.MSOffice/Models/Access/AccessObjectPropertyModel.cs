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
using System.Linq;

namespace G1ANT.Addon.MSOffice.Models.Access
{
    public class AccessObjectPropertyModel
    {
        public string Name { get; set; }
        public string Value { get; set; }
        //public ICollection<AccessObjectPropertyModel> Children { get; set; } = new List<AccessObjectPropertyModel>();

        public AccessObjectPropertyModel(AccessObjectProperty property)
        {
            Name = property.Name;
            Value = property.Value?.ToString();

            //if (property.Value is AccessObjectProperties properties)
            //{
            //    Children = properties.Cast<AccessObjectProperty>()
            //        .Select(p => new AccessObjectPropertyModel(p))
            //        .ToList();
            //}
        }
    }
}
