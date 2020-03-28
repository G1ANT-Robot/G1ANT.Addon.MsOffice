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

namespace G1ANT.Addon.MSOffice.Models.Access
{
    public class AccessObjectPropertyModel : INameModel
    {
        public string Name { get; set; }
        public string Value { get; set; }

        public AccessObjectPropertyModel(AccessObjectProperty property)
        {
            Name = property.Name;
            Value = property.Value?.ToString();
        }

        public override string ToString() => Name;
    }
}
