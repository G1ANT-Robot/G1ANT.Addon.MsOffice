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
    public class AccessObjectDetailedModel : AccessObjectModel
    {
        public ICollection<AccessObjectPropertyModel> Properties { get; set; } = new List<AccessObjectPropertyModel>();

        public AccessObjectDetailedModel(AccessObject form) : base(form)
        {
            Properties = form.Properties.Cast<AccessObjectProperty>().Select(p => new AccessObjectPropertyModel(p.Value)).ToList();
        }
    }
}
