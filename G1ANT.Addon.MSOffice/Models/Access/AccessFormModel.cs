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
    public class AccessFormModel
    {
        public string Name { get; }
        public ICollection<AccessControlModel> Controls = new List<AccessControlModel>();

        public AccessFormModel(Form form)
        {
            Name = form.Name;
            Controls = form.Controls.Cast<Control>().Select(c => new AccessControlModel(c)).ToList();
        }
    }
}
