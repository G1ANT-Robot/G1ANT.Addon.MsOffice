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
    public class AccessControlModel
    {
        public string Name { get; }
        public ICollection<AccessControlModel> Children = new List<AccessControlModel>();

        public AccessControlModel(Control control)
        {
            Name = control.Name;

            try
            {
                Children = control.Controls.Cast<Control>().Select(c => new AccessControlModel(c)).ToList();
            }
            catch { }
        }
    }
}
