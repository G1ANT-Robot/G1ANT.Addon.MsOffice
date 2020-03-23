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

namespace G1ANT.Addon.MSOffice.Models.Access.AccessObjects
{
    public class AccessObjectCollectionModel : List<AccessObjectModel>
    {
        protected void Initialize(AllObjects allObjects)
        {
            AddRange(allObjects.OfType<AccessObject>().Select(o => new AccessObjectModel(o)));
        }
    }
}
