/**
*    Copyright(C) G1ANT Ltd, All rights reserved
*    Solution G1ANT.Addon, Project G1ANT.Addon.MSOffice
*    www.g1ant.com
*
*    Licensed under the G1ANT license.
*    See License.txt file in the project root for full license information.
*
*/
using G1ANT.Addon.MSOffice.Models.Access;
using Microsoft.Office.Interop.Access;

namespace G1ANT.Addon.MSOffice.Access
{
    public interface IAccessFormControlsTreeWalker
    {
        AccessControlModel GetAccessControlByPath(Application application, string path);
        AccessControlModel GetAccessControlByPath(Application application, ControlPathModel controlPath);
    }
}
