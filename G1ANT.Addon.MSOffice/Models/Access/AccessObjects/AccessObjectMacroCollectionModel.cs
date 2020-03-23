/**
*    Copyright(C) G1ANT Ltd, All rights reserved
*    Solution G1ANT.Addon, Project G1ANT.Addon.MSOffice
*    www.g1ant.com
*
*    Licensed under the G1ANT license.
*    See License.txt file in the project root for full license information.
*
*/

namespace G1ANT.Addon.MSOffice.Models.Access.AccessObjects
{
    public class AccessObjectMacroCollectionModel : AccessObjectCollectionModel
    {
        public AccessObjectMacroCollectionModel(RotApplicationModel rotApplicationModel)
        {
            try
            {
                Initialize(rotApplicationModel.Application.CurrentProject.AllMacros);
            }
            catch
            { }
        }
    }
}
