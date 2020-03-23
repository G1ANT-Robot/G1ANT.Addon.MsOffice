/**
*    Copyright(C) G1ANT Ltd, All rights reserved
*    Solution G1ANT.Addon, Project G1ANT.Addon.MSOffice
*    www.g1ant.com
*
*    Licensed under the G1ANT license.
*    See License.txt file in the project root for full license information.
*
*/

using G1ANT.Addon.MSOffice.Models.Access.AccessObjects;
using Microsoft.Office.Interop.Access;

namespace G1ANT.Addon.MSOffice.Models.Access.Data
{
    public class AccessObjectFunctionCollectionModel : AccessObjectCollectionModel
    {
        public AccessObjectFunctionCollectionModel(RotApplicationModel rotApplicationModel)
            : this(rotApplicationModel.Application.CurrentData.AllFunctions)
        { }
        
            
        public AccessObjectFunctionCollectionModel(AllObjects functions)
        {
            try
            {
                Initialize(functions);
            }
            catch
            { }
        }
    }
}
