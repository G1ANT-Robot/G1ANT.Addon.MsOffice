/**
*    Copyright(C) G1ANT Ltd, All rights reserved
*    Solution G1ANT.Addon, Project G1ANT.Addon.MSOffice
*    www.g1ant.com
*
*    Licensed under the G1ANT license.
*    See License.txt file in the project root for full license information.
*
*/

namespace G1ANT.Addon.MSOffice.Models.Access
{
    public class AccessObjectQueryCollectionModel : AccessObjectCollectionModel
    {
        public AccessObjectQueryCollectionModel(RotApplicationModel rotApplicationModel)
        {
            try
            {
                var currentData = rotApplicationModel.Application.CurrentData;
                Initialize(currentData.AllQueries);
            }
            catch
            { }
        }
    }
}
