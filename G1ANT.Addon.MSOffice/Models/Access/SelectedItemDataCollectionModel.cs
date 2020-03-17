/**
*    Copyright(C) G1ANT Ltd, All rights reserved
*    Solution G1ANT.Addon, Project G1ANT.Addon.MSOffice
*    www.g1ant.com
*
*    Licensed under the G1ANT license.
*    See License.txt file in the project root for full license information.
*
*/

using G1ANT.Addon.MSOffice.Api.Access;
using System.Collections.Generic;
using System.Linq;

namespace G1ANT.Addon.MSOffice.Models.Access
{
    public class SelectedItemDataCollectionModel : List<ItemDataModel>
    {
        public SelectedItemDataCollectionModel(AccessControlModel control)
        {
            var isHeaderVisible = control.TryGetDynamicPropertyValue<bool>("ColumnHeads");

            AddRange(
                Enumerable
                    .Range(0, control.Control.ItemsSelected.Count)
                    .Select(i =>
                    {
                        var itemIndex = control.Control.ItemsSelected[i];
                        var itemValue = control.Control.ItemData[itemIndex]?.ToString();
                        return new ItemDataModel(isHeaderVisible ? itemIndex - 1 : itemIndex, itemValue);
                    })
            );

        }
    }
}
