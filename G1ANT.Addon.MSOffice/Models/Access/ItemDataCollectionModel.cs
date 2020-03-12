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
using System;
using System.Collections.Generic;
using System.Linq;

namespace G1ANT.Addon.MSOffice.Models.Access
{
    public class ItemDataCollectionModel : List<ItemDataModel>
    {
        public ItemDataCollectionModel(AccessControlModel control)
        {
            var isHeaderVisible = control.TryGetPropertyValue<bool>("ColumnHeads");
            var hasItemCount = control.TryGetPropertyValue("ListCount", out int itemCount);

            if (hasItemCount)
                LoadData(control, isHeaderVisible, itemCount);
            else
                LoadFallbackData(control, isHeaderVisible);
        }

        private void LoadData(AccessControlModel control, bool isHeaderVisible, int itemCount)
        {
            AddRange(
                Enumerable.Range(isHeaderVisible ? 1 : 0, itemCount).Select(
                    i => new ItemDataModel(ToIndex(i, isHeaderVisible), control.Control.ItemData[i]?.ToString())
                )
            );
        }

        private static int ToIndex(int i, bool isHeaderVisible)
        {
            return isHeaderVisible ? i - 1 : i;
        }

        private void LoadFallbackData(AccessControlModel control, bool isHeaderVisible)
        {
            var i = 0;
            while (true)
            {
                var item = control.Control.ItemData[i]?.ToString();
                if (item == "{}" || string.IsNullOrEmpty(item))
                    break;
                if (i > 0 || !isHeaderVisible)
                    Add(new ItemDataModel(ToIndex(i, isHeaderVisible), item));

                ++i;
                if (i > 10000)
                    throw new Exception($"Error finding last data item for control {control.Name}");
            }
        }
    }
}
