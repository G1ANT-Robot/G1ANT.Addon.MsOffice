/**
*    Copyright(C) G1ANT Ltd, All rights reserved
*    Solution G1ANT.Addon, Project G1ANT.Addon.MSOffice
*    www.g1ant.com
*
*    Licensed under the G1ANT license.
*    See License.txt file in the project root for full license information.
*
*/

using Microsoft.Office.Interop.Access.Dao;
using System;
using System.Collections.Generic;
using System.Linq;

namespace G1ANT.Addon.MSOffice.Models.Access.Dao
{
    internal class AccessTableDefIndexCollectionModel : List<AccessTableDefIndexModel>
    {
        public AccessTableDefIndexCollectionModel(Indexes indexes)
        {
            try
            {
                AddRange(indexes.Cast<Index>().Select(i => new AccessTableDefIndexModel(i)));
            }
            catch (Exception ex)
            {
                Add(new AccessTableDefIndexModel() { Name = ex.Message });
            }
        }
    }
}
