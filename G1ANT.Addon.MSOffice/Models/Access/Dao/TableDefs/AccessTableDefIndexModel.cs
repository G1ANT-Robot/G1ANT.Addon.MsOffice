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

namespace G1ANT.Addon.MSOffice.Models.Access.Dao
{
    internal class AccessTableDefIndexModel : INameModel
    {
        public string Name { get; set; }
        public bool IsPrimary { get; }
        public bool IsClustered { get; }
        public bool IsForeign { get; }
        public bool IgnoreNulls { get; }
        public Lazy<AccessDaoPropertyCollectionModel> Properties { get; }
        public bool IsUnique { get; }
        public int DistinctCount { get; }
        public dynamic Fields { get; }

        public AccessTableDefIndexModel() { }

        public AccessTableDefIndexModel(Index index)
        {
            Name = index.Name;
            IsPrimary = index.Primary;
            IsClustered = index.Clustered;
            IsForeign = index.Foreign;
            IgnoreNulls = index.IgnoreNulls;
            Properties = new Lazy<AccessDaoPropertyCollectionModel>(() => new AccessDaoPropertyCollectionModel(index.Properties));
            IsUnique = index.Unique;
            DistinctCount = index.DistinctCount;
            Fields = new AccessTableDefIndexFieldCollectionModel(index.Fields);
        }
    }
}
