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
    internal class AccessTableDefModel : INameModel
    {
        public string Name { get; }
        public TableDefAttributeEnum Attributes { get; }
        public Lazy<AccessDaoPropertyCollection> Properties { get; }
        public Lazy<AccessTableDefIndexCollectionModel> Indexes { get; }
        public Lazy<AccessTableDefFieldCollectionModel> Fields { get; }
        public int RecordCount { get; }
        public string SourceTableName { get; }
        public string Connect { get; }
        public DateTime DateCreated { get; }
        public DateTime LastUpdated { get; }
        public bool Updatable { get; }

        public AccessTableDefModel(TableDef tableDef)
        {
            Name = tableDef.Name;

            Attributes = (TableDefAttributeEnum)tableDef.Attributes;
            Properties = new Lazy<AccessDaoPropertyCollection>(() => new AccessDaoPropertyCollection(tableDef.Properties));
            Indexes = new Lazy<AccessTableDefIndexCollectionModel>(() => new AccessTableDefIndexCollectionModel(tableDef.Indexes));
            Fields = new Lazy<AccessTableDefFieldCollectionModel>(() => new AccessTableDefFieldCollectionModel(tableDef.Fields));

            RecordCount = tableDef.RecordCount;
            SourceTableName = tableDef.SourceTableName;
            Connect = tableDef.Connect;
            DateCreated = tableDef.DateCreated;
            LastUpdated = tableDef.LastUpdated;
            Updatable = tableDef.Updatable;
        }

        public override string ToString() => Name;
    }
}
