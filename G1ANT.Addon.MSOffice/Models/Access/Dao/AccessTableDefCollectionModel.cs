using Microsoft.Office.Interop.Access.Dao;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace G1ANT.Addon.MSOffice.Models.Access.Dao
{
    internal class AccessDaoPropertyModel : INameModel
    {
        public string Name { get; }
        public string Value { get; }
        public short Type { get; }

        public AccessDaoPropertyModel(Property property)
        {
            Name = property.Name;
            try { Value = property.Value?.ToString(); }
            catch (Exception ex) { Value = ex.Message; }
            Type = property.Type;
        }

        public override string ToString() => $"{Name}: {Value}, type: {Type}";
    }

    internal class AccessDaoPropertyCollection : List<AccessDaoPropertyModel>
    {
        public AccessDaoPropertyCollection(Microsoft.Office.Interop.Access.Dao.Properties properties)
        {
            AddRange(properties.Cast<Microsoft.Office.Interop.Access.Dao.Property>().Select(p => new AccessDaoPropertyModel(p)));
        }
    }


    internal class AccessTableDefIndexModel : INameModel
    {
        public string Name { get; set; }
        public bool IsPrimary { get; }
        public bool IsClustered { get; }
        public bool IsForeign { get; }
        public bool IgnoreNulls { get; }
        public Lazy<AccessDaoPropertyCollection> Properties { get; }
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
            Properties = new Lazy<AccessDaoPropertyCollection>(() => new AccessDaoPropertyCollection(index.Properties));
            IsUnique = index.Unique;
            DistinctCount = index.DistinctCount;
            Fields = new AccessTableDefIndexFieldCollectionModel(index.Fields);
        }
    }

    internal class AccessTableDefIndexFieldCollectionModel : List<AccessTableDefIndexFieldModel>
    {
        public AccessTableDefIndexFieldCollectionModel(IEnumerable indexFields)
        {
            AddRange(indexFields.Cast<dynamic>().Select(f => new AccessTableDefIndexFieldModel(f)));
        }
    }

    internal class AccessTableDefIndexFieldModel : INameModel
    {
        public string Name { get; }
        public Lazy<AccessDaoPropertyCollection> Properties { get; }
        public dynamic Value { get; }

        public AccessTableDefIndexFieldModel(dynamic indexField)
        {
            try
            {
                Name = indexField.Name;
                Properties = new Lazy<AccessDaoPropertyCollection>(() => new AccessDaoPropertyCollection(indexField.Properties));
                Value = indexField?.ToString();
            }
            catch { }
        }
    }

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


    internal class AccessTableDefFieldModel : INameModel
    {
        public string Name { get; }
        public string DefaultValue { get; }
        public int FieldSize { get; }
        public string ForeignName { get; }
        public bool Required { get; }
        public int Size { get; }
        public string SourceField { get; }
        public string SourceTable { get; }
        public dynamic Value { get; }

        public AccessTableDefFieldModel(Field field)
        {
            Name = field.Name;
            DefaultValue = field.DefaultValue?.ToString();
            Required = field.Required;
            Size = field.Size;
            SourceField = field.SourceField;
            SourceTable = field.SourceTable;
        }
    }

    internal class AccessTableDefFieldCollectionModel : List<AccessTableDefFieldModel>
    {
        public AccessTableDefFieldCollectionModel(Fields fields)
        {
            AddRange(fields.Cast<Field>().Select(f => new AccessTableDefFieldModel(f)));
        }
    }

    

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
    }

    internal class AccessTableDefCollectionModel : List<AccessTableDefModel>
    {
        public AccessTableDefCollectionModel(TableDefs tableDefs)
        {
            AddRange(tableDefs.Cast<TableDef>().Select(td => new AccessTableDefModel(td)));
        }
    }
}
