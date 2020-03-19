using Microsoft.Office.Interop.Access.Dao;
using System;

namespace G1ANT.Addon.MSOffice.Models.Access.Dao
{
    public class AccessQueryFieldModel
    {
        public string Name { get; }
        public string Attributes { get; }
        public short CollectionIndex { get; }
        public bool DataUpdatable { get; }
        public dynamic DefaultValue { get; }
        //public int FieldSize { get; }
        //public string ForeignName { get; }
        public short OrdinalPosition { get; }
        //public dynamic OriginalValue { get; }
        public Lazy<AccessQueryPropertyCollectionModel> Properties { get; }
        public bool Required { get; }
        public int Size { get; }
        public string SourceField { get; }
        public string SourceTable { get; }
        public string Type { get; }
        //public dynamic Value { get; }
        //public dynamic VisibleValue { get; }

        public AccessQueryFieldModel(Field field)
        {
            Name = field.Name;
            Attributes = ((FieldAttributeEnum)field.Attributes).ToString();
            CollectionIndex = field.CollectionIndex;
            DataUpdatable = field.DataUpdatable;
            DefaultValue = field.DefaultValue;
            //FieldSize = field.FieldSize;
            //ForeignName = field.ForeignName;
            OrdinalPosition = field.OrdinalPosition;
            //OriginalValue = field.OriginalValue;
            Properties = new Lazy<AccessQueryPropertyCollectionModel>(() => new AccessQueryPropertyCollectionModel(field.Properties));
            Required = field.Required;
            Size = field.Size;
            SourceField = field.SourceField;
            SourceTable = field.SourceTable;
            Type = ((DataTypeEnum)field.Type).ToString();
            //Value = field.Value;
            //VisibleValue = field.VisibleValue;
        }
    }
}