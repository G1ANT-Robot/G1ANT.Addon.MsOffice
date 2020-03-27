/**
*    Copyright(C) G1ANT Ltd, All rights reserved
*    Solution G1ANT.Addon, Project G1ANT.Addon.MSOffice
*    www.g1ant.com
*
*    Licensed under the G1ANT license.
*    See License.txt file in the project root for full license information.
*
*/

using G1ANT.Addon.MSOffice.Models.Access.Dao.Properties;
using Microsoft.Office.Interop.Access.Dao;
using System;

namespace G1ANT.Addon.MSOffice.Models.Access.Dao.Fields
{
    public class AccessDaoFieldModel : INameModel
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
        public Lazy<AccessDaoPropertyCollectionModel> Properties { get; }
        public bool Required { get; }
        public int Size { get; }
        public string SourceField { get; }
        public string SourceTable { get; }
        public string Type { get; }
        //public dynamic Value { get; }
        //public dynamic VisibleValue { get; }

        public AccessDaoFieldModel(Field field)
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
            Properties = new Lazy<AccessDaoPropertyCollectionModel>(() => new AccessDaoPropertyCollectionModel(field.Properties));
            Required = field.Required;
            Size = field.Size;
            SourceField = field.SourceField;
            SourceTable = field.SourceTable;
            Type = ((DataTypeEnum)field.Type).ToString();
            //Value = field.Value;
            //VisibleValue = field.VisibleValue;
        }

        public override string ToString() => Name;
    }
}