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

namespace G1ANT.Addon.MSOffice.Models.Access.Dao
{
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
}
