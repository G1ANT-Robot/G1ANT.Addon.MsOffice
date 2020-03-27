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

namespace G1ANT.Addon.MSOffice.Models.Access.Dao.Parameters
{
    public class AccessDaoParameterModel : INameModel, IDetailedNameModel
    {
        public string Name { get; }
        public dynamic Value { get; }
        public Lazy<AccessDaoPropertyCollectionModel> Properties { get; }
        public string Type { get; }

        public AccessDaoParameterModel(Parameter parameter)
        {
            Name = parameter.Name;
            Value = parameter.Value;
            Properties = new Lazy<AccessDaoPropertyCollectionModel>(() => new AccessDaoPropertyCollectionModel(parameter.Properties));
            Type = ((DataTypeEnum)parameter.Type).ToString();
        }

        public override string ToString() => Name;

        public string ToDetailedString() => $"{Name}: {Value}, type: {Type}";
    }
}
