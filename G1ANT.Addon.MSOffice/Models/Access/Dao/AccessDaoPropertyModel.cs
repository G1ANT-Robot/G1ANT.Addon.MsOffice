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
}
