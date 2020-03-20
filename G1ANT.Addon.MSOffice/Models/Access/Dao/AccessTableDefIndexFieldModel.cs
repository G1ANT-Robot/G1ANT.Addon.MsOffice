﻿/**
*    Copyright(C) G1ANT Ltd, All rights reserved
*    Solution G1ANT.Addon, Project G1ANT.Addon.MSOffice
*    www.g1ant.com
*
*    Licensed under the G1ANT license.
*    See License.txt file in the project root for full license information.
*
*/

using System;

namespace G1ANT.Addon.MSOffice.Models.Access.Dao
{
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
}
