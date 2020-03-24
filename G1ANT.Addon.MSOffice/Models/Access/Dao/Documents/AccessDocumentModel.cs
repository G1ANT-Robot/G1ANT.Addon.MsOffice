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

namespace G1ANT.Addon.MSOffice.Models.Access.Dao.Documents
{
    internal class AccessDocumentModel : INameModel
    {
        public string Container { get; }
        public DateTime DateCreated { get; }
        public DateTime LastUpdated { get; }
        public string Name { get; }
        public string Owner { get; }
        public int Permissions { get; }
        public AccessDaoPropertyCollectionModel Properties { get; }
        public string UserName { get; }

        public AccessDocumentModel(Document document)
        {
            try
            {
                //document.AllPermissions
                Container = document.Container;
                DateCreated = document.DateCreated;
                LastUpdated = document.LastUpdated;
                Name = document.Name;
                Owner = document.Owner;
                Permissions = document.Permissions;
                Properties = new AccessDaoPropertyCollectionModel(document.Properties);
                UserName = document.UserName;
            }
            catch { }
        }


        public override string ToString() => $"{Name}, user name: {UserName}, owner: {Owner}";
    }
}
