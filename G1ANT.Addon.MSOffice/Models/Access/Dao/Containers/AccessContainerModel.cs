/**
*    Copyright(C) G1ANT Ltd, All rights reserved
*    Solution G1ANT.Addon, Project G1ANT.Addon.MSOffice
*    www.g1ant.com
*
*    Licensed under the G1ANT license.
*    See License.txt file in the project root for full license information.
*
*/

using G1ANT.Addon.MSOffice.Models.Access.Dao.Documents;
using G1ANT.Addon.MSOffice.Models.Access.Dao.Properties;
using Microsoft.Office.Interop.Access.Dao;
using System;

namespace G1ANT.Addon.MSOffice.Models.Access.Dao.Containers
{
    internal class AccessContainerModel : INameModel
    {
        public string Name { get; }
        public Lazy<AccessDocumentCollectionModel> Documents { get; }
        public bool Inherit { get; }
        public string Owner { get; }
        public int Permissions { get; }
        public Lazy<AccessDaoPropertyCollectionModel> Properties { get; }
        public string UserName { get; }

        public AccessContainerModel(Container container)
        {
            Name = container.Name;
            //container.AllPermissions
            Documents = new Lazy<AccessDocumentCollectionModel>(
                () =>
                {
                    try { return new AccessDocumentCollectionModel(container.Documents); }
                    catch { return new AccessDocumentCollectionModel();  }
                }
            );
            Inherit = container.Inherit;
            Owner = container.Owner;
            Permissions = container.Permissions;
            Properties = new Lazy<AccessDaoPropertyCollectionModel>(
                () => {
                    try { return new AccessDaoPropertyCollectionModel(container.Properties); }
                    catch { return new AccessDaoPropertyCollectionModel(); }
                }
            );
            UserName = container.UserName;
        }

        public override string ToString() => Name;// $"{Name}, user name: {UserName}, owner: {Owner}";
    }
}
