using G1ANT.Addon.MSOffice.Models.Access.Dao;
using G1ANT.Addon.MSOffice.Models.Access.Dao.Containers;
using G1ANT.Addon.MSOffice.Models.Access.Dao.Properties;
using G1ANT.Addon.MSOffice.Models.Access.Dao.QueryDefs;
using Microsoft.Office.Interop.Access.Dao;
using System;

namespace G1ANT.Addon.MSOffice.Models.Access.Forms.Recordsets
{
    internal class AccessDatabaseModel : INameModel
    {
        private readonly Database database;

        public string Name { get; }
        public string Connect { get; }
        public Lazy<AccessConnectionModel> Connection { get; }
        public Lazy<AccessContainerCollectionModel> Containers { get; }
        public Lazy<AccessDaoPropertyCollectionModel> Properties { get; }
        public Lazy<AccessQueryDefCollectionModel> QueryDefs { get; }
        public short QueryTimeout { get; }
        public int RecordsAffected { get; }
        public Lazy<AccessRecordsetCollectionModel> Recordsets { get; }
        public Lazy<AccessTableDefCollectionModel> TableDefs { get; }

        public AccessDatabaseModel(Database database)
        {
            this.database = database;
            Name = database.Name;
            Connect = database.Connect;
            Connection = new Lazy<AccessConnectionModel>(() => new AccessConnectionModel(database.Connection));
            Containers = new Lazy<AccessContainerCollectionModel>(() => new AccessContainerCollectionModel(database.Containers));
            Properties = new Lazy<AccessDaoPropertyCollectionModel>(() => new AccessDaoPropertyCollectionModel(database.Properties));
            QueryDefs = new Lazy<AccessQueryDefCollectionModel>(() => new AccessQueryDefCollectionModel(database.QueryDefs));
            QueryTimeout = database.QueryTimeout;
            RecordsAffected = database.RecordsAffected;
            Recordsets = new Lazy<AccessRecordsetCollectionModel>(() => new AccessRecordsetCollectionModel(database.Recordsets));
            TableDefs = new Lazy<AccessTableDefCollectionModel>(() => new AccessTableDefCollectionModel(database.TableDefs));
            //database.Relations
        }

        public override string ToString() => Name;
    }
}