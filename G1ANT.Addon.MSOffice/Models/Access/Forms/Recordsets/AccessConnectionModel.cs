using G1ANT.Addon.MSOffice.Models.Access.Dao.QueryDefs;
using Microsoft.Office.Interop.Access.Dao;
using System;

namespace G1ANT.Addon.MSOffice.Models.Access.Forms.Recordsets
{
    internal class AccessConnectionModel : INameModel
    {
        public Connection Connection { get; }
        public string Connect { get; }
        public string Name { get; }
        public Lazy<AccessDatabaseModel> Database { get; }
        public int HDbc { get; }
        public bool Transactions { get; }
        public int RecordsAffected { get; }
        public bool StillExecuting { get; }
        public short QueryTimeout { get => Connection.QueryTimeout; set => Connection.QueryTimeout = value; }
        public bool Updatable { get; }
        public Lazy<AccessQueryDefCollectionModel> QueryDefs { get; }
        public Lazy<AccessRecordsetCollectionModel> Recordsets { get; }

        public AccessConnectionModel(Connection connection)
        {
            Connection = connection;

            Connect = connection.Connect;
            Name = connection.Name;
            Database = new Lazy<AccessDatabaseModel>(() => new AccessDatabaseModel(connection.Database));
            HDbc = connection.hDbc;
            Transactions = connection.Transactions;
            RecordsAffected = connection.RecordsAffected;
            try { StillExecuting = connection.StillExecuting; } catch { }
            Updatable = connection.Updatable;
            QueryDefs = new Lazy<AccessQueryDefCollectionModel>(() => new AccessQueryDefCollectionModel(connection.QueryDefs));
            Recordsets = new Lazy<AccessRecordsetCollectionModel>(() => new AccessRecordsetCollectionModel(connection.Recordsets));
        }

        public void Cancel() => Connection.Cancel();
        public void Close() => Connection.Close();

        public AccessQueryDefModel CreateQueryDef(object name, string sqlText) => new AccessQueryDefModel(Connection.CreateQueryDef(name, sqlText));

        //void Execute(string query, object options) => Connection.Execute(query, options);

        AccessRecordsetModel OpenRecordset(string name, RecordsetTypeEnum type, RecordsetOptionEnum options, LockTypeEnum lockEdit)
        {
            return new AccessRecordsetModel(Connection.OpenRecordset(name, type, options, lockEdit));
        }
    }
}