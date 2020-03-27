/**
*    Copyright(C) G1ANT Ltd, All rights reserved
*    Solution G1ANT.Addon, Project G1ANT.Addon.MSOffice
*    www.g1ant.com
*
*    Licensed under the G1ANT license.
*    See License.txt file in the project root for full license information.
*
*/

using G1ANT.Addon.MSOffice.Models.Access.Dao.Fields;
using G1ANT.Addon.MSOffice.Models.Access.Dao.Parameters;
using G1ANT.Addon.MSOffice.Models.Access.Dao.Properties;
using Microsoft.Office.Interop.Access.Dao;
using System;

namespace G1ANT.Addon.MSOffice.Models.Access.Dao.QueryDefs
{
    public class AccessQueryDefDetailsModel : INameModel
    {
        //public QueryDef Query { get; }
        public string Name { get; }
        public string SQL { get; }
        public string Connect { get; }
        public DateTime DateCreated { get; }
        public DateTime LastUpdated { get; }
        public Lazy<AccessDaoFieldCollectionModel> Fields { get; }
        public int MaxRecords { get; }
        public Lazy<AccessDaoParameterCollectionModel> Parameters { get; }
        public Lazy<AccessDaoPropertyCollectionModel> Properties { get; }
        public int RecordsAffected { get; }
        public bool ReturnsRecords { get; }
        //public bool? StillExecuting { get; }
        public string Type { get; }
        public bool Updatable { get; }

        public AccessQueryDefDetailsModel(QueryDef query)
        {
            //Query = query ?? throw new ArgumentNullException(nameof(query));
            try
            {
                Name = query.Name;
                SQL = query.SQL;
                Connect = query.Connect;
                DateCreated = query.DateCreated;
                LastUpdated = query.LastUpdated;
                Fields = new Lazy<AccessDaoFieldCollectionModel>(() => new AccessDaoFieldCollectionModel(query.Fields));
                MaxRecords = query.MaxRecords;
                Parameters = new Lazy<AccessDaoParameterCollectionModel>(() => new AccessDaoParameterCollectionModel(query.Parameters));
                Properties = new Lazy<AccessDaoPropertyCollectionModel>(() => new AccessDaoPropertyCollectionModel(query.Properties));
                RecordsAffected = query.RecordsAffected;
                ReturnsRecords = query.ReturnsRecords;
                //try { StillExecuting = query.StillExecuting; } catch { }
                Type = ((QueryDefTypeEnum)query.Type).ToString();
                Updatable = query.Updatable;
            }
            catch { }
        }
    }
}

