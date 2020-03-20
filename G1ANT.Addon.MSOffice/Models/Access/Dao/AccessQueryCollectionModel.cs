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
using System.Collections.Generic;
using System.Linq;

namespace G1ANT.Addon.MSOffice.Models.Access.Dao
{
    public class AccessQueryModel : INameModel
    {
        public string Name { get; }
        public Lazy<AccessQueryDetailsModel> Details;

        public AccessQueryModel(QueryDef query)
        {
            Name = query.Name;
            Details = new Lazy<AccessQueryDetailsModel>(() => new AccessQueryDetailsModel(query));
        }

        public override string ToString() => Name;
    }

    public class AccessQueryDetailsModel : INameModel
    {
        //public QueryDef Query { get; }
        public string Name { get; }
        public string SQL { get; }
        public string Connect { get; }
        public DateTime DateCreated { get; }
        public DateTime LastUpdated { get; }
        public AccessQueryFieldCollectionModel Fields { get; }
        public int MaxRecords { get; }
        public AccessQueryParameterCollectionModel Parameters { get; }
        public AccessQueryPropertyCollectionModel Properties { get; }
        public int RecordsAffected { get; }
        public bool ReturnsRecords { get; }
        //public bool? StillExecuting { get; }
        public string Type { get; }
        public bool Updatable { get; }

        public AccessQueryDetailsModel(QueryDef query)
        {
            //Query = query ?? throw new ArgumentNullException(nameof(query));
            try
            {
                Name = query.Name;
                SQL = query.SQL;
                Connect = query.Connect;
                DateCreated = query.DateCreated;
                LastUpdated = query.LastUpdated;
                Fields = new AccessQueryFieldCollectionModel(query.Fields);
                MaxRecords = query.MaxRecords;
                Parameters = new AccessQueryParameterCollectionModel(query.Parameters);
                Properties = new AccessQueryPropertyCollectionModel(query.Properties);
                RecordsAffected = query.RecordsAffected;
                ReturnsRecords = query.ReturnsRecords;
                //try { StillExecuting = query.StillExecuting; } catch { }
                Type = ((QueryDefTypeEnum)query.Type).ToString();
                Updatable = query.Updatable;
            }
            catch { }
        }
    }

    public class AccessQueryCollectionModel : List<AccessQueryModel>
    {

        public AccessQueryCollectionModel(RotApplicationModel rotApplicationModel) : this(rotApplicationModel.Application.CurrentDb())
        { }

        public AccessQueryCollectionModel(Database database)
        {
            AddRange(
                database.QueryDefs
                    .OfType<QueryDef>()
                    .Select(q => new AccessQueryModel(q))
            );
        }
    }
}

