/**
*    Copyright(C) G1ANT Ltd, All rights reserved
*    Solution G1ANT.Addon, Project G1ANT.Addon.MSOffice
*    www.g1ant.com
*
*    Licensed under the G1ANT license.
*    See License.txt file in the project root for full license information.
*
*/

using G1ANT.Addon.MSOffice.Models.Access.Dao;
using G1ANT.Language;
using Microsoft.Office.Interop.Access.Dao;
using System;
using System.Collections.Generic;
using System.Linq;

namespace G1ANT.Addon.MSOffice.Models.Access
{
    public class AccessQueryModel
    {
        public QueryDef Query { get; }
        public string Name { get; }
        public string SQL { get; }
        public string Connect { get; }
        public DateTime DateCreated { get; }
        public DateTime LastUpdated { get; }
        public AccessQueryFieldCollectionModel Fields { get; }
        public int MaxRecords { get; }
        public AccessQueryParameterCollectionModel Parameters { get; }
        public dynamic Prepare { get; }
        public AccessQueryPropertyCollectionModel Properties { get; }
        public int RecordsAffected { get; }
        public bool ReturnsRecords { get; }
        public bool StillExecuting { get; }
        public string Type { get; }
        public bool Updatable { get; }

        public AccessQueryModel(QueryDef query)
        {
            Query = query ?? throw new ArgumentNullException(nameof(query));

            Name = query.Name;
            SQL = query.SQL;
            Connect = query.Connect;
            DateCreated = query.DateCreated;
            LastUpdated = query.LastUpdated;
            Fields = new AccessQueryFieldCollectionModel(query.Fields);
            MaxRecords = query.MaxRecords;
            Parameters = new AccessQueryParameterCollectionModel(query.Parameters);
            Prepare = query.Prepare?.ToString();
            Properties = new AccessQueryPropertyCollectionModel(query.Properties);
            RecordsAffected = query.RecordsAffected;
            ReturnsRecords = query.ReturnsRecords;
            StillExecuting = query.StillExecuting;
            Type = ((QueryDefTypeEnum)query.Type).ToString();
            Updatable = query.Updatable;
        }

    }

    public class AccessQueryCollectionModel : List<AccessQueryModel>
    {
        public AccessQueryCollectionModel(RotApplicationModel rotApplicationModel)
        {
            try
            {
                AddRange(
                    rotApplicationModel.Application.CurrentDb().QueryDefs
                        .OfType<QueryDef>()
                        .Select(q => new AccessQueryModel(q))
                );
            }
            catch (Exception ex)
            {
                RobotMessageBox.Show(ex.Message);
            }
        }
    }
}
