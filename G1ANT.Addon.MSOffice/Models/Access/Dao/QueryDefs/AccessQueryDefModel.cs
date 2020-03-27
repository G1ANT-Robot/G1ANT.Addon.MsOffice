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
using System.Text;

namespace G1ANT.Addon.MSOffice.Models.Access.Dao.QueryDefs
{
    public class AccessQueryDefModel : INameModel, IDetailedNameModel
    {
        public string Name { get; }
        public Lazy<AccessQueryDefDetailsModel> Details;

        public AccessQueryDefModel(QueryDef queryDef)
        {
            Name = queryDef.Name;
            Details = new Lazy<AccessQueryDefDetailsModel>(() => new AccessQueryDefDetailsModel(queryDef));
        }


        public override string ToString() => Name;

        public string ToDetailedString()
        {
            var result = new StringBuilder();

            result.AppendLine($"Name: {Name}");
            //result.AppendLine($"Type: {Type}");
            //result.AppendLine($"DateCreated: {DateCreated}");
            //result.AppendLine($"DateModified: {LastUpdated}");
            //result.AppendLine($"Connect: {Connect}");
            ////result.AppendLine($"Fields: {string.Join(", ", Fields.Select(f => f.Name))}");
            //result.AppendLine($"RecordsAffected: {RecordsAffected}");
            //result.AppendLine($"ReturnsRecords: {ReturnsRecords}");
            //result.AppendLine($"SQL: {SQL}");
            //result.AppendLine($"Type: {Type}");
            //result.AppendLine($"Updatable: {Updatable}");

            return result.ToString();
        }
    }
}

