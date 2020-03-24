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
using System.Collections.Generic;

namespace G1ANT.Addon.MSOffice.Models.Access.Dao.Documents
{
    internal class AccessDocumentCollectionModel : List<AccessDocumentModel>
    {
        public AccessDocumentCollectionModel()
        { }

        public AccessDocumentCollectionModel(Microsoft.Office.Interop.Access.Dao.Documents documents)
        {
            try
            {
                foreach (Document document in documents)
                {
                    try
                    {
                        var model = new AccessDocumentModel(document);
                        Add(model);
                    }
                    catch { }
                }
            }
            catch { }
        }
    }
}
