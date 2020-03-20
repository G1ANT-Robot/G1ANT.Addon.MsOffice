﻿/**
*    Copyright(C) G1ANT Ltd, All rights reserved
*    Solution G1ANT.Addon, Project G1ANT.Addon.MSOffice
*    www.g1ant.com
*
*    Licensed under the G1ANT license.
*    See License.txt file in the project root for full license information.
*
*/

using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace G1ANT.Addon.MSOffice.Models.Access.Dao
{
    internal class AccessTableDefIndexFieldCollectionModel : List<AccessTableDefIndexFieldModel>
    {
        public AccessTableDefIndexFieldCollectionModel(IEnumerable indexFields)
        {
            AddRange(indexFields.Cast<dynamic>().Select(f => new AccessTableDefIndexFieldModel(f)));
        }
    }
}
