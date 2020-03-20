﻿/**
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
using System.Linq;

namespace G1ANT.Addon.MSOffice.Models.Access.Dao
{
    public class AccessQueryParameterCollectionModel : List<AccessQueryParameterModel>
    {
        public AccessQueryParameterCollectionModel(Parameters parameters)
        {
            AddRange(parameters.Cast<Parameter>().Select(p => new AccessQueryParameterModel(p)));
        }
    }
}