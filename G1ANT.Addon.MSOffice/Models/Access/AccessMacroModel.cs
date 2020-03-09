/**
*    Copyright(C) G1ANT Ltd, All rights reserved
*    Solution G1ANT.Addon, Project G1ANT.Addon.MSOffice
*    www.g1ant.com
*
*    Licensed under the G1ANT license.
*    See License.txt file in the project root for full license information.
*
*/

using Newtonsoft.Json;
using System;

namespace G1ANT.Addon.MSOffice.Models.Access
{
    public class AccessMacroModel
    {
        public string Name { get; }

        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        public dynamic CurrentView { get; }

        public int Attributes { get; }
        public DateTime DateCreated { get; }
        public DateTime DateModified { get; }
        public string FullName { get; }
        public dynamic Type { get; }

        public AccessMacroModel(dynamic macro)
        {
            Name = macro.Name;
            try { CurrentView = macro.CurrentView; } catch { }
            Attributes = macro.Attributes;
            DateCreated = macro.DateCreated;
            DateModified = macro.DateModified;
            FullName = macro.FullName;
            Type = macro.Type;

            // if (macro.Properties.Count > 0) Properties = 
        }
    }
}
