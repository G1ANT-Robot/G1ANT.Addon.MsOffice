/**
*    Copyright(C) G1ANT Ltd, All rights reserved
*    Solution G1ANT.Addon, Project G1ANT.Addon.MSOffice
*    www.g1ant.com
*
*    Licensed under the G1ANT license.
*    See License.txt file in the project root for full license information.
*
*/
using Microsoft.Office.Interop.Access;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;

namespace G1ANT.Addon.MSOffice.Models.Access
{
    public class AccessControlModel
    {
        public string Name { get; }

        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        public dynamic Caption { get; }

        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        public dynamic Value { get; }

        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        public ICollection<AccessControlModel> Children;

        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        public AccessPropertiesModel Properties;


        public AccessControlModel(Control control, bool getProperties = true)
        {
            Name = control.Name;

            if (getProperties && control.Properties.Count > 0)
            {
                var properties = control.Properties.OfType<dynamic>().ToList();
                Caption = properties.FirstOrDefault(p => p.Name == "Caption")?.Value?.ToString();
                Value = properties.FirstOrDefault(p => p.Name == "Value")?.Value?.ToString();
                Properties = new AccessPropertiesModel(control.Properties);
            }

            try
            {
                Children = control.Controls.Count == 0 
                    ? null 
                    : control.Controls.Cast<Control>().Select(c => new AccessControlModel(c, getProperties)).ToList();
            }
            catch { }
        }
    }
}
