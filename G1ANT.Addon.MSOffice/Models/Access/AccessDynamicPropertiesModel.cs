/**
*    Copyright(C) G1ANT Ltd, All rights reserved
*    Solution G1ANT.Addon, Project G1ANT.Addon.MSOffice
*    www.g1ant.com
*
*    Licensed under the G1ANT license.
*    See License.txt file in the project root for full license information.
*
*/

using System.Collections.Generic;
using System.Linq;

namespace G1ANT.Addon.MSOffice.Models.Access
{
    public class AccessDynamicPropertiesModel : Dictionary<string, string>
    {
        public AccessDynamicPropertiesModel(IEnumerable<dynamic> properties)
        {
            properties.Select(p =>
            {
                try
                {
                    if (p.Name != "InSelection")
                    {
                        return new
                        {
                            Name = ((object)p.Name).ToString(),
                            Value = ((object)p.Value).ToString()
                        };
                    }
                }
                catch { }
                return null;
            })
            .Where(p => p != null && !string.IsNullOrEmpty(p.Value))
            .ToList()
            .ForEach(p => Add(p.Name, p.Value));
        }

        public AccessDynamicPropertiesModel(Microsoft.Office.Interop.Access.Properties properties) : this(properties.OfType<dynamic>())
        { }
    }
}
