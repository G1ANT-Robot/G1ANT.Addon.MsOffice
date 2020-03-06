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
using System.Collections.Generic;
using System.Linq;

namespace G1ANT.Addon.MSOffice.Models.Access
{
    public class AccessFormModel
    {
        public string Name { get; }
        public string Value { get; }
        public string Caption { get; }
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        public AccessFormModel Form { get; }
        public string FormName { get; }
        public int Hwnd { get; }
        public int InsideWidth { get; }
        public short Width { get; }
        public short Height { get; }
        public short X { get; }
        public short Y { get; }
        public int InsideHeight { get; }

        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        public ICollection<AccessControlModel> Controls;

        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        public AccessPropertiesModel Properties { get; }


        public AccessFormModel(Form form, bool getProperties)
        {
            Name = form.Name;
            Value = form.accValue;

            Caption = form.Caption;
            Form = form.Form != null && form.Form != form ? new AccessFormModel(form.Form, getProperties) : null;
            FormName = form.FormName;
            Hwnd = form.Hwnd;
            InsideWidth = form.InsideWidth;
            InsideHeight = form.InsideHeight;
            Width = form.WindowWidth;
            Height = form.WindowHeight;
            X = form.WindowLeft;
            Y = form.WindowTop;

            Properties = getProperties && form.Properties.Count == 0 ? null : new AccessPropertiesModel(form.Properties);
            Controls = form.Controls.Count == 0 ? null : form.Controls.Cast<Control>().Select(c => new AccessControlModel(c, getProperties)).ToList();
        }
    }
}
