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
    public class AccessFormModel : IComparable, INameModel
    {
        public string Name { get; }

        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        public string Value { get; }

        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        public string Caption { get; }

        [JsonIgnore]
        public Form Form { get; }

        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        public string FormName { get; }

        public int Hwnd { get; }
        public int InsideWidth { get; }
        public short Width { get; }
        public short Height { get; }
        public short X { get; }
        public short Y { get; }
        public int InsideHeight { get; }

        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        public ICollection<AccessControlModel> Controls { get; private set; }

        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        public AccessPropertiesModel Properties { get; }


        public AccessFormModel(Form form, bool getFormProperties, bool getControls, bool getControlsProperties)
        {
            Form = form ?? throw new ArgumentNullException(nameof(form));

            Name = form.Name;
            Value = form.accValue;

            Caption = form.Caption;
            FormName = form.FormName;
            Hwnd = form.Hwnd;
            InsideWidth = form.InsideWidth;
            InsideHeight = form.InsideHeight;
            Width = form.WindowWidth;
            Height = form.WindowHeight;
            X = form.WindowLeft;
            Y = form.WindowTop;

            Properties = !getFormProperties || form.Properties.Count == 0 ? null : new AccessPropertiesModel(form.Properties);
            if (getControls)
                LoadControls(getControlsProperties);
        }


        public void LoadControls(bool getControlsProperties)
        {
            if (Form.Controls.Count > 0)
                Controls = Form.Controls.Cast<Control>().Select(c => new AccessControlModel(c, getControlsProperties, false)).ToList();
        }

        public int CompareTo(object obj)
        {
            if (obj == null)
                return 1;

            if (!(obj is AccessFormModel))
                return 1;

            var model = (AccessFormModel)obj;

            if (Form.Application.hWndAccessApp() != model.Form.Application.hWndAccessApp())
                return 1;

            return model.Name == this.Name ? 0 : 1; // names of forms seem to be unique
        }
    }
}
