/**
*    Copyright(C) G1ANT Ltd, All rights reserved
*    Solution G1ANT.Addon, Project G1ANT.Addon.MSOffice
*    www.g1ant.com
*
*    Licensed under the G1ANT license.
*    See License.txt file in the project root for full license information.
*
*/

using G1ANT.Addon.MSOffice.Helpers.Access;
using Microsoft.Office.Interop.Access;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;

namespace G1ANT.Addon.MSOffice.Models.Access
{
    public class AccessControlModel : IComparable, INameModel
    {
        public string Name { get; }
        public string Type { get; }
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        public dynamic Caption { get; private set; }

        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        public dynamic Value { get; private set; }

        [JsonIgnore]
        public _Control Control { get; }

        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        public ICollection<AccessControlModel> Children;

        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        public AccessDynamicPropertiesModel Properties;


        public AccessControlModel(_Control control, bool getProperties = true, bool getChildrenRecursively = true)
        {
            Control = control ?? throw new ArgumentNullException(nameof(control));
            Name = control.Name;
            Type = ((AcControlType)this.TryGetDynamicPropertyValue<int>("ControlType")).ToString();
            Caption = this.TryGetDynamicPropertyValue<string>("Caption");
            Value = this.TryGetDynamicPropertyValue<string>("Value");

            if (getProperties && control.Properties.Count > 0)
                LoadProperties();

            if (getChildrenRecursively)
                LoadChildren(getProperties);
        }

        public void LoadProperties()
        {
            Properties = new AccessDynamicPropertiesModel(Control.Properties);
        }

        internal void SetValue(string value)
        {
            try
            {
                ((dynamic)Control).Value = value;
            }
            catch (COMException ex)
            {
                throw new Exception("Error setting the value", ex);
            }
        }

        internal object GetValue()
        {
            try
            {
                return ((dynamic)Control).Value;
            }
            catch (COMException ex)
            {
                throw new Exception("Error getting the value", ex);
            }
        }

        internal List<NameValueModel> GetProperties(bool getValues = true)
        {
            return TypeDescriptor.GetProperties(Control)
                .Cast<PropertyDescriptor>()
                .Select(pd => new NameValueModel(pd.Name, getValues ? pd.GetValue(Control) : null))
                .ToList();
        }

        internal AccessDynamicPropertiesModel GetDynamicProperties() => new AccessDynamicPropertiesModel(Control.Properties);

        internal T GetPropertyValue<T>(string name) => (T)Convert.ChangeType(GetPropertyValue(name), typeof(T));

        internal object GetPropertyValue(string name)
        {
            try
            {
                var property = TypeDescriptor.GetProperties(Control)[name];
                return property.GetValue(Control);
            }
            catch (COMException ex)
            {
                throw new Exception($"Error getting the property {name} value", ex);
            }
        }

        internal void SetPropertyValue(string name, object value)
        {
            try
            {
                var property = TypeDescriptor.GetProperties(Control)[name];
                property.SetValue(Control, value);
            }
            catch (COMException ex)
            {
                throw new Exception($"Error setting the value of property {name}", ex);
            }
        }

        public void SetFocus() => Control.SetFocus();

        public AccessControlModel GetParent()
        {
            if (Control.Parent is _Control)
                return new AccessControlModel(Control.Parent, false, false);

            return null;
        }

        public List<ItemDataModel> GetItemsSelected() => new SelectedItemDataCollectionModel(this);
        public ItemDataCollectionModel GetItems() => new ItemDataCollectionModel(this);
        public bool IsItemSelected(int index) => Control.Selected[index] != 0;
        public void SetItemSelected(int index, bool selected) => Control.Selected[index] = selected ? 1 : 0;



        public int GetChildrenCount()
        {
            try
            {
                return Control is _Control ? Control.Controls.Count : 0;
            }
            catch { }
            return 0;
        }

        public void LoadChildren(bool getProperties = true)
        {
            try
            {
                if (Control.Controls.Count > 0)
                    Children = Control.Controls.Cast<Control>().Select(c => new AccessControlModel(c, getProperties, true)).ToList();
            }
            catch { }
        }

        public AccessFormModel GetForm()
        {
            return Control.Form != null ? new AccessFormModel(Control.Form, false, false, false) : null;
        }


        public void Blink(string propertyName = "ForeColor")
        {
            const string fallbackPropertyName = "FontUnderline";
            var originColor = this.TryGetDynamicPropertyValue<int>(propertyName);
            var originFontUnderline = this.TryGetDynamicPropertyValue<bool>(fallbackPropertyName);

            for (var i = 0; i < 3; ++i)
            {
                this.SetDynamicPropertyValue(propertyName, (originColor & 0xffffff) ^ 0xffffff);
                this.SetDynamicPropertyValue(fallbackPropertyName, !originFontUnderline);
                Thread.Sleep(500);
                this.SetDynamicPropertyValue(propertyName, originColor);
                this.SetDynamicPropertyValue(fallbackPropertyName, originFontUnderline);
                Thread.Sleep(500);
            }
        }

        public int CompareTo(object obj)
        {
            if (obj == null)
                return 1;

            if (!(obj is AccessControlModel))
                return 1;

            var model = (AccessControlModel)obj;

            if (Control.Application.hWndAccessApp() != model.Control.Application.hWndAccessApp())
                return 1;

            return model.Name == Name ? 0 : 1; // names of controls seem to be unique
        }
    }
}
