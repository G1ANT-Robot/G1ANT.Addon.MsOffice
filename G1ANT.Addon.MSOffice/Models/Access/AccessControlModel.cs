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
using System.Text;
using System.Threading;

namespace G1ANT.Addon.MSOffice.Models.Access
{
    internal class AccessControlModel : IComparable, INameModel, IDetailedNameModel
    {
        public string Name { get; }
        public string Type { get; }

        public int Top { get; }
        public int Left { get; }
        public int Width { get; }
        public int Height { get; }

        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        public string Caption { get; private set; }

        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        public dynamic Value { get; private set; }

        [JsonIgnore]
        public _Control Control { get; }

        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        public ICollection<AccessControlModel> Children;

        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        public AccessDynamicPropertyCollectionModel Properties;


        public AccessControlModel(_Control control, bool getProperties = true, bool getChildrenRecursively = true)
        {
            Control = control ?? throw new ArgumentNullException(nameof(control));
            Name = control.Name;
            Type = ((AcControlType)this.TryGetDynamicPropertyValue<int>("ControlType")).ToString();
            Caption = TryGetPropertyValue<string>("Caption");
            Value = TryGetPropertyValue<string>("Value");

            Left = TryGetPropertyValue<int>("Left");
            Top = TryGetPropertyValue<int>("Top");
            Width = TryGetPropertyValue<int>("Width");
            Height = TryGetPropertyValue<int>("Height");

            if (getProperties && control.Properties.Count > 0)
                LoadProperties();

            if (getChildrenRecursively)
                LoadChildren(getProperties);
        }


        public void LoadProperties()
        {
            Properties = new AccessDynamicPropertyCollectionModel(Control.Properties);
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

        internal AccessDynamicPropertyCollectionModel GetDynamicProperties() => new AccessDynamicPropertyCollectionModel(Control.Properties);

        internal T TryGetPropertyValue<T>(string name)
        {
            try
            {
                return (T)Convert.ChangeType(GetPropertyValue(name), typeof(T));
            }
            catch
            {
                return default(T);
            }
        }

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
        public bool IsItemSelected(int index) => Control.Selected[ToItemIndex(index)] != 0;
        public void SetItemSelected(int index, bool selected) => Control.Selected[ToItemIndex(index)] = selected ? 1 : 0;


        private int ToItemIndex(int i)
        {
            var isHeaderVisible = TryGetPropertyValue<bool>("ColumnHeads");
            return isHeaderVisible ? i + 1 : i;
        }

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
            if (!this.TryGetDynamicPropertyValue(propertyName, out int originColor))
            {
                Shake();
                return;
            }
            this.TryGetDynamicPropertyValue(fallbackPropertyName, out bool originFontUnderline);

            for (var i = 0; i < 6; ++i)
            {
                this.TrySetDynamicPropertyValue(propertyName, (originColor & 0xffffff) ^ 0xffffff);
                this.TrySetDynamicPropertyValue(fallbackPropertyName, !originFontUnderline);
                Thread.Sleep(200);
                this.TrySetDynamicPropertyValue(propertyName, originColor);
                this.TrySetDynamicPropertyValue(fallbackPropertyName, originFontUnderline);
                Thread.Sleep(200);
            }
        }

        private void Shake()
        {
            var random = new Random();
            for (var i = 0; i < 50; ++i)
            {
                Control.Move(Top + random.Next(-200, 200), Left + random.Next(-200, 200));
                Thread.Sleep(50);
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

        public override string ToString() => $"{Caption} {Name} {Type} {Value}";

        public string ToDetailedString()
        {
            var result = new StringBuilder();

            result.AppendLine($"Type: {Type}\r\n");
            result.AppendLine($"Name: {Name}");
            result.AppendLine($"Caption: {Caption}");
            if (Value != null)
                result.AppendLine($"Value: {Value}");

            return result.ToString();
        }
    }
}
