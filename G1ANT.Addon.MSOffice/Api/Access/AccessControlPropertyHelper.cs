using G1ANT.Addon.MSOffice.Models.Access;
using System;

namespace G1ANT.Addon.MSOffice.Api.Access
{
    static class AccessControlPropertyHelper
    {
        public static T GetPropertyValue<T>(Microsoft.Office.Interop.Access.Properties properties, string propertyName)
        {
            var value = properties[propertyName].Value;
            return (T)Convert.ChangeType(value, typeof(T));
        }

        public static T TryGetPropertyValue<T>(this AccessControlModel control, string propertyName)
        {
            try
            {
                return GetPropertyValue<T>(control.Control.Properties, propertyName);
            }
            catch
            {
                return default(T);
            }
        }

        public static void SetPropertyValue<T>(this AccessControlModel control, string propertyName, T value)
        {
            control.Control.Properties[propertyName].Value = value;
        }


        public static bool TryGetPropertyValue<T>(this AccessControlModel control, string propertyName, out T value)
        {
            try
            {
                value = GetPropertyValue<T>(control.Control.Properties, propertyName);
                return true;
            }
            catch
            {
                value = default(T);
                return false;
            }
        }

        public static T TryGetPropertyValue<T>(this AccessFormModel form, string propertyName)
        {
            try
            {
                return GetPropertyValue<T>(form.Form.Properties, propertyName);
            }
            catch
            {
                return default(T);
            }
        }

        public static bool TryGetPropertyValue<T>(this AccessFormModel form, string propertyName, out T value)
        {
            try
            {
                value = GetPropertyValue<T>(form.Form.Properties, propertyName);
                return true;
            }
            catch
            {
                value = default(T);
                return false;
            }
        }

    }
}
