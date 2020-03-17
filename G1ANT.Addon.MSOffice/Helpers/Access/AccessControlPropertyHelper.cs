using G1ANT.Addon.MSOffice.Models.Access;
using System;

namespace G1ANT.Addon.MSOffice.Helpers.Access
{
    static class AccessControlPropertyHelper
    {
        public static T GetDynamicPropertyValue<T>(Microsoft.Office.Interop.Access.Properties properties, string propertyName)
        {
            var value = properties[propertyName].Value;
            return (T)Convert.ChangeType(value, typeof(T));
        }

        public static T TryGetDynamicPropertyValue<T>(this AccessControlModel control, string propertyName)
        {
            try
            {
                return GetDynamicPropertyValue<T>(control.Control.Properties, propertyName);
            }
            catch
            {
                return default(T);
            }
        }

        public static void SetDynamicPropertyValue<T>(this AccessControlModel control, string propertyName, T value)
        {
            control.Control.Properties[propertyName].Value = value;
        }


        public static bool TryGetDynamicPropertyValue<T>(this AccessControlModel control, string propertyName, out T value)
        {
            try
            {
                value = GetDynamicPropertyValue<T>(control.Control.Properties, propertyName);
                return true;
            }
            catch
            {
                value = default(T);
                return false;
            }
        }

        public static T TryGetDynamicPropertyValue<T>(this AccessFormModel form, string propertyName)
        {
            try
            {
                return GetDynamicPropertyValue<T>(form.Form.Properties, propertyName);
            }
            catch
            {
                return default(T);
            }
        }

        public static bool TryGetDynamicPropertyValue<T>(this AccessFormModel form, string propertyName, out T value)
        {
            try
            {
                value = GetDynamicPropertyValue<T>(form.Form.Properties, propertyName);
                return true;
            }
            catch
            {
                value = default(T);
                return false;
            }
        }

        public static void SetDynamicPropertyValue<T>(this AccessFormModel control, string propertyName, T value)
        {
            control.Form.Properties[propertyName].Value = value;
        }
    }
}
