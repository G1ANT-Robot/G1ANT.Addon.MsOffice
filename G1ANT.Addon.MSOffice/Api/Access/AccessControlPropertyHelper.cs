using G1ANT.Addon.MSOffice.Models.Access;
using Microsoft.Office.Interop.Access;
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

        public static T TryGetPropertyValue<T>(this _Control control, string propertyName)
        {
            try
            {
                return GetPropertyValue<T>(control.Properties, propertyName);
            }
            catch
            {
                return default(T);
            }
        }

        public static bool TryGetPropertyValue<T>(this _Control control, string propertyName, out T value)
        {
            try
            {
                value = GetPropertyValue<T>(control.Properties, propertyName);
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
