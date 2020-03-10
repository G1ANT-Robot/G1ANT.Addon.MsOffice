using Microsoft.Office.Interop.Access;
using System;

namespace G1ANT.Addon.MSOffice.Api.Access
{
    static class AccessControlPropertyHelper
    {
        public static T GetPropertyValue<T>(Control control, string propertyName)
        {
            var value = control.Properties[propertyName].Value;
            return (T)Convert.ChangeType(value, typeof(T));
        }

        public static T TryGetPropertyValue<T>(Control control, string propertyName)
        {
            try
            {
                return GetPropertyValue<T>(control, propertyName);
            }
            catch
            {
                return default(T);
            }
        }

        public static bool TryGetPropertyValue<T>(Control control, string propertyName, out T value)
        {
            try
            {
                value = GetPropertyValue<T>(control, propertyName);
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
