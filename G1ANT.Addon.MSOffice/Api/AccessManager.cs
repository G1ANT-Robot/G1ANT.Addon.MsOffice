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

namespace G1ANT.Addon.MSOffice
{
    public static class AccessManager
    {
        private static List<AccessWrapper> launchedAccesses = new List<AccessWrapper>();

        public static AccessWrapper CurrentAccess { get; private set; }

        public static AccessWrapper AddAccess()
        {
            var wrapper = new AccessWrapper();
            launchedAccesses.Add(wrapper);
            CurrentAccess = wrapper;
            return wrapper;
        }

        internal static int GetFreeId()
        {
            return launchedAccesses.Select(x => x.Id).DefaultIfEmpty(-1).Max() + 1;
        }

        //internal static bool Switch(int id)
        //{
        //    var wrapper = launchedAccesses.Where(x => x.Id == id).FirstOrDefault();
        //    CurrentAccess = wrapper ?? CurrentAccess;
        //    CurrentAccess.Show();
        //    return wrapper != null;
        //}

        //public static void Remove(AccessWrapper AccessWrapper)
        //{
        //    launchedAccesses.Remove(AccessWrapper);
        //}
    }
}
