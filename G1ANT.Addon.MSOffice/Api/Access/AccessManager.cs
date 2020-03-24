/**
*    Copyright(C) G1ANT Ltd, All rights reserved
*    Solution G1ANT.Addon, Project G1ANT.Addon.MSOffice
*    www.g1ant.com
*
*    Licensed under the G1ANT license.
*    See License.txt file in the project root for full license information.
*
*/

using G1ANT.Addon.MSOffice.Access;
using G1ANT.Addon.MSOffice.Api.Access;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace G1ANT.Addon.MSOffice
{
    internal static class AccessManager
    {
        private static List<AccessWrapper> launchedAccesses = new List<AccessWrapper>();

        internal static AccessWrapper CurrentAccess { get; private set; }

        internal static AccessWrapper AddAccess()
        {
            //if (GetOfficeAppPath("Access.Application", "msaccess.exe") == null)
            //{
            //    throw new Exception("Can't determine path to msaccess.exe");
            //}

            var wrapper = new AccessWrapper(new AccessFormControlsTreeWalker(), new RunningObjectTableService());
            launchedAccesses.Add(wrapper);
            CurrentAccess = wrapper;
            return wrapper;
        }

        internal static void KillOrphanedAccessProcesses()
        {
            var processIds = new RunningObjectTableService().GetOrphanedApplicationProcessIds("msaccess");
            foreach (var processId in processIds)
            {
                var process = Process.GetProcessById(processId);
                process.Kill();
            }
        }

        internal static int GetFreeId()
        {
            return launchedAccesses.Select(x => x.Id).DefaultIfEmpty(-1).Max() + 1;
        }

        internal static bool Switch(int id)
        {
            var wrapper = launchedAccesses.Where(x => x.Id == id).FirstOrDefault();
            CurrentAccess = wrapper ?? CurrentAccess;
            CurrentAccess.Show();
            return wrapper != null;
        }

        internal static void Remove(AccessWrapper accessWrapper)
        {
            launchedAccesses.Remove(accessWrapper);
            CurrentAccess = launchedAccesses.FirstOrDefault();
        }

        //private static string GetOfficeAppPath(string progId, string executableName)
        //{
        //    try
        //    {
        //        var oReg = Registry.LocalMachine;

        //        var oKey = oReg.OpenSubKey($@"Software\Classes\{progId}\CLSID");
        //        var clsid = oKey.GetValue("").ToString();
        //        oKey.Close();

        //        oKey = oReg.OpenSubKey($@"Software\Classes\CLSID\{clsid}\LocalServer32");
        //        var sPath = oKey.GetValue("").ToString();
        //        oKey.Close();

        //        var iPos = sPath.IndexOf(executableName, StringComparison.CurrentCultureIgnoreCase);
        //        sPath = sPath.Substring(0, iPos + executableName.Length);
        //        return sPath.Trim();
        //    }
        //    catch
        //    {
        //        return null;
        //    }
        //}
    }
}
