using G1ANT.Addon.MSOffice.Models.Access;
using Microsoft.Office.Interop.Access;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace G1ANT.Addon.MSOffice.Api.Access
{
    public class RunningObjectTableService : IRunningObjectTableService
    {
        private const uint OBJID_NATIVEOM = 0xFFFFFFF0;
        private static readonly Guid IID_IDispatch = new Guid("{00020400-0000-0000-C000-000000000046}");

        public IList<RotApplicationModel> GetApplicationInstances(string applicationProcessName)
        {
            applicationProcessName = Path.GetFileNameWithoutExtension(applicationProcessName);

            var applicationProcesses = Process.GetProcessesByName(applicationProcessName).Concat(Process.GetProcessesByName(applicationProcessName + ".EXE"));

            var result = new List<RotApplicationModel>();
            var exceptions = new List<Exception>();
            foreach (var accessProcess in applicationProcesses)
            {
                try
                {
                    result.Add(GetApplicationFromProcess(accessProcess));
                }
                catch (Exception ex)
                {
                    result.Add(new RotApplicationModel() { Name = $"Failed to load the application for process {accessProcess.Id} hwnd {accessProcess.MainWindowHandle}" });
                    exceptions.Add(ex);
                }
            }

            //if (exceptions.Any())
            //    throw new AggregateException("At least one of GetApplicationFromProcess (AccessibleObjectFromWindow) calls failed", exceptions);

            return result;
        }

        private RotApplicationModel GetApplicationFromProcess(Process process)
        {
            var mainHandle = process.MainWindowHandle;
            if (mainHandle.ToInt32() > 0)
            {
                var IID_IDispatch = RunningObjectTableService.IID_IDispatch;
                int res = OleAccWrapper.AccessibleObjectFromWindow(mainHandle, OBJID_NATIVEOM, ref IID_IDispatch, out Application app);
                if (res >= 0)
                {
                    return new RotApplicationModel()
                    {
                        Name = app.Name,
                        Application = app,
                        ApplicationMainHwnd = app.hWndAccessApp()
                    };
                }

                throw new Exception($"AccessibleObjectFromWindow returned false for hwnd {mainHandle}");
            }

            throw new Exception($"Process {process.ProcessName} id {process.Id} does not have main window");
        }


        /// <summary>
        /// Get list of process ids that do are not registered in ROT
        /// </summary>
        /// <param name="applicationProcessName"></param>
        /// <returns></returns>
        public IList<int> GetOrphanedApplicationProcessIds(string applicationProcessName)
        {
            applicationProcessName = Path.GetFileNameWithoutExtension(applicationProcessName);
            var applicationProcesses = Process.GetProcessesByName(applicationProcessName).Concat(Process.GetProcessesByName(applicationProcessName + ".EXE"));

            return applicationProcesses
                .Where(process => IsOrphaned(process))
                .Select(process => process.Id)
                .ToList();

        }

        private bool IsOrphaned(Process process)
        {
            try
            {
                var mainHandle = process.MainWindowHandle;
                if (mainHandle.ToInt32() > 0)
                {
                    var IID_IDispatch = RunningObjectTableService.IID_IDispatch;
                    int res = OleAccWrapper.AccessibleObjectFromWindow(mainHandle, OBJID_NATIVEOM, ref IID_IDispatch, out Application app);
                    return res < 0;
                }
            }
            catch { }

            return true;
        }


    }
}
