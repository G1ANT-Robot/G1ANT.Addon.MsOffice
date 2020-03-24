using G1ANT.Addon.MSOffice.Models.Access;
using Microsoft.Office.Interop.Access;
using System.Collections.Generic;

namespace G1ANT.Addon.MSOffice.Api.Access
{
    public interface IRunningObjectTableService
    {
        IList<RotApplicationModel> GetApplicationInstances(string applicationProcessName);
        
        /// <summary>
        /// Get list of process ids that do are not registered in ROT
        /// </summary>
        /// <param name="applicationProcessName"></param>
        /// <returns></returns>
        IList<int> GetOrphanedApplicationProcessIds(string applicationProcessName);

        Application GetApplicationInstance(int processId);
        Application GetNewestApplicationInstance();
    }
}