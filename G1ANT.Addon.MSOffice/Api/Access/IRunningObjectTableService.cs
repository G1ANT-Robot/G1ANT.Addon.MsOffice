using G1ANT.Addon.MSOffice.Models.Access;
using System.Collections.Generic;

namespace G1ANT.Addon.MSOffice.Api.Access
{
    public interface IRunningObjectTableService
    {
        IList<RotApplicationModel> GetApplicationInstances(string applicationProcessName);
    }
}