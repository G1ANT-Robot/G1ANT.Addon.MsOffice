using G1ANT.Addon.MSOffice.Models.Access;
using G1ANT.Addon.MSOffice.Models.Access.Dao;

namespace G1ANT.Addon.MSOffice.Controllers.Access
{
    public interface ITooltipService
    {
        string GetTooltip(AccessControlModel controlModel);
        string GetTooltip(AccessFormModel formModel);
        string GetTooltip(AccessObjectModel formModel);
        string GetTooltip(AccessQueryModel query);
    }
}