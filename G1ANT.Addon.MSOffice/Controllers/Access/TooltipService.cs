using G1ANT.Addon.MSOffice.Models.Access;
using System.Text;

namespace G1ANT.Addon.MSOffice.Controllers.Access
{
    public class TooltipService : ITooltipService
    {
        public string GetTooltip(AccessControlModel controlModel)
        {
            var result = new StringBuilder();

            result.AppendLine($"Type: {controlModel.Type}\r\n");
            result.AppendLine($"Name: {controlModel.Name}");
            result.AppendLine($"Caption: {controlModel.Caption}");
            if (controlModel.Value != null)
                result.AppendLine($"Value: {controlModel.Value}");

            return result.ToString();
        }

        public string GetTooltip(AccessFormModel formModel)
        {
            var result = new StringBuilder();

            result.AppendLine($"Name: {formModel.Name}");
            result.AppendLine($"Caption: {formModel.Caption}");
            if (formModel.FormName != formModel.Name)
                result.AppendLine($"FormName: {formModel.FormName}");
            result.AppendLine($"Height: {formModel.Height}");
            result.AppendLine($"Width: {formModel.Width}");
            result.AppendLine($"X: {formModel.X}");
            result.AppendLine($"Y: {formModel.Y}");

            return result.ToString();
        }

        public string GetTooltip(AccessObjectModel formModel)
        {
            var result = new StringBuilder();

            result.AppendLine($"Name: {formModel.Name}");
            result.AppendLine($"FullName: {formModel.FullName}");
            result.AppendLine($"Type: {formModel.Type}");
            result.AppendLine($"IsLoaded: {formModel.IsLoaded}");
            result.AppendLine($"IsWeb: {formModel.IsWeb}");
            result.AppendLine($"Attributes: {formModel.Attributes}");
            result.AppendLine($"DateCreated: {formModel.DateCreated}");
            result.AppendLine($"DateModified: {formModel.DateModified}");

            return result.ToString();
        }
    }
}
