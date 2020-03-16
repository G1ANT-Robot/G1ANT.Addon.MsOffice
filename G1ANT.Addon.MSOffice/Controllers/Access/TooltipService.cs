using G1ANT.Addon.MSOffice.Models.Access;
using G1ANT.Addon.MSOffice.Models.Access.Dao;
using System.Linq;
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
            result.AppendLine($"Type: {formModel.TypeName}");
            result.AppendLine($"IsLoaded: {formModel.IsLoaded}");
            result.AppendLine($"IsWeb: {formModel.IsWeb}");
            result.AppendLine($"Attributes: {formModel.Attributes}");
            result.AppendLine($"DateCreated: {formModel.DateCreated}");
            result.AppendLine($"DateModified: {formModel.DateModified}");

            return result.ToString();
        }


        public string GetTooltip(AccessQueryModel query)
        {
            var result = new StringBuilder();

            result.AppendLine($"Name: {query.Name}");
            result.AppendLine($"Type: {query.Type}");
            result.AppendLine($"DateCreated: {query.DateCreated}");
            result.AppendLine($"DateModified: {query.LastUpdated}");
            result.AppendLine($"Connect: {query.Connect}");
            result.AppendLine($"Fields: {string.Join(", ", query.Fields.Select(f => f.Name))}");
            result.AppendLine($"Parameters: {string.Join(", ", query.Parameters.Select(p => p.Name))}");
            result.AppendLine($"Prepare: {query.Prepare}");
            result.AppendLine($"Properties: {string.Join(", ", query.Properties.Select(p => p.Name))}");
            result.AppendLine($"Properties: {query.Query}");
            result.AppendLine($"RecordsAffected: {query.RecordsAffected}");
            result.AppendLine($"ReturnsRecords: {query.ReturnsRecords}");
            result.AppendLine($"SQL: {query.SQL}");
            result.AppendLine($"StillExecuting: {query.StillExecuting}");
            result.AppendLine($"Type: {query.Type}");
            result.AppendLine($"Updatable: {query.Updatable}");

            return result.ToString();
        }
    }
}
