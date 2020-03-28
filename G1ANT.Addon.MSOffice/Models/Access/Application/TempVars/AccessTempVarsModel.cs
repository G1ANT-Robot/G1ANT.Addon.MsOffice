using Microsoft.Office.Interop.Access;

namespace G1ANT.Addon.MSOffice.Models.Access.Application.TempVars
{
    internal class AccessTempVarsModel : INameModel
    {
        public string Name { get; }
        public string Value { get; }

        public AccessTempVarsModel(TempVar tempVar)
        {
            Name = tempVar.Name;
            try { Value = tempVar.Value?.ToString(); } catch { }
        }

        public override string ToString() => Name;
    }
}
