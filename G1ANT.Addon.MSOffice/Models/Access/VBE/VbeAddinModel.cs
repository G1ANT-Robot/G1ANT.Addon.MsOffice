using Microsoft.Vbe.Interop;

namespace G1ANT.Addon.MSOffice.Models.Access.VBE
{
    internal class VbeAddinModel : INameModel, IDetailedNameModel
    {
        public string Name => Description;
        public string Description { get; }
        public dynamic Object { get; }
        public string Guid { get; }
        public string ProgId { get; }

        public VbeAddinModel(AddIn addIn)
        {
            Description = addIn.Description;
            try
            {
                Object = addIn.Object?.ToString();
                Guid = addIn.Guid;
                ProgId = addIn.ProgId;
            }
            catch { }
        }

        public string ToDetailedString() => $"{Description}, {Object}, {Guid}";
        public override string ToString() => Name;
    }
}
