using Microsoft.Vbe.Interop;

namespace G1ANT.Addon.MSOffice.Models.Access.VBE
{
    internal class VbeAddinModel
    {
        public string Description { get; }
        public dynamic Object { get; }
        public string Guid { get; }
        public string ProgId { get; }

        public VbeAddinModel(AddIn addIn)
        {
            Description = addIn.Description;
            Object = addIn.Object?.ToString();
            Guid = addIn.Guid;
            ProgId = addIn.ProgId;
        }

        public override string ToString() => $"{Description}, {Object}, {Guid}";
    }
}
