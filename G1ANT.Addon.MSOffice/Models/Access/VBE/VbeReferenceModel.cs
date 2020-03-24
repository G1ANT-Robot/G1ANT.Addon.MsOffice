using Microsoft.Vbe.Interop;

namespace G1ANT.Addon.MSOffice.Models.Access.VBE
{
    internal class VbeReferenceModel : INameModel
    {
        public string Name { get; }
        public string FullPath { get; }
        public string Guid { get; }
        public string Description { get; }
        public bool IsBroken { get; }
        public vbext_RefKind Type { get; }
        public bool BuiltIn { get; }
        public string Version { get; }

        public VbeReferenceModel(Reference reference)
        {
            Name = reference.Name;
            FullPath = reference.FullPath;
            Guid = reference.Guid;
            Description = reference.Description;
            IsBroken = reference.IsBroken;
            Type = reference.Type;
            BuiltIn = reference.BuiltIn;
            Version = $"{reference.Major}.{reference.Minor}";
        }

    }
}
