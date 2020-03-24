using Microsoft.Vbe.Interop;

namespace G1ANT.Addon.MSOffice.Models.Access.VBE
{
    internal class VbeProjectModel : INameModel
    {
        public string FileName { get; }
        public string BuildFileName { get; }
        //public VbeProjectCollectionModel Projects { get; }
        public string Name { get; }
        public string Description { get; }
        public vbext_VBAMode Mode { get; }
        public vbext_ProjectType Type { get; }
        public bool Saved { get; }
        public VbeReferenceCollectionModel References { get; }

        public VbeProjectModel(VBProject project)
        {
            Name = project.Name;
            Description = project.Description;
            FileName = project.FileName;
            BuildFileName = project.BuildFileName;
            //if (project.Collection.Count > 0)
            //    Projects = new VbeProjectCollectionModel(project.Collection);
            Mode = project.Mode;
            Type = project.Type;
            Saved = project.Saved;
            if (project.References.Count > 0)
                References = new VbeReferenceCollectionModel(project.References);
        }

        public override string ToString() => $"{Name}, {Description}, {FileName}, mode: {Mode}, type: {Type}, saved: {Saved}";
    }
}
