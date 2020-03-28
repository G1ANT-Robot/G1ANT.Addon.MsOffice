using Microsoft.Office.Interop.Access;
using System;

namespace G1ANT.Addon.MSOffice.Models.Access
{
    public class AccessCurrentProjectModel : INameModel
    {
        public string BaseConnectionString { get; }
        public AcFileFormat FileFormat { get; }
        public string FullName { get; }
        //public ImportExportSpecifications ImportExportSpecifications { get; }
        public bool IsConnected { get; }
        public bool IsSQLBackend { get; }
        public bool IsTrusted { get; }
        public bool IsWeb { get; }
        public string Name { get; }
        public string Path { get; }
        public AcProjectType ProjectType { get; }
        public Lazy<AccessObjectPropertyCollectionModel> Properties { get; }
        public string WebSite { get; }

        public AccessCurrentProjectModel(_CurrentProject project)
        {
            BaseConnectionString = project.BaseConnectionString;
            FileFormat = project.FileFormat;
            FullName = project.FullName;
            //ImportExportSpecifications = project.ImportExportSpecifications;
            IsConnected = project.IsConnected;
            IsSQLBackend = project.IsSQLBackend;
            IsTrusted = project.IsTrusted;
            IsWeb = project.IsWeb;
            Name = project.Name;
            Path = project.Path;
            ProjectType = project.ProjectType;
            Properties = new Lazy<AccessObjectPropertyCollectionModel>(() => new AccessObjectPropertyCollectionModel(project.Properties));
            WebSite = project.WebSite;
        }

        public override string ToString() => Name;
    }
}
