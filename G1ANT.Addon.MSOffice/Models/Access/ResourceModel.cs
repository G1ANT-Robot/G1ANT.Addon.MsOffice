using Microsoft.Office.Interop.Access;
using System;

namespace G1ANT.Addon.MSOffice.Models.Access
{
    internal class ResourceModel : INameModel
    {
        public SharedResource Resource { get; }
        public string Name { get; }
        public AcResourceType Type { get; }
        public string TypeName { get; }

        public ResourceModel(SharedResource resource)
        {
            Resource = resource ?? throw new ArgumentNullException(nameof(resource));
            Name = resource.Name;
            Type = resource.Type;
            TypeName = resource.Type.ToString();
        }


        public void Delete() => Resource.Delete();

        public override string ToString() => $"{Name} {TypeName}";

    }
}
