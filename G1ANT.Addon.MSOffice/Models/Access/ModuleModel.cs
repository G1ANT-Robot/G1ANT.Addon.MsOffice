using Microsoft.Office.Interop.Access;
using System;

namespace G1ANT.Addon.MSOffice.Models.Access
{
    internal class ModuleModel : INameModel
    {
        public string Name { get; }
        public AcModuleType Type { get; }
        public Module Module { get; }
        public string TypeName { get; }
        public int CountOfDeclarationLines { get; }
        public int CountOfLines { get; }
        public string Code { get; }

        public ModuleModel(Module module)
        {
            Module = module ?? throw new ArgumentNullException(nameof(module));
            Name = module.Name;
            Type = module.Type;
            TypeName = module.Type.ToString();
            CountOfDeclarationLines = module.CountOfDeclarationLines;
            CountOfLines = module.CountOfLines;

            Code = module.Lines[0, CountOfLines];
        }

        public override string ToString() => $"{Name} {TypeName}";

    }
}
