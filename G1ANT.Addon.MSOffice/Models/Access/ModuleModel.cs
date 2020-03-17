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

            Code = module.Lines[0, CountOfLines - 1];
        }

        public void AddFromFile(string path) => Module.AddFromFile(path);
        public void AddFromString(string code) => Module.AddFromString(code);
        public void InsertLines(int from, string code) => Module.InsertLines(from, code);
        public void InsertText(string text) => Module.InsertText(text);
        public void InsertText(int line, string code) => Module.ReplaceLine(line, code);

        public void DeleteLines(int from, int count) => Module.DeleteLines(from, count);

        public override string ToString() => $"{Name} {TypeName}";
    }
}
