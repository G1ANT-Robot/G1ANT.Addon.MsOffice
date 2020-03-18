using Microsoft.Office.Interop.Access;

namespace G1ANT.Addon.MSOffice.Models.Access.VBE
{
    internal class TempVarModel : INameModel
    {
        public string Name { get; }
        public dynamic Value { get; }

        public TempVarModel(TempVar tempVar)
        {
            Name = tempVar.Name;
            Value = tempVar.Value;
        }


        public override string ToString() => $"{Name}: {Value}";
    }
}
