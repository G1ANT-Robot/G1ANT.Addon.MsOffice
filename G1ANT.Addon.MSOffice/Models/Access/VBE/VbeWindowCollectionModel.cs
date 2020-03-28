using System.Collections.Generic;
using System.Linq;

namespace G1ANT.Addon.MSOffice.Models.Access.VBE
{
    internal class VbeWindowCollectionModel : List<VbeWindowModel>
    {
        public VbeWindowCollectionModel(Microsoft.Vbe.Interop.Windows windows)
        {
            AddRange(windows.Cast<Microsoft.Vbe.Interop.Window>().Select(w => new VbeWindowModel(w)));
        }

        public override string ToString() => "Windows";
    }
}
