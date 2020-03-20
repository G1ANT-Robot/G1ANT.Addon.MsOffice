using Microsoft.Office.Interop.Access;
using System.Collections.Generic;
using System.Linq;

namespace G1ANT.Addon.MSOffice.Models.Access.Printers
{
    internal class AccessPrinterCollectionModel : List<AccessPrinterModel>
    {
        public AccessPrinterCollectionModel(Microsoft.Office.Interop.Access.Printers printers)
        {
            AddRange(printers.Cast<Printer>().Select(p => new AccessPrinterModel(p)));
        }
    }
}
