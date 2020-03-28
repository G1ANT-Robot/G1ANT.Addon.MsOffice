using System;

namespace G1ANT.Addon.MSOffice.Models.Access
{
    public class RotApplicationModel : IComparable, INameModel
    {
        public string Name { get; set; }
        public Microsoft.Office.Interop.Access.Application Application { get; set; }
        public int ApplicationMainHwnd { get; set; }
        public int ProcessId { get; set; }

        public int CompareTo(object obj)
        {
            return obj?.ToString() == ToString() ? 0 : 1;
        }

        public override string ToString()
        {
            return $"{Name} {Application?.CurrentProject.FullName} (id {ApplicationMainHwnd})";
        }

    }
}
