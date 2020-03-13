using Microsoft.Office.Interop.Access;
using System;

namespace G1ANT.Addon.MSOffice.Models.Access
{
    public class RotApplicationModel : IComparable
    {
        public string Name { get; set; }
        public Application Application { get; set; }
        public int ApplicationMainHwnd { get; set; }

        public int CompareTo(object obj)
        {
            return obj?.ToString() == this.ToString() ? 0 : 1;
        }

        public override string ToString()
        {
            return $"{Name} {Application.CurrentProject.Name} (id {ApplicationMainHwnd})";
        }
       
    }
}
