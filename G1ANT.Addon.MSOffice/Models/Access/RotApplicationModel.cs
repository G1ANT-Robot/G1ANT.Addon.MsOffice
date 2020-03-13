using Microsoft.Office.Interop.Access;

namespace G1ANT.Addon.MSOffice.Models.Access
{
    public class RotApplicationModel
    {
        public string Name { get; set; }
        public Application Application { get; set; }
        public int ApplicationMainHwnd { get; set; }

        public override string ToString()
        {
            return $"{Name} {Application.CurrentProject.Name} (id {ApplicationMainHwnd})";
        }
       
    }
}
