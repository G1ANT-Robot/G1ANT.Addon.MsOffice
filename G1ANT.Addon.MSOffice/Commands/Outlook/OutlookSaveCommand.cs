using G1ANT.Language;
using Microsoft.Office.Interop.Excel;
using mshtml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace G1ANT.Addon.MSOffice.Commands.Outlook
{
    [Command(Name = "outlook.save", Tooltip = "This command saves an email to specified file")]
    public class OutlookSaveCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Tooltip = "Path to the file where message will be stored")]
            public TextStructure Path { get; set; } 

            [Argument(Tooltip = "Mail structure to be saved")]
            public OutlookMailStructure Mail { get; set; }
        }
        public OutlookSaveCommand(AbstractScripter scripter) : base(scripter)
        {
        }

        public void Execute(Arguments arguments)
        {
            var outlookManager = OutlookManager.CurrentOutlook;
            if (outlookManager != null)
            {
                outlookManager.SaveAs(arguments.Mail?.Value, arguments.Path?.Value);
            }
            else
                throw new NullReferenceException("Current Outlook is not set.");
        }
    }
}
