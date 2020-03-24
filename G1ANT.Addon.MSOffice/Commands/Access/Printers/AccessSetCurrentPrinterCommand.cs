/**
*    Copyright(C) G1ANT Ltd, All rights reserved
*    Solution G1ANT.Addon, Project G1ANT.Addon.MSOffice
*    www.g1ant.com
*
*    Licensed under the G1ANT license.
*    See License.txt file in the project root for full license information.
*
*/

using G1ANT.Language;

namespace G1ANT.Addon.MSOffice.Commands.Access.Printers
{
    [Command(Name = "access.setcurrentprinters", Tooltip = "Set current printer of Access. Use `access.getprinters` command to get list of available printers")]
    public class AccessSetCurrentPrinterCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Required = true, Tooltip = "Name of the printer to set")]
            public TextStructure Name { get; set; }
        }

        public AccessSetCurrentPrinterCommand(AbstractScripter scripter) : base(scripter)
        { }

        public void Execute(Arguments arguments)
        {
            AccessManager.CurrentAccess.SetCurrentPrinter(arguments.Name.Value);
        }
    }
}