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

namespace G1ANT.Addon.MSOffice.Commands.Access.Data
{
    [Command(Name = "access.closefunction", Tooltip = "This command closes and optionally saves an Access Database Diagram")]
    public class AccessCloseFunctionCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Tooltip = "Name of the function to close", Required = true)]
            public TextStructure Name { get; set; }

            [Argument(Tooltip = "Set to true to save changes, false to discard them, don't set any value to prompt user", Required = true)]
            public BooleanStructure SaveChanges { get; set; }
        }

        public AccessCloseFunctionCommand(AbstractScripter scripter) : base(scripter)
        { }

        public void Execute(Arguments arguments)
        {
            AccessManager.CurrentAccess.Close(Microsoft.Office.Interop.Access.AcObjectType.acFunction, arguments.Name.Value, arguments.SaveChanges?.Value);
        }
    }
}