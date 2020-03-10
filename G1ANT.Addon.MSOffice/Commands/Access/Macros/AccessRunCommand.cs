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

namespace G1ANT.Addon.MSOffice.Access
{
    [Command(Name = "access.procedures.run", Tooltip = "This command runs an existing Access procedure (sub) or function")]
    public class AccessRunCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Tooltip = "Name of the macro to be executed", Required = true)]
            public TextStructure ProcedureName { get; set; }
        }

        public AccessRunCommand(AbstractScripter scripter) : base(scripter)
        { }

        public void Execute(Arguments arguments)
        {
            AccessManager.CurrentAccess.Run(arguments.ProcedureName.Value);
        }
    }
}