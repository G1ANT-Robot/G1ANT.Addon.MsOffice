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

namespace G1ANT.Addon.MSOffice.Commands.Access.Macros
{
    [Command(Name = "access.runprocedure", Tooltip = "This command runs an existing Access procedure (sub) or function")]
    public class AccessRunProcedureCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Tooltip = "Name of the procedure to be executed", Required = true)]
            public TextStructure ProcedureName { get; set; }
        }

        public AccessRunProcedureCommand(AbstractScripter scripter) : base(scripter)
        { }

        public void Execute(Arguments arguments)
        {
            AccessManager.CurrentAccess.RunProcedure(arguments.ProcedureName.Value);
        }
    }
}