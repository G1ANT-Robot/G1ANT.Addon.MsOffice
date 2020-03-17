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
using System;

namespace G1ANT.Addon.MSOffice.Commands.Access
{
    [Command(Name = "access.join", Tooltip = "This command joins to an existing Access instance")]
    public class AccessJoinCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Required = true, Tooltip = "Id of Access process to join to")]
            public IntegerStructure ProcessId { get; set; }

            [Argument(Tooltip = "Name of a variable where a currently opened Access process number is stored. It can be used in the `access.switch` command")]
            public VariableStructure Result { get; set; } = new VariableStructure("result");
        }

        public AccessJoinCommand(AbstractScripter scripter) : base(scripter)
        { }

        public void Execute(Arguments arguments)
        {
            try
            {
                var access = AccessManager.AddAccess();
                access.JoinToExistingInstance(arguments.ProcessId.Value);

                Scripter.Variables.SetVariableValue(arguments.Result.Value, new IntegerStructure(access.Id));
            }
            catch (Exception ex)
            {
                //if (ex.GetType() == typeof(COMException) && ex.Message.Contains("80040154"))
                //    throw new Exception("Could not find Microsoft Office on computer. Please make sure it is installed and try again.");
                throw new ApplicationException($"Problem occured while joining to Access instance. ProcessId: '{arguments.ProcessId.Value}'", ex);
            }
        }
    }
}