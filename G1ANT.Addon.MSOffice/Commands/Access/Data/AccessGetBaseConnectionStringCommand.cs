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
    [Command(Name = "access.getbaseconnectionstring", Tooltip = "This command gets Access base db connection string (CurrentProject.BaseConnectionString)")]
    public class AccessGetBaseConnectionStringCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Tooltip = "Name of a variable where the command's result will be stored")]
            public VariableStructure Result { get; set; } = new VariableStructure("result");
        }

        public AccessGetBaseConnectionStringCommand(AbstractScripter scripter) : base(scripter)
        { }

        public void Execute(Arguments arguments)
        {
            var result = AccessManager.CurrentAccess.GetBaseConnectionString();
            Scripter.Variables.SetVariableValue(arguments.Result.Value, new TextStructure(result));
        }
    }
}