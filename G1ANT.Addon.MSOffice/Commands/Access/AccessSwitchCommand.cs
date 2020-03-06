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
    [Command(Name = "access.switch",Tooltip = "This command switches between Access instances open with `access.open` command", NeedsDelay = true)]
    public class AccessSwitchCommand : Command
	{
        public class Arguments : CommandArguments
        {
            [Argument(Required = true, Tooltip = "Id of an Access window that was returned by the `access.open` command")]
            public IntegerStructure Id { get; set; }

            [Argument(Tooltip = "Name of a variable where the command's result will be stored (true or false)")]
            public VariableStructure Result { get; set; } = new VariableStructure("result");          
        }

        public AccessSwitchCommand(AbstractScripter scripter) : base(scripter)
        {
        }

        public void Execute(Arguments arguments)
        {
            var result = AccessManager.Switch(arguments.Id.Value);

            Scripter.Variables.SetVariableValue(arguments.Result.Value, new BooleanStructure(result));
        }
    }
}
