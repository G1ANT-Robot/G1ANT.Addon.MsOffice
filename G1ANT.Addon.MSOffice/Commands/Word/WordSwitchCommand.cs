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

namespace G1ANT.Addon.MSOffice
{
    [Command(Name = "word.switch",Tooltip = "This command switches between open Word instances", NeedsDelay = true)]

    public class WordSwitchCommand : Command
	{
        public class Arguments : CommandArguments
        {
            [Argument(Required = true, Tooltip = "ID of a Word window that was specified while using the `word.open` command")]
            public IntegerStructure Id { get; set; }

            [Argument(Tooltip = "Name of a variable where the command's result will be stored (true or false)")]
            public VariableStructure Result { get; set; } = new VariableStructure("result");          


        }
        public WordSwitchCommand(AbstractScripter scripter) : base(scripter)
        {
        }
        public void Execute(Arguments arguments)
        {
            int id = arguments.Id.Value;
            if (WordManager.Switch(id))
            {
                Scripter.Variables.SetVariableValue(arguments.Result.Value, new Language.BooleanStructure(true));
            }
        }
    }
}
