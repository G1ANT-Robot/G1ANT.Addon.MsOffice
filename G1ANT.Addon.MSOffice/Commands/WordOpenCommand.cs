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
    [Command(Name = "word.open",Tooltip = "This command opens a Word instance with an blank document or a specified file", NeedsDelay = true)]
    public class WordOpenCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Required = true, Tooltip = "Path to a file that has to be opened; if not specified, Word will be opened with a blank document")]
            public TextStructure Path { get; set; } = new TextStructure(string.Empty);

            [Argument(Tooltip = "Name of a variable where the instance's ID will be stored")]
            public VariableStructure Result { get; set; } = new VariableStructure("result");


        }

        public WordOpenCommand(AbstractScripter scripter) : base(scripter)
        {
        }

        public void Execute(Arguments arguments)
        {
            string path = arguments.Path.Value;

            WordWrapper wordWraper = WordManager.AddWord();
            wordWraper.Open(path);
            Scripter.Variables.SetVariableValue(arguments.Result.Value, new Language.IntegerStructure(wordWraper.Id));
        }
    }
}
