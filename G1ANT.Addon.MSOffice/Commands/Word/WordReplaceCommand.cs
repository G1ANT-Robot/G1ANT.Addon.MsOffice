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
    [Command(Name = "word.replace",Tooltip = "This command replaces specified text in a document", NeedsDelay = true)]
    public class WordReplaceCommand : Command
	{
        public class Arguments : CommandArguments
        {
            [Argument(Required = true, Tooltip = "Text to be found in a document")]
            public TextStructure From { get; set; } = new TextStructure(string.Empty);

            [Argument(Required = true, Tooltip = "Text to be replaced in a document")]
            public TextStructure To { get; set; } = new TextStructure(string.Empty);

            [Argument(Tooltip = "If set to `true`, then the search is case sensitive")]
            public BooleanStructure MatchCase { get; set; } = new BooleanStructure(false);

            [Argument(Tooltip = "If set to `true`, only whole words are replaced")]
            public BooleanStructure WholeWords { get; set; } = new BooleanStructure(false);


        }
        public WordReplaceCommand(AbstractScripter scripter) : base(scripter)
        {
        }

        public void Execute(Arguments arguments)
        {
            var wrapper = WordManager.CurrentWord;
            wrapper.ReplaceWord(arguments.From.Value, arguments.To.Value, arguments.MatchCase.Value, arguments.WholeWords.Value);
        }
    }
}
