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
    [Command(Name = "word.runmacro",Tooltip = "This command runs a macro in the currently active Word instance", NeedsDelay = true)]
    public class WordRunMacroCommand : Command
	{
        public class Arguments : CommandArguments
        {
            [Argument(Required = true, Tooltip = "Name of a macro")]
            public TextStructure Name { get; set; } = new TextStructure(string.Empty);

            [Argument(Tooltip = "Arguments for the specified macro")]
            public TextStructure Args { get; set; } 


        }

        public WordRunMacroCommand(AbstractScripter scripter) : base(scripter)
        {
        }

        public void Execute(Arguments arguments)
        {
            string macroName = arguments.Name.Value;

            string args = arguments.Args != null ? arguments.Args.Value : null;
            var wrapper = WordManager.CurrentWord;
            wrapper.RunMacro(macroName, args);
        }
    }
}
