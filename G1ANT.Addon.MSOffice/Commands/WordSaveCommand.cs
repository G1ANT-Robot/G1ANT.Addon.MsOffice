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
    [Command(Name = "word.save",Tooltip = "This command saves the currently active Word document to a specified file", NeedsDelay = true)]

    public class WordSaveCommand : Command
	{
        public class Arguments : CommandArguments
        {
            [Argument(Required = true, Tooltip = "Specifies path, where the current Word document will be saved. If a filename is not specified and the document has never been saved, the default name is used (for example, `Doc1.docx`)")]
            public TextStructure Path { get; set; } = new TextStructure(string.Empty);

        }

        public WordSaveCommand(AbstractScripter scripter) : base(scripter)
        {
        }

        public void Execute(Arguments arguments)
        {
            WordWrapper wordWrapper = WordManager.CurrentWord;
            wordWrapper.Save(arguments.Path.Value);
        }
    }
}
