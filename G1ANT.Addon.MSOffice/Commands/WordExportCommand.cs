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
    [Command(Name = "word.export", Tooltip = "This command exports a document from the currently active Word instance to a specified file in either .pdf or .xps format", NeedsDelay = true)]

    public class WordExportCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Tooltip = "Path to the exported file; if not specified, the file will be saved in the location of the source file")]
            public TextStructure Path { get; set; } = new TextStructure(string.Empty);

            [Argument(Tooltip = "Type of the exported file: `pdf` or `xps`); if not specified, the type will be defined by the extension of the exported filename")]
            public TextStructure Type { get; set; } = new TextStructure(string.Empty);
        }
        public WordExportCommand(AbstractScripter scripter) : base(scripter)
        {
        }


        public void Execute(Arguments arguments)
        {
            string path = arguments.Path.Value;
            string type = arguments.Type != null ? arguments.Type.Value : null;
            WordWrapper wordWrapper = WordManager.CurrentWord;
            wordWrapper.Export(path, type);
        }
    }
}
