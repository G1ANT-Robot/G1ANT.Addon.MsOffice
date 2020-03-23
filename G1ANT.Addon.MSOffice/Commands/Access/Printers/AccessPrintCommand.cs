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

namespace G1ANT.Addon.MSOffice.Commands.Access.Printers
{
    [Command(Name = "access.print", Tooltip = "This command prints the active object in the open Access database. You can print datasheets, reports, forms, data access pages and modules")]
    public class AccessPrintCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Tooltip = "Range to print. Possible values are `Pages` (use to print specific page range), `PrintAll` (the default), `Selection` (selected part of the object)")]
            public TextStructure PrintRange { get; set; } = new TextStructure("PrintAll");

            [Argument(Tooltip = "The first page to print, a valid page number in the active form or datasheet. This argument is required if you specify `Pages` for the `PrintRange` argument")]
            public IntegerStructure PageFrom { get; set; } = new IntegerStructure(0);

            [Argument(Tooltip = "The last page to print, a valid page number in the active form or datasheet. This argument is required if you specify `Pages` for the `PrintRange` argument.")]
            public IntegerStructure PageTo { get; set; } = new IntegerStructure(0);

            [Argument(Tooltip = "Specifies the print quality. Possible values are `High`, `Medium`, `Low`, `Draft`. The defaul value is `High`")]
            public TextStructure PrintQuality { get; set; } = new TextStructure("PrintQuality");

            [Argument(Tooltip = "The number of copies to print. 1 by default.")]
            public IntegerStructure Copies { get; set; } = new IntegerStructure(1);

            [Argument(Tooltip = "Set to true to collate copies and to false to print without collating. The default is true")]
            public BooleanStructure CollateCopies { get; set; } = new BooleanStructure(true);
        }

        public AccessPrintCommand(AbstractScripter scripter) : base(scripter)
        { }

        public void Execute(Arguments arguments)
        {
            AccessManager.CurrentAccess.PrintActiveObject();
        }
    }
}