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

namespace G1ANT.Addon.MSOffice.Commands.Access.Projects
{
    [Command(Name = "access.importobjectfromxml", Tooltip = "This command allows to import XML data and/or schema information into Microsoft SQL Server 2000 Desktop Engine (MSDE 2000), Microsoft SQL Server 7.0 (or later) or the Microsoft Access database engine")]
    public class AccessImportObjectFromXmlCommand : Command
    {
        public class Arguments : CommandArguments
        {
    [Argument(Tooltip = "Path to a file for imported data", Required = true)]
            public TextStructure Path { get; set; }

            [Argument(Tooltip = "Specifies the option to use when importing XML files. Possible values are StructureOnly, StructureAndData, AppendData. The default value is StructureAndData.")]
            public TextStructure Options { get; set; } = new TextStructure("StructureAndData");
        }

        public AccessImportObjectFromXmlCommand(AbstractScripter scripter) : base(scripter)
        { }

        public void Execute(Arguments arguments)
        {
            AccessManager.CurrentAccess.ImportXml(arguments.Path.Value, arguments.Options.Value);
        }
    }
}