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
    [Command(Name = "access.exportobjecttoxml", Tooltip = "This command allows to export XML data, schemas, and presentation information from Microsoft SQL Server 2000 Desktop Engine (MSDE 2000), Microsoft SQL Server 6.5 or later, or the Microsoft Access database engine")]
    public class AccessExportObjectToXmlCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Tooltip = "Type of object to export. Possible values: Form, Function, Query, Report ServerView, StoredProcedure, Table", Required = true)]
            public TextStructure TypeOfObjectToExport { get; set; }

            [Argument(Tooltip = "Name of the object to export", Required = true)]
            public TextStructure ObjectName { get; set; }

            [Argument(Tooltip = "Path to a file for exported data", Required = true)]
            public TextStructure Path { get; set; }

            [Argument(Tooltip = "The path for the exported schema information. If this argument is omitted, schema information is not exported to a separate XML file")]
            public TextStructure PathForSchema { get; set; }

            [Argument(Tooltip = "The path for the exported presentation information. If this argument is omitted, presentation information is not exported")]
            public TextStructure PathForPresentation { get; set; }

            [Argument(Tooltip = "The directory for exported images. If this argument is omitted, images are not exported")]
            public TextStructure DirectoryForImages { get; set; }

            [Argument(Tooltip = "Set encoding for the files: true for UTF16, false for UTF8. False is default.")]
            public BooleanStructure UseUTF16ForEncoding { get; set; } = new BooleanStructure(false);

            [Argument(Tooltip = "Other behaviors associated with exporting to XML. Possible values: EmbedSchema, ExcludePrimaryKeyAndIndexes, RunFromServer, LiveReportSource, PersistReportML, ExportAllTableAndFieldProperties")]
            public TextStructure OtherFlags { get; set; }

            [Argument(Tooltip = "A valid WHERE part of SQL w/o `where` keyword. Specifies a subset of records to be exported.")]
            public TextStructure WhereCondition { get; set; }

            [Argument(Tooltip = "Specifies additional tables to export. This argument is ignored if the OtherFlags argument is set to acLiveReportSource")]
            public TextStructure AdditionalData { get; set; }
        }

        public AccessExportObjectToXmlCommand(AbstractScripter scripter) : base(scripter)
        { }

        public void Execute(Arguments arguments)
        {
            AccessManager.CurrentAccess.ExportXML(
                arguments.TypeOfObjectToExport.Value,
                arguments.ObjectName.Value,
                arguments.Path.Value,
                arguments.PathForSchema?.Value,
                arguments.PathForPresentation?.Value,
                arguments.DirectoryForImages?.Value,
                arguments.UseUTF16ForEncoding.Value,
                arguments.OtherFlags?.Value,
                arguments.WhereCondition?.Value,
                arguments.AdditionalData?.Value
            );
        }
    }
}