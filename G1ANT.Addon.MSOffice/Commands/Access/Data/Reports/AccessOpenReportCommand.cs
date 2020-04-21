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

namespace G1ANT.Addon.MSOffice.Commands.Access.Data.Reports
{
    [Command(Name = "access.openreport", Tooltip = "This command opens an Access Report")]
    public class AccessOpenReportCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Tooltip = "Name of the Report to open", Required = true)]
            public TextStructure Name { get; set; }

            [Argument(Tooltip = "Open mode. `Normal` is default. Possible values (taken from AcView): Normal, Design, Preview, PivotTable, PivotChart, Report, Layout")]
            public TextStructure ViewType { get; set; } = new TextStructure("Normal");

            [Argument(Tooltip = "True to open readonly. It is the default value")]
            public BooleanStructure Readonly { get; set; } = new BooleanStructure(true);

            [Argument(Tooltip = "True to create new Report. False by default")]
            public BooleanStructure CreateNew { get; set; } = new BooleanStructure(false);
        }

        public AccessOpenReportCommand(AbstractScripter scripter) : base(scripter)
        { }

        public void Execute(Arguments arguments)
        {
            AccessManager.CurrentAccess.OpenReport(arguments.Name.Value, arguments.ViewType.Value, arguments.CreateNew.Value, arguments.Readonly.Value);
        }
    }
}