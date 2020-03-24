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

namespace G1ANT.Addon.MSOffice.Commands.Access.Data.Queries
{
    [Command(Name = "access.openquery", Tooltip = "This command opens an Access Query")]
    public class AccessOpenQueryCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Tooltip = "Name of the query to open", Required = true)]
            public TextStructure Name { get; set; }

            [Argument(Tooltip = "Open mode. `Normal` is default. Possible values (taken from AcView): Normal, Design, Preview, PivotTable, PivotChart, Report, Layout")]
            public TextStructure ViewType { get; set; } = new TextStructure("Normal");

            [Argument(Tooltip = "True to open readonly. It is the default value")]
            public BooleanStructure Readonly { get; set; } = new BooleanStructure(true);

            [Argument(Tooltip = "True to create new query. False by default")]
            public BooleanStructure CreateNew { get; set; } = new BooleanStructure(false);
        }

        public AccessOpenQueryCommand(AbstractScripter scripter) : base(scripter)
        { }

        public void Execute(Arguments arguments)
        {
            AccessManager.CurrentAccess.OpenQuery(arguments.Name.Value, arguments.ViewType.Value, arguments.CreateNew.Value, arguments.Readonly.Value);
        }
    }
}