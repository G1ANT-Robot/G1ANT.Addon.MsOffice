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

namespace G1ANT.Addon.MSOffice.Commands.Access.Forms
{
    [Command(Name = "access.openform", Tooltip = "This command opens an Access Form")]
    public class AccessOpenFormCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Tooltip = "Name of the form to open", Required = true)]
            public TextStructure Name { get; set; }

            [Argument(Tooltip = "Open mode. `Normal` is default. Possible values (taken from AcFormView): Normal, Design, Preview, FormDS, FormPivotTable, FormPivotChart, Layout")]
            public TextStructure ViewType { get; set; } = new TextStructure("Normal");

            [Argument(Tooltip = "True to open readonly. It is the default value")]
            public BooleanStructure Readonly { get; set; } = new BooleanStructure(true);

            [Argument(Tooltip = "True to create new form. False by default")]
            public BooleanStructure CreateNew { get; set; } = new BooleanStructure(false);

            [Argument(Tooltip = "Open properties of form. False by default")]
            public BooleanStructure OpenProperties { get; set; } = new BooleanStructure(false);

            [Argument(Tooltip = "Specifies the window mode in which the form opens. The default value is `WindowNormal`. Possible values (taken from AcWindowMode): WindowNormal, Hidden, Icon, Dialog")]
            public TextStructure WindowMode { get; set; } = new TextStructure("WindowNormal");

            [Argument(Tooltip = "Valid name of a query in the current database")]
            public TextStructure FilterName { get; set; }

            [Argument(Tooltip = "Valid SQL WHERE clause without the word WHERE.")]
            public TextStructure WhereCondition { get; set; }

            [Argument(Tooltip = "Used to set the form's OpenArgs property. This setting can then be used by code in a form module, such as the Open event procedure. The OpenArgs property can also be referred to in macros and expressions")]
            public TextStructure OpenArgs { get; set; }
        }

        public AccessOpenFormCommand(AbstractScripter scripter) : base(scripter)
        { }

        public void Execute(Arguments arguments)
        {
            AccessManager.CurrentAccess.OpenForm(
                arguments.Name.Value,
                arguments.ViewType.Value,
                arguments.CreateNew.Value,
                arguments.Readonly.Value,
                arguments.OpenProperties.Value,
                arguments.WindowMode.Value,
                arguments.FilterName?.Value,
                arguments.WhereCondition?.Value,
                arguments.OpenArgs?.Value
           );
        }
    }
}