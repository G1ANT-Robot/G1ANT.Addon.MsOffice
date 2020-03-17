/**
*    Copyright(C) G1ANT Ltd, All rights reserved
*    Solution G1ANT.Addon, Project G1ANT.Addon.MSOffice
*    www.g1ant.com
*
*    Licensed under the G1ANT license.
*    See License.txt file in the project root for full license information.
*
*/

using G1ANT.Addon.MSOffice.Api.Access;
using G1ANT.Language;

namespace G1ANT.Addon.MSOffice.Commands.Access.Forms.Properties
{
    [Command(Name = "access.setformdynamicproperty", Tooltip = "Get value of form of Access application")]
    public class AccessSetFormDynamicPropertyCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Required = true, Tooltip = "Name of form")]
            public TextStructure Name { get; set; }

            [Argument(Required = true, Tooltip = "Name of property (see list of names at `Access forms and Forms tree` panel")]
            public TextStructure NameOfProperty { get; set; }

            [Argument(Required = true, Tooltip = "Value of property")]
            public TextStructure Value { get; set; }
        }

        public AccessSetFormDynamicPropertyCommand(AbstractScripter scripter) : base(scripter)
        { }

        public void Execute(Arguments arguments)
        {
            var form = AccessManager.CurrentAccess.GetForm(arguments.Name.Value);
            form.SetDynamicPropertyValue(arguments.NameOfProperty.Value, arguments.Value.Value);
        }
    }
}