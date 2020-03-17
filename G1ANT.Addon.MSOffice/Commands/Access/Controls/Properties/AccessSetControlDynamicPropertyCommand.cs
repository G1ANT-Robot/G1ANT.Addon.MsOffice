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

namespace G1ANT.Addon.MSOffice.Commands.Access.Controls.Properties
{
    [Command(Name = "access.setcontroldynamicproperty", Tooltip = "Get value of control of form of Access application")]
    public class AccessSetControlDynamicPropertyCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Required = true, Tooltip = "Path to the control. See `access.getcontrol` tooltip for path examples")]
            public TextStructure Path { get; set; }

            [Argument(Required = true, Tooltip = "Name of property (see list of names at `Access forms and controls tree` panel")]
            public TextStructure NameOfProperty { get; set; }

            [Argument(Required = true, Tooltip = "Value of property")]
            public TextStructure Value { get; set; }
        }

        public AccessSetControlDynamicPropertyCommand(AbstractScripter scripter) : base(scripter)
        { }

        public void Execute(Arguments arguments)
        {
            var control = AccessManager.CurrentAccess.GetControlByPath(arguments.Path.Value);
            control.SetDynamicPropertyValue<object>(arguments.NameOfProperty.Value, arguments.Value.Value);
        }
    }
}