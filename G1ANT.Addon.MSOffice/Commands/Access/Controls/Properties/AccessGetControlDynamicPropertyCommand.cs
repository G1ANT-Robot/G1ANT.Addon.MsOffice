/**
*    Copyright(C) G1ANT Ltd, All rights reserved
*    Solution G1ANT.Addon, Project G1ANT.Addon.MSOffice
*    www.g1ant.com
*
*    Licensed under the G1ANT license.
*    See License.txt file in the project root for full license information.
*
*/

using G1ANT.Addon.MSOffice.Api;
using G1ANT.Addon.MSOffice.Api.Access;
using G1ANT.Language;
using System;

namespace G1ANT.Addon.MSOffice.Commands.Access.Controls.Properties
{
    [Command(Name = "access.getcontroldynamicproperty", Tooltip = "Get value of dynamic property of control of form of Access application")]
    public class AccessGetControlDynamicPropertyCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Required = true, Tooltip = "Path to the control. See `access.getcontrol` tooltip for path examples")]
            public TextStructure Path { get; set; }

            [Argument(Required = true, Tooltip = "Name of property (see list of names at `Access forms and controls tree` panel")]
            public TextStructure NameOfProperty { get; set; }

            [Argument(Tooltip = "Name of a variable where the command's result will be stored")]
            public VariableStructure Result { get; set; } = new VariableStructure("result");
        }

        public AccessGetControlDynamicPropertyCommand(AbstractScripter scripter) : base(scripter)
        { }

        public void Execute(Arguments arguments)
        {
            var control = AccessManager.CurrentAccess.GetControlByPath(arguments.Path.Value);
            if (control.TryGetDynamicPropertyValue(arguments.NameOfProperty.Value, out object value))
            {
                var result = new StructureConverter().Convert(value);
                Scripter.Variables.SetVariableValue(arguments.Result.Value, result);
            }
            else throw new ApplicationException($"Error getting dynamic property value for field {arguments.NameOfProperty.Value}");
        }
    }
}