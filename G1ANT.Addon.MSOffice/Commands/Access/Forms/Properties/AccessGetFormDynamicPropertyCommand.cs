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

namespace G1ANT.Addon.MSOffice.Commands.Access.Forms.Properties
{
    [Command(Name = "access.getformdynamicproperty", Tooltip = "Get value of dynamic property of form of Access application")]
    public class AccessGetFormDynamicPropertyCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Required = true, Tooltip = "Name of the form")]
            public TextStructure Name { get; set; }

            [Argument(Required = true, Tooltip = "Name of property (see list of names at `Access forms and Forms tree` panel")]
            public TextStructure NameOfProperty { get; set; }

            [Argument(Tooltip = "Name of a variable where the command's result will be stored")]
            public VariableStructure Result { get; set; } = new VariableStructure("result");
        }

        public AccessGetFormDynamicPropertyCommand(AbstractScripter scripter) : base(scripter)
        { }

        public void Execute(Arguments arguments)
        {
            var form = AccessManager.CurrentAccess.GetForm(arguments.Name.Value);
            if (form.TryGetDynamicPropertyValue(arguments.NameOfProperty.Value, out object value))
            {
                var result = new StructureConverter().Convert(value);
                Scripter.Variables.SetVariableValue(arguments.Result.Value, result);
            }
            else throw new ApplicationException($"Error getting dynamic property value for field {arguments.NameOfProperty.Value}");
        }
    }
}