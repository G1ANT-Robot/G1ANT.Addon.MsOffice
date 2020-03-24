/**
*    Copyright(C) G1ANT Ltd, All rights reserved
*    Solution G1ANT.Addon, Project G1ANT.Addon.MSOffice
*    www.g1ant.com
*
*    Licensed under the G1ANT license.
*    See License.txt file in the project root for full license inControlation.
*
*/

using G1ANT.Language;
using Newtonsoft.Json.Linq;

namespace G1ANT.Addon.MSOffice.Commands.Access.Controls.Properties
{
    [Command(Name = "access.getcontrolproperties", Tooltip = "Get list of value and name of properties of control of form of Access application")]
    public class AccessGetControlPropertiesCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Required = true, Tooltip = "Path to the control. See `access.getcontrol` tooltip for path examples")]
            public TextStructure Path { get; set; }

            [Argument(Required = true, Tooltip = "Set to true to load values of control properties. True by default")]
            public BooleanStructure GetValues { get; set; } = new BooleanStructure(true);

            [Argument(Tooltip = "Name of a variable where the command's result will be stored")]
            public VariableStructure Result { get; set; } = new VariableStructure("result");
        }

        public AccessGetControlPropertiesCommand(AbstractScripter scripter) : base(scripter)
        { }

        public void Execute(Arguments arguments)
        {
            var Control = AccessManager.CurrentAccess.GetControlByPath(arguments.Path.Value);
            var result = Control.GetProperties(arguments.GetValues.Value);

            Scripter.Variables.SetVariableValue(arguments.Result.Value, new JsonStructure(JArray.FromObject(result)));
        }
    }
}