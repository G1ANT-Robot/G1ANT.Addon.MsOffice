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
    [Command(Name = "access.getcontroldynamicproperties", Tooltip = "Get list of value and name of dynamic properties of control of form of Access application")]
    public class AccessGetControlDynamicPropertiesCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Required = true, Tooltip = "Path to the control. See `access.getcontrol` tooltip for path examples")]
            public TextStructure Path { get; set; }

            [Argument(Tooltip = "Name of a variable where the command's result will be stored")]
            public VariableStructure Result { get; set; } = new VariableStructure("result");
        }

        public AccessGetControlDynamicPropertiesCommand(AbstractScripter scripter) : base(scripter)
        { }

        public void Execute(Arguments arguments)
        {
            var control = AccessManager.CurrentAccess.GetControlByPath(arguments.Path.Value);
            var result = control.GetDynamicProperties();

            Scripter.Variables.SetVariableValue(arguments.Result.Value, new JsonStructure(JArray.FromObject(result)));
        }
    }
}