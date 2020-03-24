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
using Newtonsoft.Json.Linq;

namespace G1ANT.Addon.MSOffice.Commands.Access.Forms.Properties
{
    [Command(Name = "access.getformproperties", Tooltip = "Get list of value and name of properties of form of Access application")]
    public class AccessGetFormPropertiesCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Required = true, Tooltip = "Name of form")]
            public TextStructure Name { get; set; }

            [Argument(Required = true, Tooltip = "Set to true to load values of form properties. True by default")]
            public BooleanStructure GetValues { get; set; } = new BooleanStructure(true);

            [Argument(Tooltip = "Name of a variable where the command's result will be stored")]
            public VariableStructure Result { get; set; } = new VariableStructure("result");
        }

        public AccessGetFormPropertiesCommand(AbstractScripter scripter) : base(scripter)
        { }

        public void Execute(Arguments arguments)
        {
            var form = AccessManager.CurrentAccess.GetForm(arguments.Name.Value);
            var result = form.GetProperties(arguments.GetValues.Value);

            Scripter.Variables.SetVariableValue(arguments.Result.Value, new JsonStructure(JArray.FromObject(result)));
        }
    }
}