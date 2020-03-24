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

namespace G1ANT.Addon.MSOffice.Commands.Access.Data
{
    [Command(Name = "access.getapplicationdetails", Tooltip = "This command Access application and opened project details")]
    public class AccessGetApplicationDetailsCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Tooltip = "Name of a variable where the command's result will be stored")]
            public VariableStructure Result { get; set; } = new VariableStructure("result");
        }

        public AccessGetApplicationDetailsCommand(AbstractScripter scripter) : base(scripter)
        { }

        public void Execute(Arguments arguments)
        {
            var result = AccessManager.CurrentAccess.GetApplicationDetails();
            Scripter.Variables.SetVariableValue(arguments.Result.Value, new JsonStructure(JObject.FromObject(result)));
        }
    }
}