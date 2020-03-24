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

namespace G1ANT.Addon.MSOffice.Commands.Access.Data.Queries
{
    [Command(Name = "access.getquerydetails", Tooltip = "This command gets details of Access Query")]
    public class AccessGetQueryDetailsCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Required = true, Tooltip = "Name of the query (one from `access.getqueries` command)")]
            public TextStructure Name { get; set; }

            [Argument(Tooltip = "Name of a variable where the command's result will be stored")]
            public VariableStructure Result { get; set; } = new VariableStructure("result");
        }

        public AccessGetQueryDetailsCommand(AbstractScripter scripter) : base(scripter)
        { }

        public void Execute(Arguments arguments)
        {
            var result = AccessManager.CurrentAccess.GetQueryDetails(arguments.Name.Value);
            Scripter.Variables.SetVariableValue(arguments.Result.Value, new JsonStructure(JObject.FromObject(result)));
        }
    }
}