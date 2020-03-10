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

namespace G1ANT.Addon.MSOffice.Access
{
    [Command(Name = "access.getitems", Tooltip = "Get items indexes and values")]
    public class AccessGetItemsCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Required = true, Tooltip = "Path to the control. See `access.findcontrol` tooltip for path examples")]
            public TextStructure Path { get; set; }

            [Argument(Tooltip = "Name of a variable where the command's result will be stored")]
            public VariableStructure Result { get; set; } = new VariableStructure("result");
        }

        public AccessGetItemsCommand(AbstractScripter scripter) : base(scripter)
        { }

        public void Execute(Arguments arguments)
        {
            var control = AccessManager.CurrentAccess.GetAccessControlByPath(arguments.Path.Value);
            var result = control.GetItems();

            Scripter.Variables.SetVariableValue(arguments.Result.Value, new JsonStructure(JArray.FromObject(result)));
        }
    }
}
