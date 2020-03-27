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

namespace G1ANT.Addon.MSOffice.Commands.Access.Forms
{
    [Command(Name = "access.getformrecordset", Tooltip = "Get detailed information about recordset of a form")]
    public class AccessGetFormRecordsetCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Tooltip = "Name of form to get the details from OR path to control of `SubForm` type (see `access.getcontrol` parameter `Name` for examples)", Required = true)]
            public TextStructure Name { get; set; }

            [Argument(Tooltip = "Name of a variable where the command's result will be stored")]
            public VariableStructure Result { get; set; } = new VariableStructure("result");
        }

        public AccessGetFormRecordsetCommand(AbstractScripter scripter) : base(scripter)
        { }

        public void Execute(Arguments arguments)
        {
            var form = AccessManager.CurrentAccess.GetForm(arguments.Name.Value, false, false, false);
            var recordset = form.Recordset.Value;

            Scripter.Variables.SetVariableValue(arguments.Result.Value, new JsonStructure(JObject.FromObject(recordset)));
        }
    }
}