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
using System;

namespace G1ANT.Addon.MSOffice
{
    [Command(Name = "access.getform", Tooltip = "Get detailed information about form defined in current project")]
    public class AccessGetFormCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Tooltip = "Name of form to get the details from", Required = true)]
            public TextStructure Name { get; set; }

            [Argument(Tooltip = "Name of a variable where the command's result will be stored")]
            public VariableStructure Result { get; set; } = new VariableStructure("result");
        }

        public AccessGetFormCommand(AbstractScripter scripter) : base(scripter)
        { }

        public void Execute(Arguments arguments)
        {
            var name = arguments.Name?.Object?.ToString();
            if (string.IsNullOrEmpty(name))
                throw new ArgumentNullException(nameof(arguments.Name));

            var currentAccess = AccessManager.CurrentAccess;
            var result = currentAccess.GetForm(name);

            Scripter.Variables.SetVariableValue(arguments.Result.Value, new JsonStructure(JObject.FromObject(result)));
        }
    }
}