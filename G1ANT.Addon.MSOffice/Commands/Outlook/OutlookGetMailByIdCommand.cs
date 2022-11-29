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
using System;
using System.Net.NetworkInformation;

namespace G1ANT.Addon.MSOffice
{
    [Command(Name = "outlook.getmailbyid", Tooltip = "This command returns email by its id", NeedsDelay = true)]
    public class OutlookGetMailByIdCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Required = true, Tooltip = "Mail id")]
            public TextStructure Id { get; set; }

            [Argument(Tooltip = "Name of a variable where the command's result will be stored")]
            public VariableStructure Result { get; set; } = new VariableStructure("result");

        }
        public OutlookGetMailByIdCommand(AbstractScripter scripter) : base(scripter)
        {
        }
        public void Execute(Arguments arguments)
        {
            var outlookManager = OutlookManager.CurrentOutlook;
            var mail = outlookManager.GetMailById(arguments.Id?.Value);
            if (mail == null)
                throw new ApplicationException($"Mail with id '{arguments.Id?.Value}' cannot be found.");
            Scripter.Variables.SetVariableValue(arguments.Result.Value, new OutlookMailStructure(mail, null, Scripter));
        }
    }
}
