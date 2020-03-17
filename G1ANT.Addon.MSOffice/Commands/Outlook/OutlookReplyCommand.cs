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

namespace G1ANT.Addon.MSOffice
{
    [Command(Name = "outlook.reply", Tooltip = "This command creates a new variable of outlookmail structure which is a reply to a specified mail")]
    public class OutlookReplyCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Required = true, Tooltip = "Mail message to be replied to")]
            public OutlookMailStructure Mail { get; set; }

            [Argument(Required = true, Tooltip = "Name of a variable where the command's result will be stored. The variable will be of outlookmail structure")]
            public VariableStructure Result { get; set; }
        }
        public OutlookReplyCommand(AbstractScripter scripter) : base(scripter)
        {
        }

        public void Execute(Arguments arguments)
        {
            var outlookManager = OutlookManager.CurrentOutlook;
            if (outlookManager != null)
            {
                var replyMail = outlookManager.Reply(arguments.Mail?.Value);
                Scripter.Variables.SetVariableValue(arguments.Result.Value, new OutlookMailStructure(replyMail, null, Scripter));
            }
            else
                throw new NullReferenceException("Current Outlook is not set.");
        }
    }
}
