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

namespace G1ANT.Addon.MSOffice.Commands.Outlook
{
    [Command(Name = "outlook.sendandreceive", Tooltip = "This command initiates immediate delivery of all undelivered messages submitted in the current session")]
    class OutlookSendAndReceiveCommand : Command
    {
        public OutlookSendAndReceiveCommand(AbstractScripter scripter) : base(scripter)
        {
        }

        public void Execute(CommandArguments arguments)
        {
            var outlookManager = OutlookManager.CurrentOutlook;
            if (outlookManager != null)
            {
                outlookManager.SendAndReceive();
            }
            else
                throw new NullReferenceException("Current Outlook is not set.");
        }
    }
}
