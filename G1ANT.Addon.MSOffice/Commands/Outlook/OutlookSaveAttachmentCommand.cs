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
    [Command(Name = "outlook.saveattachment", Tooltip = "This command saves an attachment to a file")]
    public class OutlookSaveAttachmentCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Tooltip = "Email attachment to be saved", Required = true)]
            public OutlookAttachmentStructure Attachment { get; set; }

            [Argument(Tooltip = "Path to the saved file", Required = true)]
            public PathStructure Path { get; set; }
        }
        public OutlookSaveAttachmentCommand(AbstractScripter scripter) : base(scripter)
        {
        }

        public void Execute(Arguments arguments)
        {
            var outlookManager = OutlookManager.CurrentOutlook;
            if (outlookManager != null)
            {
                outlookManager.SaveAttachment(arguments.Attachment?.Value, arguments.Path?.Value);
            }
            else
                throw new NullReferenceException("Current Outlook is not set.");
        }
    }
}
