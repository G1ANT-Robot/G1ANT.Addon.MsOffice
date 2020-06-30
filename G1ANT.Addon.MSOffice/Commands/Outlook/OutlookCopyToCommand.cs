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
using Microsoft.Office.Interop.Outlook;
using System;

namespace G1ANT.Addon.MSOffice
{
    [Command(Name = "outlook.copyto", Tooltip = "This command is used to copy an individual email message or a whole folder to another location (Outlook folder)")]
    public class OutlookCopyToCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Required = true, Tooltip = "An item (a message or a folder) to be copied")]
            public Structure Item { get; set; }

            [Argument(Required = true, Tooltip = "Destination Outlook folder")]
            public OutlookFolderStructure DestinationFolder { get; set; }
        }
        public OutlookCopyToCommand(AbstractScripter scripter) : base(scripter)
        {
        }
        public void Execute(Arguments arguments)
        {
            var outlookManager = OutlookManager.CurrentOutlook;
            if (outlookManager == null)
                throw new NullReferenceException("Current Outlook is not set.");

            switch (arguments.Item)
            {
                case OutlookMailStructure mail:
                    var mailCopy = mail.Value.Copy();
                    mailCopy.Move(arguments.DestinationFolder.Value);
                    mailCopy.Save();
                    break;

                case OutlookFolderStructure folder:
                    folder.Value.CopyTo(arguments.DestinationFolder.Value);
                    break;

                default:
                    throw new NotSupportedException($"{arguments.Item.GetType()} is not supported.");
            }
        }
    }
}
