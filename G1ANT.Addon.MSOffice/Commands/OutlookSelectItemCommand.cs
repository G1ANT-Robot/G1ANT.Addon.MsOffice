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
using System.Collections.Generic;

namespace G1ANT.Addon.MSOffice.Commands
{
    [Command(Name = "outlook.selectitem", Tooltip = "This command selects a mail or a folder element in Outlook’s user interface")]
    public class OutlookSelectItemCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Required = true, Tooltip = "Item — a mail message or a folder — to be selected")]
            public Structure Item { get; set; }
        }
        public OutlookSelectItemCommand(AbstractScripter scripter) : base(scripter)
        {
        }
        public void Execute(Arguments arguments)
        {
            var outlookManager = OutlookManager.CurrentOutlook;
            if (outlookManager != null)
            {
                if (arguments.Item is OutlookMailStructure mail)
                {
                    var explorer = outlookManager.Application.ActiveExplorer();
                    explorer.ClearSelection();
                    if (explorer.IsItemSelectableInView(mail.Value))
                        explorer.AddToSelection(mail.Value);
                    else
                        mail.Value.Display(false);
                }
                else if (arguments.Item is OutlookFolderStructure folder)
                    outlookManager.Application.ActiveExplorer().CurrentFolder = folder.Value;
                else
                    throw new NotSupportedException($"{arguments.Item.GetType()} is not supported.");
            }
            else
                throw new NullReferenceException("Current Outlook is not set.");
        }
    }
}
