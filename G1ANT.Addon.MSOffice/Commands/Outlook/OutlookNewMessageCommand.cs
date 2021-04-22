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
using System.Linq;
using System.Collections.Generic;

namespace G1ANT.Addon.MSOffice
{
    [Command(Name = "outlook.newmessage",Tooltip = "This command opens a new message window and fills it up with provided information", NeedsDelay = true)]
    public class OutlookNewMessageCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Tooltip = "Mail recipients")]
            public TextStructure To { get; set; } = new TextStructure(string.Empty);

            [Argument(Tooltip = "Mail subject")]
            public TextStructure Subject { get; set; } = new TextStructure(string.Empty);

            [Argument(Tooltip = "Mail body")]
            public TextStructure Body { get; set; } = new TextStructure(string.Empty);

            [Argument(Tooltip = "List of attachments (as their filepaths) to be included in a mail message. Elements should be separated with ❚ character (Ctrl+\\)")]
            public ListStructure Attachments { get; set; } = new ListStructure(new List<Structure>() { null });

            [Argument(Tooltip = "Name of a variable where the command's result will be stored")]
            public VariableStructure Result { get; set; } = new VariableStructure("result");

            [Argument(Required = false, Tooltip = "If set to `true`, indicates that the mail message body is in HTML")]
            public BooleanStructure IsBodyHtml { get; set; } = new BooleanStructure(false);
        }

        public OutlookNewMessageCommand(AbstractScripter scripter) : base(scripter)
        {
        }

        public void Execute(Arguments arguments)
        {
            var newMail = OutlookManager.CurrentOutlook.NewMessage(
                to: arguments.To.Value,
                subject: arguments.Subject.Value,
                body: arguments.Body.Value,
                isHtmlBody: arguments.IsBodyHtml.Value,
                attachmentPath: arguments.Attachments.Value.Where(a => a != null && !string.IsNullOrWhiteSpace(a.ToString())).Select(a => a.ToString()).ToList());

            Scripter.Variables.SetVariableValue(arguments.Result.Value, new OutlookMailStructure(newMail, null, Scripter));
        }
    }
}
