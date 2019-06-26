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



namespace G1ANT.Addon.MSOffice
{
    [Command(Name = "outlook.findmails", Tooltip = "This command searches subjects of Inbox messages and returns all emails that contain a required keyword", NeedsDelay = true)]
    public class OutlookFindMailsCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Required = true,Tooltip = "Word to be searched for in a message subject")]
            public TextStructure Search { get; set; }

            [Argument(Tooltip = "If set to `true`, G1ANT.Robot will show all emails meeting the criteria")]
            public BooleanStructure ShowMail { get; set; } = new BooleanStructure(false);

            [Argument(Tooltip = "Name of a variable where the command's result will be stored")]
            public VariableStructure Result { get; set; } = new VariableStructure("result");

        }
        public OutlookFindMailsCommand(AbstractScripter scripter) : base(scripter)
        {
        }
        public void Execute(Arguments arguments)
        {
            var outlookManager = OutlookManager.CurrentOutlook;
            var search = arguments.Search.Value;
            var showMails = arguments.ShowMail.Value;
            if (search!="")
            {
                outlookManager.FindMails(search, showMails);
            }
            Scripter.Variables.SetVariableValue(arguments.Result.Value, new Language.TextStructure(outlookManager.IsMailFound.ToString()));
        }
    }
}
