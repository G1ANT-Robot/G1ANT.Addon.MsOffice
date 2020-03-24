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

namespace G1ANT.Addon.MSOffice
{
    [Command(Name = "outlook.getrootfolders", Tooltip = "This command is used to return a list of all Outlook root folders")]
    public class OutlookGetRootFoldersCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Required = true, Tooltip = "Name of a variable where the command's result will be stored. The variable will be of list structure containing elements of outlookfolder structure")]
            public VariableStructure Result { get; set; }

        }
        public OutlookGetRootFoldersCommand(AbstractScripter scripter) : base(scripter)
        {
        }
        public void Execute(Arguments arguments)
        {
            var outlookManager = OutlookManager.CurrentOutlook;
            if (outlookManager != null)
            {
                var folders = outlookManager.GetFolders();
                var list = new List<object>();
                foreach (var f in folders)
                    list.Add(new OutlookFolderStructure(f, null, Scripter));
                Scripter.Variables.SetVariableValue(arguments.Result.Value, new ListStructure(list, null, Scripter));
            }
            else
                throw new NullReferenceException("Current Outlook is not set.");
        }
    }
}
