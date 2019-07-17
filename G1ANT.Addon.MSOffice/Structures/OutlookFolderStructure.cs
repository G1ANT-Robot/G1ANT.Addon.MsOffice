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
using System.Collections.Generic;

namespace G1ANT.Addon.MSOffice
{
    [Structure(Name = "OutlookFolder", AutoCreate = false, Tooltip = "This structure stores information about the Outlook folder, which was retrieved with the `outlook.getfolder` command")]
    public class OutlookFolderStructure : StructureTyped<MAPIFolder>
    {
        const string NameIndex = "name";
        const string FolderPathIndex = "folderpath";
        const string FoldersIndex = "folders";
        const string MailsIndex = "mails";
        const string UnreadIndex = "unread";

        /// <summary>
        /// Deprecated
        /// </summary>
        const string UnreadedIndex = "unreaded";

        public OutlookFolderStructure(string value, string format = "", AbstractScripter scripter = null) :
            base(value, format, scripter)
        {
            Init();
        }

        public OutlookFolderStructure(object value, string format = null, AbstractScripter scripter = null)
            : base(value, format, scripter)
        {
            Init();
        }

        protected void Init()
        {
            Indexes.Add(NameIndex);
            Indexes.Add(FoldersIndex);
            Indexes.Add(MailsIndex);
            Indexes.Add(UnreadIndex);
        }

        public override Structure Get(string index = "")
        {
            if (string.IsNullOrWhiteSpace(index))
                return new OutlookFolderStructure(Value, Format);

            index = index.ToLower();

            switch (index)
            {
                case NameIndex:
                    return new TextStructure(Value.Name, null, Scripter);
                case FolderPathIndex:
                    return new TextStructure(Value.FolderPath, null, Scripter);
                case FoldersIndex:
                    {
                        var outlookManager = OutlookManager.CurrentOutlook;
                        if (outlookManager != null)
                        {
                            var folders = outlookManager.GetFolders(Value);
                            var list = new List<object>();
                            foreach (var f in folders)
                                list.Add(new OutlookFolderStructure(f, null, Scripter));
                            return new ListStructure(list, null, Scripter);
                        }
                        else
                            throw new NullReferenceException("Current Outlook is not set.");
                    }
                case MailsIndex:
                case UnreadedIndex:
                case UnreadIndex:
                    {
                        var outlookManager = OutlookManager.CurrentOutlook;
                        if (outlookManager != null)
                        {                            
                            bool unreaded = index == UnreadedIndex || index == UnreadIndex;
                            var mails = outlookManager.GetMails(Value, unreaded);
                            var list = new List<object>();
                            foreach (var m in mails)
                                list.Add(new OutlookMailStructure(m, null, Scripter));
                            return new ListStructure(list, null, Scripter);
                        }
                        else
                            throw new NullReferenceException("Current Outlook is not set.");
                    }
            }

            throw new ArgumentException($"Unknown index '{index}'");
        }

        public override void Set(Structure structure, string index = null)
        {
            if (structure == null || structure.Object == null)
                throw new ArgumentNullException(nameof(structure));
            else
                throw new ArgumentException($"Unknown index '{index}'");
        }

        public override string ToString(string format)
        {
            return Get(NameIndex)?.ToString();
        }

        protected override MAPIFolder Parse(string value, string format = null)
        {
            return null;
        }
    }
}
