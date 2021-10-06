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
using System.Linq;

namespace G1ANT.Addon.MSOffice
{
    [Structure(Name = "OutlookMail", AutoCreate = false, Tooltip = "This structure stores information about a mail message, which was retrieved from the Outlook folder with the `outlook.getfolder` command")]
    public class OutlookMailStructure : StructureTyped<MailItem>
    {
        private const string IdIndex = "id";
        private const string FromIndex = "from";
        private const string CcIndex = "cc";
        private const string BccIndex = "bcc";
        private const string AccountIndex = "account";
        private const string SubjectIndex = "subject";
        private const string BodyIndex = "body";
        private const string HtmlBodyIndex = "htmlbody";
        private const string AttachmentsIndex = "attachments";
        private const string UnreadIndex = "unread";

        public OutlookMailStructure(string value, string format = "", AbstractScripter scripter = null) :
            base(value, format, scripter)
        {
            Init();
        }

        public OutlookMailStructure(object value, string format = null, AbstractScripter scripter = null)
            : base(value, format, scripter)
        {
            Init();
        }

        protected void Init()
        {
            Indexes.Add(IdIndex);
            Indexes.Add(SubjectIndex);
            Indexes.Add(AttachmentsIndex);
            Indexes.Add(BodyIndex);
            Indexes.Add(HtmlBodyIndex);
            Indexes.Add(FromIndex);
            Indexes.Add(CcIndex);
            Indexes.Add(BccIndex);
            Indexes.Add(AccountIndex);
            Indexes.Add(UnreadIndex);
        }

        public override Structure Get(string index = "")
        {
            if (string.IsNullOrWhiteSpace(index))
                return new OutlookMailStructure(Value, Format);
            switch (index.ToLower())
            {
                case IdIndex:
                    return new TextStructure(Value.EntryID, null, Scripter);
                case SubjectIndex:
                    return new TextStructure(Value.Subject, null, Scripter);
                case BodyIndex:
                    return new TextStructure(Value.Body, null, Scripter);
                case HtmlBodyIndex:
                    return new TextStructure(Value.HTMLBody, null, Scripter);
                case FromIndex:
                    return new TextStructure(Value.SenderEmailAddress, null, Scripter);
                case CcIndex:
                    return new TextStructure(GetRecipientListOfType(OlMailRecipientType.olCC), null, Scripter);
                case BccIndex:
                    return new TextStructure(GetRecipientListOfType(OlMailRecipientType.olBCC), null, Scripter);
                case AccountIndex:
                    return new TextStructure(Value.SendUsingAccount.SmtpAddress, null, Scripter);
                case UnreadIndex:
                    return new BooleanStructure(Value.UnRead, null, Scripter);
                case AttachmentsIndex:
                    {
                        var outlookManager = OutlookManager.CurrentOutlook;
                        if (outlookManager != null)
                        {
                            var attachements = outlookManager.GetAttachments(Value);
                            var list = new List<object>();
                            foreach (var a in attachements)
                                list.Add(new OutlookAttachmentStructure(a, null, Scripter));
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
            {
                throw new ArgumentNullException(nameof(structure));
            }
            else
            {
                switch (index.ToLower())
                {
                    case SubjectIndex:
                        Value.Subject = structure.ToString();
                        break;
                    case BodyIndex:
                        Value.Body = structure.ToString();
                        break;
                    case HtmlBodyIndex:
                        Value.HTMLBody = structure.ToString();
                        break;
                    case UnreadIndex:
                        Value.UnRead = Convert.ToBoolean(structure.Object);
                        break;
                    case CcIndex:
                        if (structure is TextStructure cc)
                            SetRecipientListOfType(OlMailRecipientType.olCC, cc.ToString());
                        else
                            throw new ArgumentException("Should be text separated by all emails by ';'");
                        break;
                    case BccIndex:
                        if (structure is TextStructure bcc)
                            SetRecipientListOfType(OlMailRecipientType.olBCC, bcc.ToString());
                        else
                            throw new ArgumentException("Should be text separated by all emails by ';'");
                        break;
                    case AccountIndex:
                        {
                            Accounts accounts = Value.Session.Accounts;
                            foreach (Account account in accounts)
                            {
                                // When the e-mail address matches, return the account. 
                                if (account.SmtpAddress == structure.ToString())
                                {
                                    Value.SendUsingAccount = account;
                                    Value.Sender = account.CurrentUser.AddressEntry;
                                    return;
                                }
                            }
                            throw new ArgumentException($"Cannot find outlook account '{structure.ToString()}'");
                        }
                    default:
                        throw new ArgumentException($"Unknown index '{index}'");
                }
            }
        }

        public override string ToString(string format)
        {
            return Get(FromIndex)?.ToString();
        }

        protected override MailItem Parse(string value, string format = null)
        {
            return null;
        }

        private void AddRecipient(string recipient, OlMailRecipientType recipientType)
        {
            var newRecipient = Value.Recipients.Add(recipient);
            newRecipient.Type = (int)recipientType;
        }

        private void SetRecipientListOfType(OlMailRecipientType recipientType, string recipients)
        {
            for (int idx = Value.Recipients.Count; idx > 0; idx--)
            {
                var recipient = Value.Recipients[idx];
                if (recipient.Type == (int)recipientType)
                    Value.Recipients.Remove(idx);
            }
            var list = recipients.Split(';');
            list.Where(x => !string.IsNullOrEmpty(x)).ToList().ForEach(x => AddRecipient(x, recipientType));
            Value.Recipients.ResolveAll();
        }

        private string GetRecipientListOfType(OlMailRecipientType recipientType)
        {
            return Value
                   .Recipients.Cast<Recipient>()
                   .Where(recipient => recipient.Type == (int)recipientType)
                   .Aggregate(string.Empty, (current, recipient) => current + (recipient.Address + ";"));
        }
    }
}
