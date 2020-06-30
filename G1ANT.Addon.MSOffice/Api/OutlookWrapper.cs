/**
*    Copyright(C) G1ANT Ltd, All rights reserved
*    Solution G1ANT.Addon, Project G1ANT.Addon.MSOffice
*    www.g1ant.com
*
*    Licensed under the G1ANT license.
*    See License.txt file in the project root for full license information.
*
*/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Outlook;
using System.Reflection;
using System.IO;
using System.Diagnostics;
using System.Data.Linq.SqlClient;

namespace G1ANT.Addon.MSOffice
{
    public class OutlookWrapper
    {
        private OutlookWrapper() { }
        public OutlookWrapper(int id)
        {
            this.Id = id;
        }
        public int Id { get; set; }
        public bool IsMailFound { get; set; } = false;

        private Application application;

        public Application Application
        {
            get
            {
                try
                {
                    string version = application.Version;
                }
                catch (System.Exception ex)
                {
                    throw new InvalidOperationException($"Outlook instance could not be found. Most likely, it has been closed. Message: '{ex.Message}'.");
                }
                return application;
            }
            set { application = value; }
        }

        private MailItem mailItem = null;
        private NameSpace nameSpace = null;

        public void Open(bool display)
        {
            Application = new Application();
            nameSpace = Application.GetNamespace("MAPI");
            var defFolder = nameSpace.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
            if (display)
            {
                defFolder.Display();
                Application.ActiveExplorer().Activate();
            }
        }

        public void NewMessage(string to, string subject, string body, bool isHtmlBody, List<string> attachmentPath)
        {
            mailItem = Application.CreateItem(OlItemType.olMailItem);
            mailItem.To = to;
            mailItem.Subject = subject;

            if (isHtmlBody)
            {
                mailItem.BodyFormat = OlBodyFormat.olFormatHTML;
                mailItem.HTMLBody = body;
            }
            else
            {
                mailItem.Body = body;
            }

            var attachmentIndex = 1;

            foreach (var path in attachmentPath)
            {
                if (File.Exists(path))
                {
                    FileInfo file = new FileInfo(path);
                    mailItem.Attachments.Add(file.FullName, OlAttachmentType.olByValue, attachmentIndex++, file.Name);
                }
                else
                {
                    throw new FileNotFoundException("Attachement not found: " + path);
                }
            }

            mailItem.Display();
        }

        public void DiscardMail()
        {
            mailItem.Close(OlInspectorClose.olDiscard);
        }        

        private class MailDetails
        {
            public string EntryID { get; set; } = string.Empty;
            public string Subject { get; set; } = string.Empty;
            public string FolderID { get; set; } = string.Empty;
        }

        private List<MailDetails> ReturnAllMailsDetails()
        {

            List<MailDetails> mailDetailsList = new List<MailDetails>();
            List<string> names = new List<string>();
            foreach (MAPIFolder f in nameSpace.Folders)
            {
                MAPIFolder inboxFolder = f.Store.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
                Items items = inboxFolder.Items;
                foreach (var item in inboxFolder.Items)
                {
                    MailItem tmp = item as MailItem;
                    if (tmp != null)
                    {
                        MailDetails mailDetails = new MailDetails()
                        {
                            EntryID = tmp.EntryID,
                            Subject = tmp?.Subject ?? string.Empty,
                            FolderID = f.StoreID
                        };
                        mailDetailsList.Add(mailDetails);
                    }
                }
            }
            return mailDetailsList;
        }

        private List<MailDetails> ReturnAllFoundMailsDetails(List<MailDetails> mailDetailList, string search)
        {
            int howManySubjectsFound = mailDetailList.Where(x => x.Subject.Contains(search)).Count();
            List<MailDetails> foundMailDetailsList = new List<MailDetails>();
            if (howManySubjectsFound > 0)
            {
                foreach (var item in mailDetailList)
                {
                    if (item.Subject.Contains(search))
                    {
                        MailDetails foundMailDetails = new MailDetails()
                        {
                            EntryID = item.EntryID,
                            Subject = item.Subject,
                            FolderID = item.FolderID
                    };                     
                        foundMailDetailsList.Add(foundMailDetails);
                    }
                }
            }
            return foundMailDetailsList;
        }

        public List<MAPIFolder> GetFolders(MAPIFolder parent = null)
        {
            List<MAPIFolder> result = new List<MAPIFolder>();
            var folders = parent == null ? nameSpace.Folders : parent.Folders;
            if (folders != null)
            {
                foreach (MAPIFolder f in folders)
                    result.Add(f);
            }
            return result;
        }

        public MAPIFolder GetFolderByPath(string path, MAPIFolder parent = null)
        {
            string[] pathElements = path.Split('\\');
            if (pathElements.Count() == 0)
                return null;
            var folders = parent == null ? nameSpace.Folders : parent.Folders;
            if (folders != null)
            {
                foreach (MAPIFolder f in folders)
                {
                    if (pathElements[0] == f.Name)
                    {
                        if (pathElements.Count() == 1)
                            return f;
                        string rest = string.Join("\\", pathElements.Skip(1));
                        var fountFolder = GetFolderByPath(rest, f);
                        if (fountFolder != null)
                            return fountFolder;
                    }
                }
            }
            return null;
        }

        public List<MailItem> GetMails(MAPIFolder folder, bool onlyUnreaded)
        {
            if (folder == null)
                throw new NullReferenceException();
            List<MailItem> result = new List<MailItem>();
            Items mails = onlyUnreaded == true ? folder.Items.Restrict("[Unread]=true") : folder.Items;
            foreach (var item in mails)
                if (item is MailItem mail)
                    result.Add(mail);
            return result;
        }

        public List<Attachment> GetAttachments(MailItem mail)
        {
            if (mail == null)
                throw new NullReferenceException();
            List<Attachment> result = new List<Attachment>();
            foreach (var item in mail.Attachments)
                if (item is Attachment att)
                {
                    string fileName = att.FileName;
                    if (mail.HTMLBody.Contains($"cid:{fileName}") == false)
                        result.Add(att);
                }
            return result;
        }

        public IList<MailItem> FindMails(string search, bool showMail = false)
        {

            List<MailDetails> allMailsDetails = ReturnAllMailsDetails();
            List<MailDetails> allFoundMailsDetails = ReturnAllFoundMailsDetails(allMailsDetails, search);
            IsMailFound = allFoundMailsDetails?.Count() > 0 ? true : false;
            MailItem tmpMailItem;
            List<MailItem> foundMails = new List<MailItem>();
            foreach (var item in allFoundMailsDetails)
            {
                tmpMailItem = nameSpace.GetItemFromID(item.EntryID, item.FolderID) as MailItem;
                if (tmpMailItem != null) foundMails.Add(tmpMailItem);

            }
            if (showMail)
            {
                foreach (MailItem item in foundMails)
                {
                    item.Display();
                }
               
            }
            return foundMails;
        }

        public MailItem Reply(MailItem mail)
        {
            if (mail != null)
                return mail.Reply();
            else
                throw new NullReferenceException("Mail cannot be null.");
        }

        public void SaveAttachment(Attachment attachment, string path)
        {
            if (attachment != null)
            {
                try
                {
                    var attr = File.GetAttributes(path);
                    if (attr.HasFlag(FileAttributes.Directory))
                        path = Path.Combine(path, attachment.FileName);
                }
                catch
                { }
                attachment.SaveAsFile(path);
            }
            else
                throw new NullReferenceException("Attachment cannot be null.");
        }

        public void Send(MailItem mail)
        {
            if (mail != null)
                mail.Send();
            else
                mailItem.Send();
        }

        public void Close()
        {
            try
            {
                Application.Quit();
            }
            catch (System.Exception ex)
            {
                throw new ApplicationException($"Error occured while closing current Outlook Instance. Message: {ex.Message}");
            }
        }
    }
}
