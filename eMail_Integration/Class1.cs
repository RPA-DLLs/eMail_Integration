using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.Exchange.WebServices.Data;

namespace eMail_Integration
{


    class Connection
    {
        public ExchangeService InitializeConnection(string EmailAddress, string EmailPassword)
        {
            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2013_SP1);


            service.Credentials = new WebCredentials(EmailAddress, EmailPassword);
            service.AutodiscoverUrl(EmailAddress, RedirectionUrlValidationCallback);
            
            return service;

        }

        private static bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            // The default for the validation callback is to reject the URL.
            bool result = false;

            Uri redirectionUri = new Uri(redirectionUrl);

            // Validate the contents of the redirection URL. In this simple validation
            // callback, the redirection URL is considered valid if it is using HTTPS
            // to encrypt the authentication credentials. 
            if (redirectionUri.Scheme == "https")
            {
                result = true;
            }
            return result;
        }
    }

    class ExchangeCommands
    {
        public void DownloadEmail(Item email, EmailMessage message, string AttachmentSavePath, int emailID = 0)
        {
            email.Load(new PropertySet(ItemSchema.MimeContent));
            var mimeContent = email.MimeContent;
            Regex rgx = new Regex("[^a-zA-Z0-9 ]");
            string emailPath = "";
            string emailfilename = "";

            if (message.Subject is null)
            {
                emailfilename = "NoSubject";
            }
            else
            {
                emailfilename = rgx.Replace(message.Subject.ToString(), "");
            }



            if ((AttachmentSavePath.Length + emailfilename.Length - 250) >= 0)
            {
                int maxlenght = 250 - AttachmentSavePath.Length;
                emailfilename = emailfilename.Substring(0, maxlenght);
            }

            if (emailID == 0)
            {
                emailPath = AttachmentSavePath + emailfilename + ".eml";
            }
            else
            {
                emailPath = AttachmentSavePath + emailID + @"\" + emailfilename + ".eml";
            }


            using (var fileStream = new FileStream(emailPath, FileMode.Create))
            {
                fileStream.Write(mimeContent.Content, 0, mimeContent.Content.Length);
            }

        }

        public void ChangeEmailStatus(EmailMessage message, string MarkEmailAs)
        {
            if (MarkEmailAs == "read")
            {
                message.IsRead = true;
                message.Update(ConflictResolutionMode.AutoResolve);
            }
            else if (MarkEmailAs == "unread")
            {
                message.IsRead = false;
                message.Update(ConflictResolutionMode.AutoResolve);
            }
        }

        public void MoveEmail(Item email, FolderId DestinationFolder)
        {
            email.Move(DestinationFolder);
        }

        public string DownloadAttachments(EmailMessage message, string AttachmentSavePath, string[] EmailAttachment, string RetainOriginalAttName, string emailAttachmentLocation, int emailID = 0)
        {
            try
            {


                string emailAttachments = "";
                int fileID = 0;

                foreach (Attachment attachment in message.Attachments)
                {



                    if (attachment.IsInline == false)
                    {

                        string emailAttachmentSaveLocation = "";


                        if ((EmailAttachment.Any(attachment.Name.Contains) == false) && EmailAttachment[0] != "")
                        {
                            continue;
                        }
                        string extension = "";
                        string filename = "";
                        Regex rgx = new Regex("[^a-zA-Z0-9 .]");
                        if (attachment is ItemAttachment)
                        {
                            ItemAttachment fileAttachment = attachment as ItemAttachment;
                            extension = ".eml";
                            filename = fileAttachment.Name.ToString();

                            filename = rgx.Replace(filename, "");
                            filename = filename + extension;
                        }
                        else
                        {
                            FileAttachment fileAttachment = attachment as FileAttachment;
                            filename = fileAttachment.Name.ToString();
                            filename = rgx.Replace(filename, "");
                            extension = Path.GetExtension(filename);
                        }


                        if (RetainOriginalAttName == "true")
                        {


                            if (emailID == 0)
                            {
                                emailAttachmentLocation = AttachmentSavePath + filename;
                            }
                            else
                            {
                                emailAttachmentLocation = AttachmentSavePath + emailID + @"\" + filename;
                            }

                            if (emailAttachments == "")
                            {
                                emailAttachments = filename;
                            }
                            else
                            {
                                emailAttachments = emailAttachments + "|" + filename;
                            }
                        }
                        else
                        {

                            fileID++;


                            if (emailID == 0)
                            {
                                emailAttachmentLocation = AttachmentSavePath + fileID + extension;
                            }
                            else
                            {
                                emailAttachmentLocation = AttachmentSavePath + emailID + @"\" + fileID + extension;
                            }

                            if (emailAttachments == "")
                            {
                                emailAttachments = fileID + extension;
                            }
                            else
                            {
                                emailAttachments = emailAttachments + "|" + fileID + extension;
                            }
                        }

                        emailAttachmentSaveLocation = emailAttachmentLocation;

                        if (attachment is ItemAttachment)
                        {
                            ItemAttachment fileAttachment = attachment as ItemAttachment;
                            fileAttachment.Load(EmailMessageSchema.MimeContent);

                            // MimeContent.Content will give you the byte[] for the ItemAttachment
                            // Now all you have to do is write the byte[] to a file
                            File.WriteAllBytes(emailAttachmentLocation, fileAttachment.Item.MimeContent.Content);
                        }
                        else
                        {
                            FileAttachment fileAttachment = attachment as FileAttachment;
                            fileAttachment.Load(emailAttachmentLocation);
                        }




                    }
                }


                return emailAttachments;

            }
            catch (Exception e)
            {

                throw;
            }
        }
        public SearchFilter.IsEqualTo CreateFilter(string EmailType, SearchFilter.IsEqualTo sf)
        {

            if (EmailType == "unread")
            {
                sf = new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, false);
            }
            else if (EmailType == "read")
            {
                sf = new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, true);
            }

            return sf;
        }


        public FindItemsResults<Item> SetFolder(FindItemsResults<Item> EmailFolder, string FolderName, ExchangeService service, SearchFilter.IsEqualTo sf, string EmailAddress, string MappedMailboxAddress, Int32 ItemsAmount = 1)
        {
            FindFoldersResults Subfolders = null;

            if (FolderName == "inbox")
            {
                FolderId folderToAccess = new FolderId(WellKnownFolderName.Inbox, EmailAddress);

                if (MappedMailboxAddress != "")
                {
                    folderToAccess = new FolderId(WellKnownFolderName.Inbox, MappedMailboxAddress);
                }


                if (sf != null)
                {
                    EmailFolder = service.FindItems(folderToAccess, sf, new ItemView(ItemsAmount));
                }
                else
                {
                    EmailFolder = service.FindItems(folderToAccess, new ItemView(ItemsAmount));
                }

            }
            else
            {
                Subfolders = service.FindFolders(WellKnownFolderName.Inbox, new FolderView(int.MaxValue));

                foreach (var subfolder in Subfolders.Folders)
                {
                    if (FolderName == subfolder.DisplayName.ToLower())
                    {
                        if (sf != null)
                        {
                            EmailFolder = service.FindItems(subfolder.Id, sf, new ItemView(ItemsAmount));
                        }
                        else
                        {
                            EmailFolder = service.FindItems(subfolder.Id, new ItemView(ItemsAmount));
                        }
                        Subfolders = null;
                        break;
                    }

                }

                Subfolders = null;
            }


            return EmailFolder;
        }

        public FolderId GetSubfolerID(FolderId DestinationFolder, string MoveEmailToFolder, ExchangeService service)
        {
            FindFoldersResults Subfolders = null;

            Subfolders = service.FindFolders(WellKnownFolderName.Inbox, new FolderView(int.MaxValue));

            foreach (var subfolder in Subfolders.Folders)
            {
                if (MoveEmailToFolder == subfolder.DisplayName.ToLower())
                {
                    DestinationFolder = subfolder.Id;
                    Subfolders = null;
                    break;
                }

            }
            Subfolders = null;



            return DestinationFolder;
        }



    }


    public class Integration
    {
        public List<string> SendEmail(string EmailAddress, string EmailPassword, string[] EmailTo, string[] EmailCC, string[] EmailBCC, string[] Subject, string[] Body, string[] IsReadReceiptRequested, string[] IsDeliveryReceiptRequested, string[] EmailAttachments, string[] HighEmailImportance, string[] EmailEmbedImagePath, string[] EmailEmbedImigeFileName)
        {
            List<string> results = new List<string>();
            try
            {
                string EmailCCItem = "";
                string EmailBCCItem = "";
                int i = 0;
                //Establishing connection to the Exchange server
                Connection conn = new Connection();
                ExchangeService service = conn.InitializeConnection(EmailAddress, EmailPassword);

                foreach (string item in EmailTo)
                {


                    EmailMessage message = new EmailMessage(service);

                    string EmailToItem = item.Replace(" ", "");
                    if (EmailCC[i] != "")
                    {
                        EmailCCItem = EmailCC[i].Replace(" ", "");
                        List<string> RecepientsCC = EmailCCItem.Split(';').ToList();
                        message.CcRecipients.AddRange(RecepientsCC);
                    }
                    if (EmailBCC[i] != "")
                    {
                        EmailBCCItem = EmailBCC[i].Replace(" ", "");
                        List<string> RecepientsBCC = EmailBCCItem.Split(';').ToList();
                        message.BccRecipients.AddRange(RecepientsBCC);
                    }
                    
                    

                    List<string> Recepients = EmailToItem.Split(';').ToList();



                    List<string> Attachments = EmailAttachments[i].Split('|').ToList();



                    string IsReadReceiptRequestedItem = IsReadReceiptRequested[i].ToLower();
                    string IsDeliveryReceiptRequestedItem = IsDeliveryReceiptRequested[i].ToLower();
                    HighEmailImportance[i] = HighEmailImportance[i].ToLower();

                    //Message properties
                    message.ToRecipients.AddRange(Recepients);


                    message.Subject = Subject[i];

                    //Insert Body of the email
                    message.Body = Body[i];
                    message.Body.BodyType = BodyType.HTML;

                    //Embed image in email Body
                    if (EmailEmbedImagePath[i] != "" && EmailEmbedImigeFileName[i] != "")
                    {
                        string file = EmailEmbedImagePath[i] + EmailEmbedImigeFileName[i];
                        message.Attachments.AddFileAttachment(EmailEmbedImigeFileName[i], file);
                        message.Attachments[0].IsInline = true;
                        message.Attachments[0].ContentId = EmailEmbedImigeFileName[i];
                    }

                    if (HighEmailImportance[i] == "true")
                    {
                        message.Importance = Importance.High;
                    }

                    if (IsReadReceiptRequestedItem == "true")
                    {
                        message.IsReadReceiptRequested = true;
                    }
                    else
                    {
                        message.IsReadReceiptRequested = false;
                    }
                    if (IsDeliveryReceiptRequestedItem == "true")
                    {
                        message.IsDeliveryReceiptRequested = true;
                    }
                    else
                    {
                        message.IsDeliveryReceiptRequested = false;
                    }

                    if (Attachments[0] != "")
                    {
                        foreach (string attachment in Attachments)
                        {
                            message.Attachments.AddFileAttachment(attachment);
                        }
                    }

                    message.SendAndSaveCopy();

                    message = null;


                    i++;

                    results.Add("Success");
                }

                return results;

            }
            catch (Exception ex)
            {

                results.Add("Failed: " + ex.ToString());
                throw;
            }

        }

        public string CreateSubfolder(string EmailAddress, string EmailPassword, string NewFolderName)
        {
            try
            {

                //Establishing connection to the Exchange server
                Connection conn = new Connection();
                ExchangeService service = conn.InitializeConnection(EmailAddress, EmailPassword);

                Folder folder = new Folder(service);
                folder.DisplayName = NewFolderName;
                folder.Save(WellKnownFolderName.Inbox);
                return "Success";

            }
            catch (Exception ex)
            {

                return "Failed: " + ex.ToString();
            }

        }

        public string DeleteSubfolder(string EmailAddress, string EmailPassword, string FolderName)
        {
            try
            {


                //Establishing connection to the Exchange server
                Connection conn = new Connection();
                ExchangeService service = conn.InitializeConnection(EmailAddress, EmailPassword);
                FolderName = FolderName.ToLower();
                Folder folder = null;

                FindFoldersResults Subfolders = service.FindFolders(WellKnownFolderName.Inbox, new FolderView(int.MaxValue));

                foreach (var subfolder in Subfolders.Folders)
                {
                    if (FolderName == subfolder.DisplayName.ToLower())
                    {
                        FolderId deleteFolderID = subfolder.Id;
                        folder = Folder.Bind(service, deleteFolderID);
                        folder.Delete(DeleteMode.HardDelete);
                        break;
                    }

                }
                if (folder is null)
                {
                    return "Failed";
                }

                return "Success";
            }
            catch (Exception ex)
            {

                return "Failed: " + ex.ToString();
            }
        }


        public List<string> LoopThroughMailbox(string EmailAddress, string EmailPassword, string FolderName, string EmailType, string EmailLogPath, string DateOfReceipt, string[] EmailFrom = null, string[] EmailSubject = null, string[] EmailAttachment = null, string[] EmailBody = null, string[] EmailCC = null, string[] EmailBCC = null, string AttachmentSavePath = null, string MarkEmailAs = "", string DownloadAttachment = "", string DownloadEmail = "", string MoveEmailToFolder = "", string RetainOriginalAttName = "true", string MappedMailboxAddress = "", Int32 ItemsAmount = 10)
        {
            List<string> results = new List<string>();
            try
            {

                DateTime emailreceived = new DateTime();

                string senderAddress = "";
                string emailSubject = "";
                string emailBody = "";
                string emailAttachmentLocation = "";
                string datereceived = "";
                string emailAttachments = "";
                int emailID = 0;
                FolderId DestinationFolder = null;
                string log = "";

                //Reformatting input variables to universal lower caps
                DownloadEmail = DownloadEmail.ToLower();
                FolderName = FolderName.ToLower();
                EmailType = EmailType.ToLower();
                MarkEmailAs = MarkEmailAs.ToLower();
                DownloadAttachment = DownloadAttachment.ToLower();
                MoveEmailToFolder = MoveEmailToFolder.ToLower();
                RetainOriginalAttName = RetainOriginalAttName.ToLower();
                //Establishing connection to the Exchange server
                Connection conn = new Connection();
                ExchangeService service = conn.InitializeConnection(EmailAddress, EmailPassword);
                ExchangeCommands commands = new ExchangeCommands();

                //Creating filter for unread/read or all
                if (EmailType == "")
                {
                    results.Add("Incorrect filter. Allowed only: read/unread/all");
                    return results;

                }
                SearchFilter.IsEqualTo sf = commands.CreateFilter(EmailType, null);


                //Choosing which folder to connect to

                FindItemsResults<Item> EmailFolder = commands.SetFolder(null, FolderName, service, sf, EmailAddress, MappedMailboxAddress, ItemsAmount);
                if (EmailFolder is null)
                {
                    results.Add("Unable to find subfolder");
                    return results;
                }

                //reinitializing subfolder if user requested email to be moved
                if (MoveEmailToFolder != "")
                {
                    DestinationFolder = commands.GetSubfolerID(DestinationFolder, MoveEmailToFolder, service);

                    if (DestinationFolder is null)
                    {
                        results.Add("Unable to find destination folder to move the emailr");
                        return results;
                    }
                }

                log = "Email Sender" + "," + "Email Subject" + "," + "Attachment Save Path" + "," + "Email Attachments" + "," + "Email File Name" + "," + "Date of Email Receipt";

                File.AppendAllText(EmailLogPath, log + Environment.NewLine);

                //Process each item.
                emailID = 0;
                foreach (Item email in EmailFolder.Items)
                {

                    emailAttachments = "";
                    EmailMessage message = EmailMessage.Bind(service, email.Id, new PropertySet(ItemSchema.Attachments));

                    if (email is EmailMessage)
                    {


                        string emailfilename = "";

                        message.Load();

                        if (DateOfReceipt != "")
                        {
                            emailreceived = Convert.ToDateTime(DateOfReceipt);
                            if (emailreceived.ToShortTimeString() != message.DateTimeReceived.ToShortTimeString() || emailreceived.ToShortDateString() != message.DateTimeReceived.ToShortDateString())
                            {
                                continue;
                            }
                        }
                        if (DateOfReceipt == message.DateTimeReceived.ToString() && DateOfReceipt != "")
                        {
                            continue;
                        }
                        if ((EmailFrom.Any(message.Sender.Address.Contains) == false) && EmailFrom[0] != "")
                        {
                            continue;
                        }

                        if (message.Subject != null)
                        {
                            if ((EmailSubject.Any(message.Subject.Contains) == false) && EmailSubject[0] != "")
                            {
                                continue;
                            }
                        }
                        if (message.Body.Text != null)
                        {


                            if ((EmailBody.Any(message.Body.Text.Contains) == false) && EmailBody[0] != "")
                            {
                                continue;
                            }
                        }

                        if (DownloadAttachment == "true" || DownloadEmail == "true")
                        {

                            emailID++;


                            if (Directory.Exists(AttachmentSavePath + emailID) == false)
                            {
                                Directory.CreateDirectory(AttachmentSavePath + emailID);
                            }


                            if (DownloadAttachment == "true")
                            {
                                emailAttachments = commands.DownloadAttachments(message, AttachmentSavePath, EmailAttachment, RetainOriginalAttName, emailAttachmentLocation, emailID);
                            }

                            if (DownloadEmail == "true")
                            {
                                commands.DownloadEmail(email, message, AttachmentSavePath, emailID);
                            }
                        }

                        senderAddress = message.Sender.Address;

                        if (message.Subject is null)
                        {
                            emailSubject = "NoSubject";
                        }
                        else
                        {
                            emailSubject = message.Subject;
                        }
                        

                        emailBody = message.Body.Text;
                        datereceived = message.DateTimeReceived.ToString();

                        log = senderAddress + "," + emailSubject.Replace(",", "").ToString() + "," + AttachmentSavePath + emailID + @"\" + "," + emailAttachments.Replace(",", "").ToString() + "," + emailfilename + ".eml" + "," + datereceived;

                        File.AppendAllText(EmailLogPath, log + Environment.NewLine);
                        results.Add(senderAddress + "#|#" + emailSubject + "#|#" + AttachmentSavePath + emailID + @"\" + "#|#" + emailAttachments + "#|#" + emailfilename + ".eml" + "#|#" + datereceived);

                        commands.ChangeEmailStatus(message, MarkEmailAs);

                        if (MoveEmailToFolder != "")
                        {
                            commands.MoveEmail(email, DestinationFolder);
                        }

                    }
                }
                if (senderAddress == "")
                {
                    results.Add("No email available");

                }
                return results;

            }
            catch (Exception ex)
            {

                results.Add("Failed: " + ex.ToString());
                return results;
            }
        }
        public string GetOneEmail(string EmailAddress, string EmailPassword, string FolderName, string EmailType, string EmailLogPath, string DateOfReceipt, string[] EmailFrom = null, string[] EmailSubject = null, string[] EmailAttachment = null, string[] EmailBody = null, string[] EmailCC = null, string[] EmailBCC = null, string AttachmentSavePath = null, string MarkEmailAs = "", string DownloadAttachment = "", string DownloadEmail = "", string MoveEmailToFolder = "", string RetainOriginalAttName = "true", string MappedMailboxAddress = "")
        {
            string results = "";
            try
            {

                DateTime emailreceived = new DateTime();

                string senderAddress = "";
                string emailSubject = "";
                string emailBody = "";
                string emailAttachmentLocation = "";
                string datereceived = "";
                string emailAttachments = "";
                FolderId DestinationFolder = null;
                string log = "";

                //Reformatting input variables to universal lower caps
                DownloadEmail = DownloadEmail.ToLower();
                FolderName = FolderName.ToLower();
                EmailType = EmailType.ToLower();
                MarkEmailAs = MarkEmailAs.ToLower();
                DownloadAttachment = DownloadAttachment.ToLower();
                MoveEmailToFolder = MoveEmailToFolder.ToLower();
                RetainOriginalAttName = RetainOriginalAttName.ToLower();

                //Establishing connection to the Exchange server
                Connection conn = new Connection();
                ExchangeService service = conn.InitializeConnection(EmailAddress, EmailPassword);

                ExchangeCommands commands = new ExchangeCommands();



                //Creating filter for unread/read or all
                if (EmailType == "")
                {
                    return "Incorrect filter. Allowed only: read/unread/all";
                }
                SearchFilter.IsEqualTo sf = commands.CreateFilter(EmailType, null);


                //Choosing which folder to connect to

                FindItemsResults<Item> EmailFolder = commands.SetFolder(null, FolderName, service, sf, EmailAddress, MappedMailboxAddress);
                if (EmailFolder is null)
                {
                    return "Unable to find subfolder";
                }

                //reinitializing subfolder if user requested email to be moved
                if (MoveEmailToFolder != "")
                {
                    DestinationFolder = commands.GetSubfolerID(DestinationFolder, MoveEmailToFolder, service);

                    if (DestinationFolder is null)
                    {
                        return "Unable to find destination folder to move the email";
                    }
                }

                log = "Email Sender" + "," + "Email Subject" + "," + "Attachment Save Path" + "," + "Email Attachments" + "," + "Email File Name" + "," + "Date of Email Receipt";
                File.AppendAllText(EmailLogPath, log + Environment.NewLine);

                //Process each item.

                foreach (Item email in EmailFolder.Items)
                {

                    emailAttachments = "";
                    EmailMessage message = EmailMessage.Bind(service, email.Id, new PropertySet(ItemSchema.Attachments));

                    if (email is EmailMessage)
                    {

                        string emailfilename = "";

                        message.Load();

                        if (DateOfReceipt != "")
                        {
                            emailreceived = Convert.ToDateTime(DateOfReceipt);
                            if (emailreceived.ToShortTimeString() != message.DateTimeReceived.ToShortTimeString() || emailreceived.ToShortDateString() != message.DateTimeReceived.ToShortDateString())
                            {
                                continue;
                            }
                        }

                        if ((EmailFrom.Any(message.Sender.Address.Contains) == false) && EmailFrom[0] != "")
                        {
                            continue;
                        }

                        if (message.Subject != null)
                        {
                            if ((EmailSubject.Any(message.Subject.Contains) == false) && EmailSubject[0] != "")
                            {
                                continue;
                            }
                        }

                        if (message.Body.Text != null)
                        {


                            if ((EmailBody.Any(message.Body.Text.Contains) == false) && EmailBody[0] != "")
                            {
                                continue;
                            }
                        }

                        if (DownloadAttachment == "true" || DownloadEmail == "true")
                        {


                            if (Directory.Exists(AttachmentSavePath) == false)
                            {
                                Directory.CreateDirectory(AttachmentSavePath);
                            }


                            if (DownloadAttachment == "true")
                            {
                                emailAttachments = commands.DownloadAttachments(message, AttachmentSavePath, EmailAttachment, RetainOriginalAttName, emailAttachmentLocation);
                            }

                            if (DownloadEmail == "true")
                            {
                                commands.DownloadEmail(email, message, AttachmentSavePath);
                            }
                        }

                        senderAddress = message.Sender.Address;

                        if (message.Subject is null)
                        {
                            emailSubject = "NoSubject";
                        }
                        else
                        {
                            emailSubject = message.Subject;
                        }

                        emailBody = message.Body.Text;
                        datereceived = message.DateTimeReceived.ToString();

                        log = senderAddress + "," + emailSubject.Replace(",", "").ToString() + "," + AttachmentSavePath + "," + emailAttachments.Replace(",", "").ToString() + "," + emailfilename + ".eml" + "," + datereceived;
                        File.AppendAllText(EmailLogPath, log + Environment.NewLine);

                        results = senderAddress + "#|#" + emailSubject + "#|#" + AttachmentSavePath + "#|#" + emailAttachments + "#|#" + emailfilename + ".eml" + "#|#" + datereceived;

                        commands.ChangeEmailStatus(message, MarkEmailAs);

                        if (MoveEmailToFolder != "")
                        {
                            commands.MoveEmail(email, DestinationFolder);
                        }

                        break;
                    }
                }
                if (senderAddress == "")
                {
                    results = "No email available";
                }
                return results;

            }
            catch (Exception ex)
            {
                return "Failed: " + ex.ToString();
            }
        }

    }
}
