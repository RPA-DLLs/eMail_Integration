using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Exchange.WebServices.Data;
using System.Text.RegularExpressions;
using Microsoft.Exchange.WebServices.Autodiscover;
using System.Net;

namespace Testing
{
    class Program
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


                string emailfilename = rgx.Replace(message.Subject.ToString(), "");

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

            public void MoveEmail (Item email, FolderId DestinationFolder)
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
                            if (attachment is ItemAttachment)
                            {
                                ItemAttachment fileAttachment = attachment as ItemAttachment;
                                extension = ".eml";
                                filename = fileAttachment.Name.ToString();
                                Regex rgx = new Regex("[^a-zA-Z0-9 ]");


                                filename = rgx.Replace(filename, "");
                                filename = filename + extension;
                            }
                            else
                            {
                                FileAttachment fileAttachment = attachment as FileAttachment;
                                filename = fileAttachment.Name.ToString();
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

                    int i = 0;
                    //Establishing connection to the Exchange server
                    Connection conn = new Connection();
                    ExchangeService service = conn.InitializeConnection(EmailAddress, EmailPassword);

                    foreach (string item in EmailTo)
                    {




                        EmailMessage message = new EmailMessage(service);

                        string EmailToItem = item.Replace(" ", "");
                        string EmailCCItem = item.Replace(" ", "");
                        string EmailBCCItem = item.Replace(" ", "");

                        List<string> Recepients = EmailToItem.Split(';').ToList();
                        List<string> RecepientsCC = EmailCCItem.Split(';').ToList();
                        List<string> RecepientsBCC = EmailBCCItem.Split(';').ToList();

                        List<string> Attachments = EmailAttachments[i].Split('|').ToList();



                        string IsReadReceiptRequestedItem = IsReadReceiptRequested[i].ToLower();
                        string IsDeliveryReceiptRequestedItem = IsDeliveryReceiptRequested[i].ToLower();
                        HighEmailImportance[i] = HighEmailImportance[i].ToLower();

                        //Message properties
                        message.ToRecipients.AddRange(Recepients);
                        message.CcRecipients.AddRange(RecepientsCC);
                        message.BccRecipients.AddRange(RecepientsBCC);
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

                    FindItemsResults<Item> EmailFolder = commands.SetFolder(null, FolderName, service, sf, EmailAddress, MappedMailboxAddress,ItemsAmount);
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


                            if ((EmailSubject.Any(message.Subject.Contains) == false) && EmailSubject[0] != "")
                            {
                                continue;
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
                                    commands.DownloadEmail(email,message, AttachmentSavePath,emailID);
                                }
                            }

                            senderAddress = message.Sender.Address;
                            emailSubject = message.Subject;
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

                    FindItemsResults<Item>  EmailFolder = commands.SetFolder(null,FolderName,service,sf,EmailAddress, MappedMailboxAddress);
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


                            if ((EmailSubject.Any(message.Subject.Contains) == false) && EmailSubject[0] != "")
                            {
                                continue;
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
                            emailSubject = message.Subject;
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
            static void Main(string[] args)
            {
                //string[] EmailTo = new string[] { "lukasz.blaszczyk@sbdinc.com; lukasz.blaszczyk@sbdinc.com", "lukasz.blaszczyk@sbdinc.com" };
                //string[] EmailCC = new string[] { "lukasz.blaszczyk@sbdinc.com; lukasz.blaszczyk@sbdinc.com", "lukasz.blaszczyk@sbdinc.com" };
                //string[] EmailBCC = new string[] { "lukasz.blaszczyk@sbdinc.com; lukasz.blaszczyk@sbdinc.com", "lukasz.blaszczyk@sbdinc.com" };
                //string[] Subject = new string[] { "MBOT test exchange", "Another test" };
                //string[] Body = new string[] { @"<html><head></head><body><img width=245 height=97 id=""1"" src=""cid:ARpic.jpg""></body></html>" + "<html xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:w=\"urn:schemas-microsoft-com:office:word\" xmlns:m=\"http://schemas.microsoft.com/office/2004/12/omml\" xmlns=\"http://www.w3.org/TR/REC-html40\">\r\n<head>\r\n<meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\">\r\n<meta name=\"Generator\" content=\"Microsoft Word 15 (filtered medium)\">\r\n<style><!--\r\n/* Font Definitions */\r\n@font-face\r\n\t{font-family:\"Cambria Math\";\r\n\tpanose-1:2 4 5 3 5 4 6 3 2 4;}\r\n@font-face\r\n\t{font-family:Calibri;\r\n\tpanose-1:2 15 5 2 2 2 4 3 2 4;}\r\n/* Style Definitions */\r\np.MsoNormal, li.MsoNormal, div.MsoNormal\r\n\t{margin:0in;\r\n\tmargin-bottom:.0001pt;\r\n\tfont-size:11.0pt;\r\n\tfont-family:\"Calibri\",sans-serif;}\r\na:link, span.MsoHyperlink\r\n\t{mso-style-priority:99;\r\n\tcolor:#0563C1;\r\n\ttext-decoration:underline;}\r\na:visited, span.MsoHyperlinkFollowed\r\n\t{mso-style-priority:99;\r\n\tcolor:#954F72;\r\n\ttext-decoration:underline;}\r\nspan.EmailStyle17\r\n\t{mso-style-type:personal-compose;\r\n\tfont-family:\"Calibri\",sans-serif;\r\n\tcolor:windowtext;}\r\n.MsoChpDefault\r\n\t{mso-style-type:export-only;\r\n\tfont-family:\"Calibri\",sans-serif;}\r\n@page WordSection1\r\n\t{size:8.5in 11.0in;\r\n\tmargin:1.0in 1.0in 1.0in 1.0in;}\r\ndiv.WordSection1\r\n\t{page:WordSection1;}\r\n--></style><!--[if gte mso 9]><xml>\r\n<o:shapedefaults v:ext=\"edit\" spidmax=\"1026\" />\r\n</xml><![endif]--><!--[if gte mso 9]><xml>\r\n<o:shapelayout v:ext=\"edit\">\r\n<o:idmap v:ext=\"edit\" data=\"1\" />\r\n</o:shapelayout></xml><![endif]-->\r\n</head>\r\n<body lang=\"EN-US\" link=\"#0563C1\" vlink=\"#954F72\">\r\n<div class=\"WordSection1\">\r\n<p class=\"MsoNormal\">Hello <b>SANTA MARIA SEEDS INC,</b><o:p></o:p></p>\r\n<p class=\"MsoNormal\"><o:p>&nbsp;</o:p></p>\r\n<p class=\"MsoNormal\">This is a courtesy reminder that your Account# <b>1371004622</b> has a balance of<b> $3,320.00</b>.&nbsp; Please take some time to browse this email and help us resolve this at your earliest convenience.&nbsp; For your convenience, we have added\r\n all the tools in one place to quickly and easily resolve any issues.&nbsp; <o:p></o:p></p>\r\n<p class=\"MsoNormal\"><o:p>&nbsp;</o:p></p>\r\n<p class=\"MsoNormal\"><o:p>&nbsp;</o:p></p>\r\n<p class=\"MsoNormal\"><b><span style=\"font-size:14.0pt\">MAKE A PAYMENT<o:p></o:p></span></b></p>\r\n<p class=\"MsoNormal\"><b><span style=\"font-size:14.0pt\"><o:p>&nbsp;</o:p></span></b></p>\r\n<p class=\"MsoNormal\">To set up <b>AutoPay</b> and/or to make a one-time payment please visit our\r\n<span class=\"MsoHyperlink\"><a href=\"https://secure.billtrust.com/stanleycss/ig/signin\">Payment Portal</a></span>.<o:p></o:p></p>\r\n<p class=\"MsoNormal\"><o:p>&nbsp;</o:p></p>\r\n<p class=\"MsoNormal\">Or,<o:p></o:p></p>\r\n<p class=\"MsoNormal\"><o:p>&nbsp;</o:p></p>\r\n<p class=\"MsoNormal\">Find the attached remittance information <b><u>Below.<o:p></o:p></u></b></p>\r\n<p class=\"MsoNormal\"><b><u><o:p><span style=\"text-decoration:none\">&nbsp;</span></o:p></u></b></p>\r\n<p class=\"MsoNormal\"><b><u><o:p><span style=\"text-decoration:none\">&nbsp;</span></o:p></u></b></p>\r\n<p class=\"MsoNormal\"><b><span style=\"font-size:14.0pt\">ACCOUNT HELP<o:p></o:p></span></b></p>\r\n<p class=\"MsoNormal\"><o:p>&nbsp;</o:p></p>\r\n<p class=\"MsoNormal\"><b>Billing questions<o:p></o:p></b></p>\r\n<p class=\"MsoNormal\">Please reach out to your designated accounting specialist at\r\n<b><a href=\"mailto:Santiago.Quiroz@sbdinc.com\">Santiago.Quiroz@sbdinc.com</a></b> or\r\n<b>317-436-9866</b>.<o:p></o:p></p>\r\n<p class=\"MsoNormal\"><o:p>&nbsp;</o:p></p>\r\n<p class=\"MsoNormal\"><b>Customer Service<o:p></o:p></b></p>\r\n<p class=\"MsoNormal\">Please visit our <span class=\"MsoHyperlink\"><a href=\"https://www.stanleysecuritysolutions.com/customerservice\">Customer Service Portal</a></span> or call (855)578-2653 to speak to a representative.<o:p></o:p></p>\r\n<p class=\"MsoNormal\"><o:p>&nbsp;</o:p></p>\r\n<p class=\"MsoNormal\"><o:p>&nbsp;</o:p></p>\r\n<p class=\"MsoNormal\"><o:p>&nbsp;</o:p></p>\r\n<p class=\"MsoNormal\"><b>Thank you,<o:p></o:p></b></p>\r\n<p class=\"MsoNormal\"><b><o:p>&nbsp;</o:p></b></p>\r\n<p class=\"MsoNormal\">SSS Collections Team<o:p></o:p></p>\r\n<p class=\"MsoNormal\"><o:p>&nbsp;</o:p></p>\r\n<p class=\"MsoNormal\"><o:p>&nbsp;</o:p></p>\r\n<p class=\"MsoNormal\"><o:p>&nbsp;</o:p></p>\r\n<p class=\"MsoNormal\"><o:p>&nbsp;</o:p></p>\r\n<p class=\"MsoNormal\" align=\"center\" style=\"margin-bottom:10.0pt;text-align:center;line-height:115%\">\r\n<b><span style=\"font-size:22.0pt;line-height:115%;font-family:&quot;Times New Roman&quot;,serif\">**IMPORTANT NOTICE**<o:p></o:p></span></b></p>\r\n<p class=\"MsoNormal\" align=\"center\" style=\"text-align:center;text-autospace:none\">\r\n<b><span style=\"font-size:12.0pt;font-family:&quot;Arial&quot;,sans-serif;color:black\">To ensure proper credit to your account, all Stanley CSS payments must be mailed to the following address:<o:p></o:p></span></b></p>\r\n<p class=\"MsoNormal\" align=\"center\" style=\"text-align:center;text-autospace:none\">\r\n<b><span style=\"font-size:12.0pt;font-family:&quot;Arial&quot;,sans-serif;color:black\"><o:p>&nbsp;</o:p></span></b></p>\r\n<p class=\"MsoNormal\" align=\"center\" style=\"text-align:center;text-autospace:none\">\r\n<b><span style=\"font-size:14.0pt;font-family:&quot;Arial&quot;,sans-serif;color:#0070C0\">STANLEY CONVERGENT SECURITY SOLUTIONS<o:p></o:p></span></b></p>\r\n<p class=\"MsoNormal\" align=\"center\" style=\"text-align:center;text-autospace:none\">\r\n<b><span style=\"font-size:14.0pt;font-family:&quot;Arial&quot;,sans-serif;color:#0070C0\">DEPT CH 10651<o:p></o:p></span></b></p>\r\n<p class=\"MsoNormal\" align=\"center\" style=\"text-align:center;text-autospace:none\">\r\n<b><span style=\"font-size:14.0pt;font-family:&quot;Arial&quot;,sans-serif;color:#0070C0\">PALATINE, IL 60055-0651<o:p></o:p></span></b></p>\r\n<p class=\"MsoNormal\" align=\"center\" style=\"text-align:center;text-autospace:none\">\r\n<b><span style=\"font-size:14.0pt;font-family:&quot;Arial&quot;,sans-serif;color:red\"><o:p>&nbsp;</o:p></span></b></p>\r\n<p class=\"MsoNormal\" align=\"center\" style=\"text-align:center;text-autospace:none\">\r\n<b><span style=\"font-size:12.0pt;font-family:&quot;Arial&quot;,sans-serif;color:black\">Failure to send your payments to this address will result in significant delays in posting your payment. This address is included on the remittance coupon of your invoice(s). We strongly\r\n recommend returning the coupon with your payment to avoid complications that can result when remittance advice is not included.<o:p></o:p></span></b></p>\r\n<p class=\"MsoNormal\" align=\"center\" style=\"text-align:center;text-autospace:none\">\r\n<b><span style=\"font-size:12.0pt;font-family:&quot;Arial&quot;,sans-serif;color:black\"><o:p>&nbsp;</o:p></span></b></p>\r\n<p class=\"MsoNormal\" align=\"center\" style=\"text-align:center;text-autospace:none\">\r\n<b><span style=\"font-size:12.0pt;font-family:&quot;Arial&quot;,sans-serif;color:black\"><o:p>&nbsp;</o:p></span></b></p>\r\n<p class=\"MsoNormal\" align=\"center\" style=\"text-align:center;text-autospace:none\">\r\n<b><span style=\"font-size:12.0pt;font-family:&quot;Arial&quot;,sans-serif;color:black\">Do not send correspondence to the payment address. Refer to the top of your&nbsp; invoice(s) for the proper correspondence address.<o:p></o:p></span></b></p>\r\n<p class=\"MsoNormal\" align=\"center\" style=\"text-align:center;text-autospace:none\">\r\n<b><span style=\"font-size:12.0pt;font-family:&quot;Arial&quot;,sans-serif;color:black\"><o:p>&nbsp;</o:p></span></b></p>\r\n<p class=\"MsoNormal\" align=\"center\" style=\"text-align:center;text-autospace:none\">\r\n<b><span style=\"font-size:12.0pt;font-family:&quot;Arial&quot;,sans-serif;color:black\">If you currently remit your payment via electronic funds transfer, please be sure that send payments are to the account noted below:<o:p></o:p></span></b></p>\r\n<p class=\"MsoNormal\" align=\"center\" style=\"text-align:center;text-autospace:none\">\r\n<b><span style=\"font-size:12.0pt;font-family:&quot;Arial&quot;,sans-serif;color:black\"><o:p>&nbsp;</o:p></span></b></p>\r\n<p class=\"MsoNormal\" align=\"center\" style=\"text-align:center;text-autospace:none\">\r\n<b><span style=\"font-size:12.0pt;font-family:&quot;Arial&quot;,sans-serif;color:black\"><o:p>&nbsp;</o:p></span></b></p>\r\n<p class=\"MsoNormal\" align=\"center\" style=\"text-align:center;text-autospace:none\">\r\n<b><span style=\"font-size:14.0pt;font-family:&quot;Arial&quot;,sans-serif;color:#0070C0\">BANK NAME: MELLON BANK<o:p></o:p></span></b></p>\r\n<p class=\"MsoNormal\" align=\"center\" style=\"text-align:center;text-autospace:none\">\r\n<b><span style=\"font-size:14.0pt;font-family:&quot;Arial&quot;,sans-serif;color:#0070C0\">ACCOUNT NUMBER: 0127327<o:p></o:p></span></b></p>\r\n<p class=\"MsoNormal\" align=\"center\" style=\"text-align:center;text-autospace:none\">\r\n<b><span style=\"font-size:14.0pt;font-family:&quot;Arial&quot;,sans-serif;color:#0070C0\">ROUTING NUMBER: 043000261<o:p></o:p></span></b></p>\r\n<p class=\"MsoNormal\" style=\"text-autospace:none\"><b><span style=\"font-size:14.0pt;font-family:&quot;Arial&quot;,sans-serif\"><o:p>&nbsp;</o:p></span></b></p>\r\n<p class=\"MsoNormal\" style=\"text-autospace:none\"><b><span style=\"font-size:14.0pt;font-family:&quot;Arial&quot;,sans-serif\"><o:p>&nbsp;</o:p></span></b></p>\r\n<p class=\"MsoNormal\" align=\"center\" style=\"text-align:center;text-autospace:none\">\r\n<b><span style=\"font-size:12.0pt;font-family:&quot;Arial&quot;,sans-serif;color:black\">Always include your invoice number in your file transmission to ensure proper posting of your payment(s). Please send remittance information to the following address to ensure prompt\r\n posting of payments:<o:p></o:p></span></b></p>\r\n<p class=\"MsoNormal\" align=\"center\" style=\"text-align:center;text-autospace:none\">\r\n<b><span style=\"font-size:12.0pt;font-family:&quot;Arial&quot;,sans-serif;color:black\"><o:p>&nbsp;</o:p></span></b></p>\r\n<p class=\"MsoNormal\" align=\"center\" style=\"text-align:center;text-autospace:none\">\r\n<b><u><span style=\"font-size:12.0pt;font-family:&quot;Arial&quot;,sans-serif;color:blue\"><a href=\"mailto:CSS-ACHRemittance@sbdinc.com\"><span style=\"color:blue\">CSS-ACHRemittance@sbdinc.com</span></a></span></u></b><b><span style=\"font-size:12.0pt;font-family:&quot;Arial&quot;,sans-serif;color:#0070C0\"><o:p></o:p></span></b></p>\r\n<p class=\"MsoNormal\" align=\"center\" style=\"text-align:center;text-autospace:none\">\r\n<b><span style=\"font-size:12.0pt;font-family:&quot;Arial&quot;,sans-serif;color:#0070C0\"><o:p>&nbsp;</o:p></span></b></p>\r\n<p class=\"MsoNormal\" align=\"center\" style=\"text-align:center;text-autospace:none\">\r\n<b><span style=\"font-size:12.0pt;font-family:&quot;Arial&quot;,sans-serif\">This address is case sensitive and must appear as indicated or your message will not be received.<o:p></o:p></span></b></p>\r\n<p class=\"MsoNormal\" style=\"text-autospace:none\"><b><span style=\"font-size:12.0pt;font-family:&quot;Arial&quot;,sans-serif;color:black\"><o:p>&nbsp;</o:p></span></b></p>\r\n<p class=\"MsoNormal\" align=\"center\" style=\"text-align:center;text-autospace:none\">\r\n<b><span style=\"font-size:12.0pt;font-family:&quot;Arial&quot;,sans-serif;color:black\">Should you need to send your payment by UPS or Fed Ex, the payment should be sent to:<o:p></o:p></span></b></p>\r\n<p class=\"MsoNormal\" style=\"text-autospace:none\"><b><span style=\"font-size:12.0pt;font-family:&quot;Arial&quot;,sans-serif;color:black\"><o:p>&nbsp;</o:p></span></b></p>\r\n<p class=\"MsoNormal\" align=\"center\" style=\"text-align:center;text-autospace:none\">\r\n<b><span style=\"font-size:14.0pt;font-family:&quot;Arial&quot;,sans-serif;color:black\">STANLEY CONVERGENT SECURITY SOLUTIONS<o:p></o:p></span></b></p>\r\n<p class=\"MsoNormal\" align=\"center\" style=\"text-align:center;text-autospace:none\">\r\n<b><span style=\"font-size:14.0pt;font-family:&quot;Arial&quot;,sans-serif;color:black\">5505 N CUMBERLAND AVENUE, SUITE 307<o:p></o:p></span></b></p>\r\n<p class=\"MsoNormal\" align=\"center\" style=\"text-align:center;text-autospace:none\">\r\n<b><span style=\"font-size:14.0pt;font-family:&quot;Arial&quot;,sans-serif;color:black\">DEPT CH-10651<o:p></o:p></span></b></p>\r\n<p class=\"MsoNormal\" align=\"center\" style=\"text-align:center;text-autospace:none\">\r\n<b><span style=\"font-size:14.0pt;font-family:&quot;Arial&quot;,sans-serif;color:black\">CHICAGO, IL 60656-1471<o:p></o:p></span></b></p>\r\n<p class=\"MsoNormal\" align=\"center\" style=\"text-align:center;text-autospace:none\">\r\n<b><span style=\"font-size:12.0pt;font-family:&quot;Arial&quot;,sans-serif;color:black\"><o:p>&nbsp;</o:p></span></b></p>\r\n<p class=\"MsoNormal\"><o:p>&nbsp;</o:p></p>\r\n<p class=\"MsoNormal\"><o:p>&nbsp;</o:p></p>\r\n<p class=\"MsoNormal\"><o:p>&nbsp;</o:p></p>\r\n<p class=\"MsoNormal\"><o:p>&nbsp;</o:p></p>\r\n</div>\r\n</body>\r\n</html>\r\n" };
                //string[] IsReadReceiptRequested = new string[] { "false", "false" };
                //string[] IsDeliveryReceiptRequested = new string[] { "false", "false" };
                //string[] EmailAttachments = new string[] { "" };
                //string[] EmailEmbedImage = new string[] { @"C:\Users\LXB0906\Desktop\ARpic.jpg" };
                //Integration test = new Integration();
                //test.SendEmail("lukasz.blaszczyk@sbdinc.com", "Beta321%", EmailTo, EmailCC, EmailBCC, Subject, Body, IsReadReceiptRequested, IsDeliveryReceiptRequested, EmailAttachments, EmailEmbedImage);

                string[] EmailFrom = new string[] { "" };
                string[] EmailCC = new string[] { "" };
                string[] EmailBCC = new string[] { };
                string[] Subject = new string[] { "" };
                string[] Body = new string[] { "" };

                string[] EmailAttachments = new string[] { "" };
                Integration test = new Integration();
                //test.GetOneEmail("FNC-RPA-CRP-PTP4014@sbdinc.com", "w!mcE3Eclg", "inbox", "all", "C:\\Users\\LXB0906\\Desktop\\test.csv", "", EmailFrom, Subject, EmailAttachments, Body, EmailCC, EmailBCC, @"C:\Users\LXB0906\Desktop\exchange\", "true", "true", "true", "", "false", "");
                test.LoopThroughMailbox("lukasz.blaszczyk@sbdinc.com", "Beta321^", "inbox", "all", "C:\\Users\\LXB0906\\Desktop\\test.csv", "", EmailFrom, Subject, EmailAttachments, Body, EmailCC, EmailBCC, @"C:\Users\LXB0906\Desktop\exchange\fakefakefakefakefakefakefakefakefakefakefakefakefakefakefakefakefakefakefakefakefakefakefakefakefakefake\", "unread", "true", "true", "OOB", "true", "",4);


            }
        }

        
    }
}