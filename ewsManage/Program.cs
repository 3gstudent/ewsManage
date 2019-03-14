using System;
using Microsoft.Exchange.WebServices.Data;
using System.Net;
using System.IO;
using System.Text;

namespace ewsManage
{
    class Program
    {
        public static string UserString;
        public static string PwdString;
        public static string ewsPathString;
        public static string AutodiscoverUrlString;
        public static string FolderString;
        public static string IdString;
        public static string SearchString;
        public static void ShowUsage()
        {
            string Usage = @"
Author: 3gstudent
Usage:
    ews_tool.exe -CerValidation <Yes or No> -ExchangeVersion <Version> -u <Username> -p <Password> -ewsPath <EWS url> -Mode <command> -<Command> <command>
    ews_tool.exe -CerValidation <Yes or No> -ExchangeVersion <Version> -u <Username> -p <Password> -AutodiscoverUrl <Autodiscover url> -Mode <command> -<Command> <command>
    ews_tool.exe -CerValidation <Yes or No> -ExchangeVersion <Version> -use the default credentials -ewsPath <EWS url> -Mode <command> -<Command> <command>
-CerValidation:
    Yes/No
-ExchangeVersion:
    Exchange2007_SP1/Exchange2010/Exchange2010_SP1/Exchange2010_SP2/Exchange2013/Exchange2013_SP1
-ewsPath/-AutodiscoverUrl:
    You should choose one 
-Mode:
    ListMail/ListUnreadMail/ListFolder -Foler:
        Inbox/Drafts/SentItems/DeletedItems/Outbox/JunkEmail        
    SaveAttachment/ClearAllAttachment/DeleteMail/ViewMail -Id
        You can get the Id by using ListMail/ListUnreadMail
    AddAttachment/DeleteAttachment -Id -AttachmentFile
        You need set 2 parameters
    ListMailofFolder/ListUnreadMailofFolder -Id
        You can get the Id by using ListFolder
    SearchMail -String
        search folder:Inbox/Drafts/SentItems/DeletedItems/Outbox/JunkEmail
        search location:Subject/Attachment name/MessageBody
    ReadXML -Path
        use EWS SOAP to send command

eg.
    ewsManage.exe -CerValidation Yes -ExchangeVersion Exchange2013_SP1 -u test1 -p test123! -ewsPath https://test.com/ews/Exchange.asmx -Mode ListUnreadMail -Folder Inbox
    ewsManage.exe -CerValidation No -ExchangeVersion Exchange2013_SP1 -use the default credentials -AutodiscoverUrl test1@test.com -Mode ListMail -Folder SentItems
    ewsManage.exe -CerValidation No -ExchangeVersion Exchange2013_SP1 -u test1 -p test123! -ewsPath https://test.com/ews/Exchange.asmx -Mode ListFolder -Folder Inbox
    ewsManage.exe -CerValidation No -ExchangeVersion Exchange2013_SP1 -u test1 -p test123! -ewsPath https://test.com/ews/Exchange.asmx -Mode ListMailofFolder -Id AAMaADFlMjRjMdM2LTgxZTUtNGRmZC05ZDQyLTMzNDFlMzBmZWY1NwAzAAAAAAAR9UOK286vT6HjUgukBQGmAQBHzR2O8KNmTcffGwlY0A76AAAAADfqAAA=
    ewsManage.exe -CerValidation No -ExchangeVersion Exchange2013_SP1 -u test1 -p test123! -ewsPath https://test.com/ews/Exchange.asmx -Mode ReadXML -Path ews.xml
";
            Console.WriteLine(Usage);
        }

        static void Main(string[] args)
        {
            if (args.Length < 14)
            {
                ShowUsage();
                Environment.Exit(0);
            }
            ExchangeService service = null;
            if ((args[0] != "-CerValidation"))
            {
                ShowUsage();
                Environment.Exit(0);
            }
            if ((args[1] == "No"))
            {
                ServicePointManager.ServerCertificateValidationCallback = (sender, certificate, chain, sslPolicyErrors) => { return true; };
            }
            Console.WriteLine("[+]CerValidation:{0}", args[1]);

            if ((args[2] != "-ExchangeVersion"))
            {
                ShowUsage();
                Environment.Exit(0);
            }
            if ((args[3] == "Exchange2007_SP1"))
            {
                service = new ExchangeService(ExchangeVersion.Exchange2007_SP1);
            }
            else if ((args[3] == "Exchange2010"))
            {
                service = new ExchangeService(ExchangeVersion.Exchange2010);
            }
            else if ((args[3] == "Exchange2010_SP1"))
            {
                service = new ExchangeService(ExchangeVersion.Exchange2010_SP1);
            }
            else if ((args[3] == "Exchange2010_SP2"))
            {
                service = new ExchangeService(ExchangeVersion.Exchange2010_SP2);
            }
            else if ((args[3] == "Exchange2013"))
            {
                service = new ExchangeService(ExchangeVersion.Exchange2013);
            }
            else if ((args[3] == "Exchange2013_SP1"))
            {
                service = new ExchangeService(ExchangeVersion.Exchange2013_SP1);
            }
            else
            {
                ShowUsage();
                Environment.Exit(0);
            }
            Console.WriteLine("[+]ExchangeVersion:{0}", args[3]);
            if (args[4] == "-u")
            {
                UserString = args[5];
                Console.WriteLine("[+]User:{0}", args[5]);
                if ((args[6] != "-p"))
                {
                    ShowUsage();
                    Environment.Exit(0);
                }
                PwdString = args[7];
                Console.WriteLine("[+]Password:{0}", args[7]);
                service.Credentials = new WebCredentials(UserString, PwdString);
            }
            else if (args[4] == "-use")
            {
                if (args[5] == "the")
                {
                    if (args[6] == "default")
                    {
                        if (args[7] == "credentials")
                        {
                            service.UseDefaultCredentials = true;
                            Console.WriteLine("[+]Use the default credentials");
                        }
                    }
                }
                else
                {
                    ShowUsage();
                    Environment.Exit(0);
                }
            }
            else
            {
                ShowUsage();
                Environment.Exit(0);
            }
            if ((args[8] == "-ewsPath"))
            {
                Console.WriteLine("[+]ewsPath:{0}", args[9]);
                ewsPathString = args[9];
                try
                {
                    service.Url = new Uri(ewsPathString);
                }
                catch (Exception e)
                {
                    Console.WriteLine("{0}", e.Message);
                    Environment.Exit(0);
                }
            }
            else if ((args[8] == "-AutodiscoverUrl"))
            {
                Console.WriteLine("[+]AutodiscoverUrl:{0}", args[9]);
                AutodiscoverUrlString = args[9];
                try
                {
                    service.AutodiscoverUrl(AutodiscoverUrlString, RedirectionUrlValidationCallback);
                    Console.WriteLine("[+]ewsPath:{0}", service.Url);
                }
                catch (Exception e)
                {
                    Console.WriteLine("{0}", e.Message);
                    Environment.Exit(0);
                }
            }
            else
            {
                ShowUsage();
                Environment.Exit(0);
            }
            if ((args[10] != "-Mode"))
            {
                ShowUsage();
                Environment.Exit(0);
            }

            if ((args[11] == "ListMail"))
            {
                Console.WriteLine("[+]Mode:{0}", args[11]);
                try
                {
                    ItemView view = new ItemView(int.MaxValue);
                    PropertySet itempropertyset = new PropertySet(BasePropertySet.FirstClassProperties);
                    itempropertyset.RequestedBodyType = BodyType.Text;
                    view.PropertySet = itempropertyset;
                    if ((args[12] != "-Folder"))
                    {
                        ShowUsage();
                        Environment.Exit(0);
                    }
                    FindItemsResults<Item> findResults = null;
                    FolderString = args[13];
                    Console.WriteLine("[+]Folder:{0}", args[13]);
                    if (args[13] == "Inbox")
                    {
                        findResults = service.FindItems(WellKnownFolderName.Inbox, view);
                    }
                    else if (args[13] == "Outbox")
                    {
                        findResults = service.FindItems(WellKnownFolderName.Outbox, view);
                    }
                    else if (args[13] == "DeletedItems")
                    {
                        findResults = service.FindItems(WellKnownFolderName.DeletedItems, view);
                    }
                    else if (args[13] == "Drafts")
                    {
                        findResults = service.FindItems(WellKnownFolderName.Drafts, view);
                    }
                    else if (args[13] == "SentItems")
                    {
                        findResults = service.FindItems(WellKnownFolderName.SentItems, view);
                    }
                    else if (args[13] == "JunkEmail")
                    {
                        findResults = service.FindItems(WellKnownFolderName.JunkEmail, view);
                    }
                    else
                    {
                        ShowUsage();
                        Environment.Exit(0);
                    }
                    foreach (Item item in findResults.Items)
                    {
                        Console.WriteLine("\r\n");
                        if (item.Subject != null)
                        {
                            Console.WriteLine("[*]Subject:{0}", item.Subject);
                        }
                        else
                        {
                            Console.WriteLine("[*]Subject:<null>");
                        }
                        Console.WriteLine("[*]HasAttachments:{0}", item.HasAttachments);
                        if (item.HasAttachments)
                        {
                            EmailMessage message = EmailMessage.Bind(service, item.Id, new PropertySet(ItemSchema.Attachments));
                            foreach (Attachment attachment in message.Attachments)
                            {
                                FileAttachment fileAttachment = attachment as FileAttachment;
                                fileAttachment.Load();
                                Console.WriteLine(" - Attachments:{0}", fileAttachment.Name);
                            }
                        }
                        Console.WriteLine("[*]ItemId:{0}", item.Id);
                        Console.WriteLine("[*]DateTimeCreated:{0}", item.DateTimeCreated);
                        Console.WriteLine("[*]DateTimeReceived:{0}", item.DateTimeReceived);
                        Console.WriteLine("[*]DateTimeSent:{0}", item.DateTimeSent);
                        Console.WriteLine("[*]DisplayCc:{0}", item.DisplayCc);
                        Console.WriteLine("[*]DisplayTo:{0}", item.DisplayTo);
                        Console.WriteLine("[*]InReplyTo:{0}", item.InReplyTo);
                        Console.WriteLine("[*]Size:{0}", item.Size);
                        item.Load(itempropertyset);
                        if (item.Body.ToString().Length > 100)
                        {
                            item.Body = item.Body.ToString().Substring(0, 100);
                            Console.WriteLine("[*]MessageBody(too big,only output 100):{0}", item.Body);
                        }
                        else
                        {
                            Console.WriteLine("[*]MessageBody:{0}", item.Body);
                        }
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine("[!]{0}", e.Message);
                    Environment.Exit(0);
                }
            }

            else if ((args[11] == "ListUnreadMail"))
            {
                Console.WriteLine("[+]Mode:{0}", args[11]);
                try
                {
                    SearchFilter sf = new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, false);
                    ItemView view = new ItemView(int.MaxValue);
                    PropertySet itempropertyset = new PropertySet(BasePropertySet.FirstClassProperties);
                    itempropertyset.RequestedBodyType = BodyType.Text;
                    view.PropertySet = itempropertyset;
                    if ((args[12] != "-Folder"))
                    {
                        ShowUsage();
                        Environment.Exit(0);
                    }
                    FindItemsResults<Item> findResults = null;
                    FolderString = args[13];
                    Console.WriteLine("[+]Folder:{0}", args[13]);
                    if (args[13] == "Inbox")
                    {
                        findResults = service.FindItems(WellKnownFolderName.Inbox, sf, view);
                    }
                    else if (args[13] == "Outbox")
                    {
                        findResults = service.FindItems(WellKnownFolderName.Outbox, sf, view);
                    }
                    else if (args[13] == "DeletedItems")
                    {
                        findResults = service.FindItems(WellKnownFolderName.DeletedItems, sf, view);
                    }
                    else if (args[13] == "Drafts")
                    {
                        findResults = service.FindItems(WellKnownFolderName.Drafts, sf, view);
                    }
                    else if (args[13] == "SentItems")
                    {
                        findResults = service.FindItems(WellKnownFolderName.SentItems, sf, view);
                    }
                    else if (args[13] == "JunkEmail")
                    {
                        findResults = service.FindItems(WellKnownFolderName.JunkEmail, sf, view);
                    }
                    else
                    {
                        ShowUsage();
                        Environment.Exit(0);
                    }
                    foreach (Item item in findResults.Items)
                    {
                        Console.WriteLine("\r\n");
                        EmailMessage email = EmailMessage.Bind(service, item.Id);
                        if (email.Subject != null)
                        {
                            Console.WriteLine("[*]Subject:{0}", email.Subject);
                        }
                        else
                        {
                            Console.WriteLine("[*]Subject:<null>");
                        }
                        Console.WriteLine("[*]HasAttachments:{0}", email.HasAttachments);
                        if (email.HasAttachments)
                        {
                            EmailMessage message = EmailMessage.Bind(service, email.Id, new PropertySet(ItemSchema.Attachments));
                            foreach (Attachment attachment in message.Attachments)
                            {
                                FileAttachment fileAttachment = attachment as FileAttachment;
                                fileAttachment.Load();
                                Console.WriteLine(" - Attachments:{0}", fileAttachment.Name);
                            }
                        }
                        Console.WriteLine("[*]ItemId:{0}", email.Id);
                        Console.WriteLine("[*]DateTimeCreated:{0}", email.DateTimeCreated);
                        Console.WriteLine("[*]DateTimeReceived:{0}", email.DateTimeReceived);
                        Console.WriteLine("[*]DateTimeSent:{0}", email.DateTimeSent);
                        Console.WriteLine("[*]DisplayCc:{0}", email.DisplayCc);
                        Console.WriteLine("[*]DisplayTo:{0}", email.DisplayTo);
                        Console.WriteLine("[*]InReplyTo:{0}", email.InReplyTo);
                        Console.WriteLine("[*]Size:{0}", item.Size);
                        email.Load(itempropertyset);
                        if (email.Body.ToString().Length > 100)
                        {
                            email.Body = email.Body.ToString().Substring(0, 100);
                            Console.WriteLine("[*]MessageBody(too big,only output 100):{0}", email.Body);
                        }
                        else
                        {
                            Console.WriteLine("[*]MessageBody:{0}", email.Body);
                        }
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine("[!]{0}", e.Message);
                    Environment.Exit(0);
                }
            }

            if ((args[11] == "ListFolder"))
            {
                Console.WriteLine("[+]Mode:{0}", args[11]);
                try
                {

                    if ((args[12] != "-Folder"))
                    {
                        ShowUsage();
                        Environment.Exit(0);
                    }
                    FolderString = args[13];
                    Console.WriteLine("[+]Folder:{0}", args[13]);

                    FindFoldersResults findResults = null;
                    FolderView view = new FolderView(int.MaxValue) { Traversal = FolderTraversal.Deep };

                    if (args[13] == "Inbox")
                    {
                        findResults = service.FindFolders(WellKnownFolderName.Inbox, view);
                    }
                    else if (args[13] == "Outbox")
                    {
                        findResults = service.FindFolders(WellKnownFolderName.Outbox, view);
                    }
                    else if (args[13] == "DeletedItems")
                    {
                        findResults = service.FindFolders(WellKnownFolderName.DeletedItems, view);
                    }
                    else if (args[13] == "Drafts")
                    {
                        findResults = service.FindFolders(WellKnownFolderName.Drafts, view);
                    }
                    else if (args[13] == "SentItems")
                    {
                        findResults = service.FindFolders(WellKnownFolderName.SentItems, view);
                    }
                    else if (args[13] == "JunkEmail")
                    {
                        findResults = service.FindFolders(WellKnownFolderName.JunkEmail, view);
                    }
                    else
                    {
                        ShowUsage();
                        Environment.Exit(0);
                    }
                    foreach (Folder folder in findResults.Folders)
                    {
                        Console.WriteLine("\r\n");
                        Console.WriteLine("[*]DisplayName:{0}", folder.DisplayName);
                        Console.WriteLine("[*]Id:{0}", folder.Id);
                        Console.WriteLine("[*]TotalCount:{0}", folder.TotalCount);
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine("[!]{0}", e.Message);
                    Environment.Exit(0);
                }
            }

            else if ((args[11] == "ListMailofFolder"))
            {
                Console.WriteLine("[+]Mode:{0}", args[11]);
                try
                {
                    if ((args[12] != "-Id"))
                    {
                        ShowUsage();
                        Environment.Exit(0);
                    }
                    IdString = args[13];
                    Console.WriteLine("[+]Id:{0}", args[13]);

                    Folder Folders = Folder.Bind(service, IdString);
                    FindItemsResults<Item> findResults = null;
                    ItemView view = new ItemView(int.MaxValue);
                    PropertySet itempropertyset = new PropertySet(BasePropertySet.FirstClassProperties);
                    itempropertyset.RequestedBodyType = BodyType.Text;
                    view.PropertySet = itempropertyset;
                    findResults = Folders.FindItems(view);
                    foreach (Item item in findResults.Items)
                    {
                        Console.WriteLine("\r\n");
                        if (item.Subject != null)
                        {
                            Console.WriteLine("[*]Subject:{0}", item.Subject);
                        }
                        else
                        {
                            Console.WriteLine("[*]Subject:<null>");
                        }

                        Console.WriteLine("[*]HasAttachments:{0}", item.HasAttachments);
                        if (item.HasAttachments)
                        {
                            EmailMessage message = EmailMessage.Bind(service, item.Id, new PropertySet(ItemSchema.Attachments));
                            foreach (Attachment attachment in message.Attachments)
                            {
                                FileAttachment fileAttachment = attachment as FileAttachment;
                                fileAttachment.Load();
                                Console.WriteLine(" - Attachments:{0}", fileAttachment.Name);
                            }
                        }
                        Console.WriteLine("[*]ItemId:{0}", item.Id);
                        Console.WriteLine("[*]DateTimeCreated:{0}", item.DateTimeCreated);
                        Console.WriteLine("[*]DateTimeReceived:{0}", item.DateTimeReceived);
                        Console.WriteLine("[*]DateTimeSent:{0}", item.DateTimeSent);
                        Console.WriteLine("[*]DisplayCc:{0}", item.DisplayCc);
                        Console.WriteLine("[*]DisplayTo:{0}", item.DisplayTo);
                        Console.WriteLine("[*]InReplyTo:{0}", item.InReplyTo);
                        Console.WriteLine("[*]Size:{0}", item.Size);
                        item.Load(itempropertyset);
                        if (item.Body.ToString().Length > 100)
                        {
                            item.Body = item.Body.ToString().Substring(0, 100);
                            Console.WriteLine("[*]MessageBody(too big,only output 100):{0}", item.Body);
                        }
                        else
                        {
                            Console.WriteLine("[*]MessageBody:{0}", item.Body);
                        }
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine("[!]{0}", e.Message);
                    Environment.Exit(0);
                }
            }

            else if ((args[11] == "ListUnreadMailofFolder"))
            {
                Console.WriteLine("[+]Mode:{0}", args[11]);
                try
                {
                    if ((args[12] != "-Id"))
                    {
                        ShowUsage();
                        Environment.Exit(0);
                    }
                    IdString = args[13];
                    Console.WriteLine("[+]Id:{0}", args[13]);

                    SearchFilter sf = new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, false);
                    Folder Folders = Folder.Bind(service, IdString);
                    FindItemsResults<Item> findResults = null;
                    ItemView view = new ItemView(int.MaxValue);
                    PropertySet itempropertyset = new PropertySet(BasePropertySet.FirstClassProperties);
                    itempropertyset.RequestedBodyType = BodyType.Text;
                    view.PropertySet = itempropertyset;
                    findResults = Folders.FindItems(sf, view);
                    foreach (Item item in findResults.Items)
                    {
                        Console.WriteLine("\r\n");
                        EmailMessage email = EmailMessage.Bind(service, item.Id);
                        if (email.Subject != null)
                        {
                            Console.WriteLine("[*]Subject:{0}", email.Subject);
                        }
                        else
                        {
                            Console.WriteLine("[*]Subject:<null>");
                        }

                        Console.WriteLine("[*]HasAttachments:{0}", email.HasAttachments);
                        if (email.HasAttachments)
                        {
                            EmailMessage message = EmailMessage.Bind(service, email.Id, new PropertySet(ItemSchema.Attachments));
                            foreach (Attachment attachment in message.Attachments)
                            {
                                FileAttachment fileAttachment = attachment as FileAttachment;
                                fileAttachment.Load();
                                Console.WriteLine(" - Attachments:{0}", fileAttachment.Name);
                            }
                        }
                        Console.WriteLine("[*]ItemId:{0}", email.Id);
                        Console.WriteLine("[*]DateTimeCreated:{0}", email.DateTimeCreated);
                        Console.WriteLine("[*]DateTimeReceived:{0}", email.DateTimeReceived);
                        Console.WriteLine("[*]DateTimeSent:{0}", email.DateTimeSent);
                        Console.WriteLine("[*]DisplayCc:{0}", email.DisplayCc);
                        Console.WriteLine("[*]DisplayTo:{0}", email.DisplayTo);
                        Console.WriteLine("[*]InReplyTo:{0}", email.InReplyTo);
                        Console.WriteLine("[*]Size:{0}", email.Size);
                        email.Load(itempropertyset);
                        if (email.Body.ToString().Length > 100)
                        {
                            email.Body = email.Body.ToString().Substring(0, 100);
                            Console.WriteLine("[*]MessageBody(too big,only output 100):{0}", email.Body);
                        }
                        else
                        {
                            Console.WriteLine("[*]MessageBody:{0}", email.Body);
                        }
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine("[!]{0}", e.Message);
                    Environment.Exit(0);
                }
            }

            else if ((args[11] == "SaveAttachment"))
            {
                Console.WriteLine("[+]Mode:{0}", args[11]);
                try
                {
                    if ((args[12] != "-Id"))
                    {
                        ShowUsage();
                        Environment.Exit(0);
                    }
                    IdString = args[13];
                    Console.WriteLine("[+]Id:{0}", args[13]);
                    EmailMessage message = EmailMessage.Bind(service, IdString, new PropertySet(ItemSchema.Attachments));
                    foreach (Attachment attachment in message.Attachments)
                    {
                        FileAttachment fileAttachment = attachment as FileAttachment;
                        Console.WriteLine("[+]Attachments:{0}", fileAttachment.Name);
                        fileAttachment.Load(fileAttachment.Name);
                        Console.WriteLine("\r\n[+]SaveAttachment success");
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine("[!]{0}", e.Message);
                    Environment.Exit(0);
                }
            }

            else if ((args[11] == "AddAttachment"))
            {
                Console.WriteLine("[+]Mode:{0}", args[11]);
                try
                {
                    if ((args[12] != "-Id"))
                    {
                        ShowUsage();
                        Environment.Exit(0);
                    }
                    IdString = args[13];
                    Console.WriteLine("[+]Id:{0}", args[13]);
                    if ((args[14] != "-AttachmentFile"))
                    {
                        ShowUsage();
                        Environment.Exit(0);
                    }
                    Console.WriteLine("[+]AttachmentFile:{0}", args[15]);
                    EmailMessage message = EmailMessage.Bind(service, IdString);
                    message.Attachments.AddFileAttachment(args[15]);
                    message.Update(ConflictResolutionMode.AlwaysOverwrite);
                    Console.WriteLine("\r\n[+]AddAttachment success");
                }
                catch (Exception e)
                {
                    Console.WriteLine("[!]{0}", e.Message);
                    Environment.Exit(0);
                }
            }

            else if ((args[11] == "ClearAllAttachment"))
            {
                Console.WriteLine("[+]Mode:{0}", args[11]);
                try
                {
                    if ((args[12] != "-Id"))
                    {
                        ShowUsage();
                        Environment.Exit(0);
                    }
                    IdString = args[13];
                    Console.WriteLine("[+]Id:{0}", args[13]);
                    EmailMessage message = EmailMessage.Bind(service, IdString, new PropertySet(ItemSchema.Attachments));
                    message.Attachments.Clear();
                    message.Update(ConflictResolutionMode.AlwaysOverwrite);
                    Console.WriteLine("\r\n[+]ClearAllAttachment success");
                }
                catch (Exception e)
                {
                    Console.WriteLine("[!]{0}", e.Message);
                    Environment.Exit(0);
                }
            }

            else if ((args[11] == "DeleteAttachment"))
            {
                Console.WriteLine("[+]Mode:{0}", args[11]);
                try
                {
                    if ((args[12] != "-Id"))
                    {
                        ShowUsage();
                        Environment.Exit(0);
                    }
                    IdString = args[13];
                    Console.WriteLine("[+]Id:{0}", args[13]);
                    if ((args[14] != "-AttachmentFile"))
                    {
                        ShowUsage();
                        Environment.Exit(0);
                    }
                    Console.WriteLine("[+]AttachmentFile:{0}", args[15]);
                    EmailMessage message = EmailMessage.Bind(service, IdString, new PropertySet(ItemSchema.Attachments));
                    foreach (Attachment attachment in message.Attachments)
                    {
                        if (attachment.Name == args[15])
                        {
                            message.Attachments.Remove(attachment);
                            break;
                        }
                    }
                    message.Update(ConflictResolutionMode.AlwaysOverwrite);
                    Console.WriteLine("\r\n[+]DeleteAttachment success");
                }
                catch (Exception e)
                {
                    Console.WriteLine("[!]{0}", e.Message);
                    Environment.Exit(0);
                }
            }

            else if ((args[11] == "DeleteMail"))
            {
                Console.WriteLine("[+]Mode:{0}", args[11]);
                try
                {
                    if ((args[12] != "-Id"))
                    {
                        ShowUsage();
                        Environment.Exit(0);
                    }
                    IdString = args[13];
                    Console.WriteLine("[+]Id:{0}", args[13]);
                    EmailMessage message = EmailMessage.Bind(service, IdString);
                    message.Delete(DeleteMode.SoftDelete);
                    Console.WriteLine("\r\n[+]DeleteMail success");
                }
                catch (Exception e)
                {
                    Console.WriteLine("[!]{0}", e.Message);
                    Environment.Exit(0);
                }
            }

            else if ((args[11] == "SearchMail"))
            {
                Console.WriteLine("[+]Mode:{0}", args[11]);
                try
                {
                    if ((args[12] != "-String"))
                    {
                        ShowUsage();
                        Environment.Exit(0);
                    }
                    SearchString = args[13];
                    Console.WriteLine("[+]SearchString:{0}", args[13]);
                    SearchMail(service, WellKnownFolderName.Inbox, SearchString);
                    SearchMail(service, WellKnownFolderName.Outbox, SearchString);
                    SearchMail(service, WellKnownFolderName.DeletedItems, SearchString);
                    SearchMail(service, WellKnownFolderName.Drafts, SearchString);
                    SearchMail(service, WellKnownFolderName.SentItems, SearchString);
                    SearchMail(service, WellKnownFolderName.JunkEmail, SearchString);
                    Console.WriteLine("\r\n[+]SearchMail done");
                }
                catch (Exception e)
                {
                    Console.WriteLine("[!]{0}", e.Message);
                    Environment.Exit(0);
                }
            }

            else if ((args[11] == "ExportMail"))
            {
                Console.WriteLine("[+]Mode:{0}", args[11]);
                try
                {
                    if ((args[12] != "-Folder"))
                    {
                        ShowUsage();
                        Environment.Exit(0);
                    }
                    IdString = args[13];
                    Console.WriteLine("[+]Folder:{0}", args[13]);
                    Console.WriteLine("[+]SavePath:{0}.eml", args[13]);
                    Folder inbox = null;
                    if (args[13] == "Inbox")
                    {
                        inbox = Folder.Bind(service, WellKnownFolderName.Inbox);
                    }
                    else if (args[13] == "Outbox")
                    {
                        inbox = Folder.Bind(service, WellKnownFolderName.Outbox);
                    }
                    else if (args[13] == "DeletedItems")
                    {
                        inbox = Folder.Bind(service, WellKnownFolderName.DeletedItems);
                    }
                    else if (args[13] == "Drafts")
                    {
                        inbox = Folder.Bind(service, WellKnownFolderName.Drafts);
                    }
                    else if (args[13] == "SentItems")
                    {
                        inbox = Folder.Bind(service, WellKnownFolderName.SentItems);
                    }
                    else if (args[13] == "JunkEmail")
                    {
                        inbox = Folder.Bind(service, WellKnownFolderName.JunkEmail);
                    }
                    else
                    {
                        Console.WriteLine("[!]Don't support this folder");
                        Environment.Exit(0);
                    }
                    ItemView view = new ItemView(int.MaxValue);
                    view.PropertySet = new PropertySet(BasePropertySet.IdOnly);
                    FindItemsResults<Item> results = inbox.FindItems(view);
                    int i = 0;
                    foreach (var item in results)
                    {
                        i++;
                        PropertySet props = new PropertySet(EmailMessageSchema.MimeContent);
                        var email = EmailMessage.Bind(service, item.Id, props);
                        string emlFileName = args[13] + ".eml";
                        using (FileStream fs = new FileStream(emlFileName, FileMode.Append, FileAccess.Write))
                        {
                            fs.Write(email.MimeContent.Content, 0, email.MimeContent.Content.Length);
                        }
                    }
                    Console.WriteLine("\r\n[+]ExportMail done,total number:{0}", i);
                }
                catch (Exception e)
                {
                    Console.WriteLine("[!]{0}", e.Message);
                    Environment.Exit(0);
                }
            }

            else if ((args[11] == "ViewMail"))
            {
                Console.WriteLine("[+]Mode:{0}", args[11]);
                try
                {
                    if ((args[12] != "-Id"))
                    {
                        ShowUsage();
                        Environment.Exit(0);
                    }
                    IdString = args[13];
                    Console.WriteLine("[+]Id:{0}", args[13]);
                    ViewMail(service, WellKnownFolderName.Inbox, IdString);
                    ViewMail(service, WellKnownFolderName.Outbox, IdString);
                    ViewMail(service, WellKnownFolderName.DeletedItems, IdString);
                    ViewMail(service, WellKnownFolderName.Drafts, IdString);
                    ViewMail(service, WellKnownFolderName.SentItems, IdString);
                    ViewMail(service, WellKnownFolderName.JunkEmail, IdString);
                    Console.WriteLine("\r\n[+]ViewMail done");
                }
                catch (Exception e)
                {
                    Console.WriteLine("[!]{0}", e.Message);
                    Environment.Exit(0);
                }
            }

            else if ((args[11] == "ReadXML"))
            {
                Console.WriteLine("[+]Mode:{0}", args[11]);
                if ((args[12] != "-Path"))
                {
                    ShowUsage();
                    Environment.Exit(0);
                }
                Console.WriteLine("[+]XML path:{0}", args[13]);
                Console.WriteLine("[+]Out path:out.xml");
                StreamReader sendData = new StreamReader(args[13], Encoding.Default);
                byte[] sendDataByte = Encoding.UTF8.GetBytes(sendData.ReadToEnd());
                sendData.Close();
                try
                {
                    HttpWebRequest request = (HttpWebRequest)WebRequest.Create(service.Url);
                    request.Method = "POST";
                    request.ContentType = "text/xml";
                    request.ContentLength = sendDataByte.Length;
                    request.AllowAutoRedirect = false;
                    if (args[4] == "-use")
                    {
                        request.Credentials = CredentialCache.DefaultCredentials;
                    }
                    else
                    {
                        request.Credentials = new NetworkCredential(UserString, PwdString);
                    }

                    Stream requestStream = request.GetRequestStream();
                    requestStream.Write(sendDataByte, 0, sendDataByte.Length);
                    requestStream.Close();

                    HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                    if (response.StatusCode != HttpStatusCode.OK)
                    {
                        throw new WebException(response.StatusDescription);
                    }
                    Stream receiveStream = response.GetResponseStream();
                    StreamReader readStream = new StreamReader(receiveStream, Encoding.UTF8);
                    String receiveString = readStream.ReadToEnd();
                    response.Close();
                    readStream.Close();
                    StreamWriter receiveData = new StreamWriter("out.xml");
                    receiveData.Write(receiveString);
                    receiveData.Close();
                }

                catch (WebException e)
                {
                    Console.WriteLine("[!]{0}", e.Message);
                    Environment.Exit(0);
                }
                Console.WriteLine("\r\n[+]ReadXML done");
            }

        }

        private static void ViewMail(ExchangeService service, WellKnownFolderName folder, string IdString)
        {
            ItemView view = new ItemView(int.MaxValue);
            PropertySet itempropertyset = new PropertySet(BasePropertySet.FirstClassProperties);
            itempropertyset.RequestedBodyType = BodyType.Text;
            view.PropertySet = itempropertyset;
            FindItemsResults<Item> findResults = null;
            findResults = service.FindItems(folder, view);
            foreach (Item item in findResults.Items)
            {
                if (item.Id.ToString() == IdString)
                {
                    Console.WriteLine("\r\n");
                    Console.WriteLine("[+]Mail found");
                    Console.WriteLine("[*]Folder:{0}", folder);
                    if (item.Subject != null)
                    {
                        Console.WriteLine("[*]Subject:{0}", item.Subject);
                    }
                    else
                    {
                        Console.WriteLine("[*]Subject:<null>");
                    }
                    Console.WriteLine("[*]HasAttachments:{0}", item.HasAttachments);
                    if (item.HasAttachments)
                    {
                        EmailMessage message = EmailMessage.Bind(service, item.Id, new PropertySet(ItemSchema.Attachments));
                        foreach (Attachment attachment in message.Attachments)
                        {
                            FileAttachment fileAttachment = attachment as FileAttachment;
                            fileAttachment.Load();
                            Console.WriteLine(" - Attachments:{0}", fileAttachment.Name);
                        }
                    }
                    Console.WriteLine("[*]ItemId:{0}", item.Id);
                    Console.WriteLine("[*]DateTimeCreated:{0}", item.DateTimeCreated);
                    Console.WriteLine("[*]DateTimeReceived:{0}", item.DateTimeReceived);
                    Console.WriteLine("[*]DateTimeSent:{0}", item.DateTimeSent);
                    Console.WriteLine("[*]DisplayCc:{0}", item.DisplayCc);
                    Console.WriteLine("[*]DisplayTo:{0}", item.DisplayTo);
                    Console.WriteLine("[*]InReplyTo:{0}", item.InReplyTo);
                    Console.WriteLine("[*]Categories:{0}", item.Categories);
                    Console.WriteLine("[*]Culture:{0}", item.Culture);
                    Console.WriteLine("[*]IsFromMe:{0}", item.IsFromMe);
                    Console.WriteLine("[*]ItemClass:{0}", item.ItemClass);
                    Console.WriteLine("[*]Size:{0}", item.Size);
                    item.Load(itempropertyset);
                    Console.WriteLine("[*]MessageBody:{0}", item.Body);
                }
            }
        }

        private static void SearchMail(ExchangeService service, WellKnownFolderName folder, string SearchString)
        {
            ItemView view = new ItemView(int.MaxValue);
            PropertySet itempropertyset = new PropertySet(BasePropertySet.FirstClassProperties);
            itempropertyset.RequestedBodyType = BodyType.Text;
            view.PropertySet = itempropertyset;
            FindItemsResults<Item> findResults = null;
            findResults = service.FindItems(folder, view);
            foreach (Item item in findResults.Items)
            {
                if (item.Subject.Contains(SearchString))
                {
                    Console.WriteLine("\r\n");
                    Console.WriteLine("[+]Match mail found,Subject match");
                    Console.WriteLine("[*]Folder:{0}", folder);
                    Console.WriteLine("[*]Subject:{0}", item.Subject);
                    Console.WriteLine("[*]Size:{0}", item.Size);
                    Console.WriteLine("[*]ItemId:{0}", item.Id);
                }
                if (item.HasAttachments)
                {
                    EmailMessage message = EmailMessage.Bind(service, item.Id, new PropertySet(ItemSchema.Attachments));
                    foreach (Attachment attachment in message.Attachments)
                    {
                        FileAttachment fileAttachment = attachment as FileAttachment;
                        fileAttachment.Load();
                        if (fileAttachment.Name.Contains(SearchString))
                        {
                            Console.WriteLine("\r\n");
                            Console.WriteLine("[+]Match mail found,Attachment match");
                            Console.WriteLine("[*]Folder:{0}", folder);
                            Console.WriteLine("[*]Subject:{0}", item.Subject);
                            Console.WriteLine(" - Attachments:{0}", fileAttachment.Name);
                            Console.WriteLine("[*]Size:{0}", item.Size);
                            Console.WriteLine("[*]ItemId:{0}", item.Id);
                        }
                    }
                }
                item.Load(itempropertyset);
                if (item.Body.Text != null)
                {
                    if (item.Body.Text.Contains(SearchString))
                    {
                        Console.WriteLine("\r\n");
                        Console.WriteLine("[+]Match mail found,MessageBody match");
                        Console.WriteLine("[*]Folder:{0}", folder);
                        Console.WriteLine("[*]Subject:{0}", item.Subject);
                        Console.WriteLine("[*]Size:{0}", item.Size);
                        Console.WriteLine("[*]ItemId:{0}", item.Id);
                    }
                }
            }
        }

        private static bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            bool result = false;
            Uri redirectionUri = new Uri(redirectionUrl);
            if (redirectionUri.Scheme == "https")
            {
                result = true;
            }
            return result;
        }
    }
}
