# ewsManage
My exercise of using Exchange Web Service(EWS)

Author: 3gstudent

License: BSD 3-Clause

---

The following functions are currently supported:

- Support EWS Managed API and EWS SOAP XML message
- Use the default credentials of the logged on user or another user
- Accept all certificates, regardless of why they are invalid(optional)
- List mails at specified locations, including file names in attachments and mail body 
(Judging the length of mail content, if more than 100 characters, only the first 100 characters will be displayed.)
- List unread messages at specified locations, including file names in attachments and mail body 
(Judging the length of mail content, if more than 100 characters, only the first 100 characters will be displayed.)
- List custom folders in the specified location (traverse all subfolders)
- View all messages under custom files
- View unread messages under custom files
- Save all messages in the specified location (in EML format)
- Save attachments in specified mail (specified ID)
- Add attachments (specified ID) to the specified message
- Delete attachments to specified messages (specified ID)
- Delete all attachments to a specified message
- Search for mail with specified keywords (common location, search title, attachment name and mail body)
- Delete the specified message (specified ID)
- View the specific content of a message (specified ID)
- Send mail (using EWS SOAP)
- Read the XML file and send commands through EWS SOAP

Location to support queries and operations:

- Inbox
- Drafts
- Sent Items
- Deleted Items
- Outbox
- Junk Email

eg.

ewsManage.exe -CerValidation Yes -ExchangeVersion Exchange2013_SP1 -u test1 -p test123! -ewsPath https://test.com/ews/Exchange.asmx -Mode ListUnreadMail -Folder Inbox

ewsManage.exe -CerValidation No -ExchangeVersion Exchange2013_SP1 -use the default credentials -AutodiscoverUrl test1@test.com -Mode ListMail -Folder SentItems

ewsManage.exe -CerValidation No -ExchangeVersion Exchange2013_SP1 -u test1 -p test123! -ewsPath https://test.com/ews/Exchange.asmx -Mode ListFolder -Folder Inbox

ewsManage.exe -CerValidation No -ExchangeVersion Exchange2013_SP1 -u test1 -p test123! -ewsPath https://test.com/ews/Exchange.asmx -Mode ListMailofFolder -Id AAMaADFlMjRjMdM2LTgxZTUtNGRmZC05ZDQyLTMzNDFlMzBmZWY1NwAzAAAAAAAR9UOK286vT6HjUgukBQGmAQBHzR2O8KNmTcffGwlY0A76AAAAADfqAAA=

ewsManage.exe -CerValidation No -ExchangeVersion Exchange2013_SP1 -u test1 -p test123! -ewsPath https://test.com/ews/Exchange.asmx -Mode ExportMail -Folder Inbox

ewsManage.exe -CerValidation No -ExchangeVersion Exchange2013_SP1 -u test1 -p test123! -ewsPath https://test.com/ews/Exchange.asmx -Mode SaveAttachment -Id AAMzADFlMjRjMzM3LTgxZTUzNGRmZC25ZDQyLTMaNDFlMzBwZWY1NwBGAAAAAAAR8UOK236vT6HjUnujBQGmBwBHzR1O8KNmTrjfGwlY0A56AAAAAAEKAABHzR1O8KNmTrjfGzlY2A75AAAAABxFAAA=

ewsManage.exe -CerValidation No -ExchangeVersion Exchange2013_SP1 -u test1 -p test123! -ewsPath https://test.com/ews/Exchange.asmx -Mode AddAttachment -Id AAMzADFlMjRjMzM3LTgxZTUzNGRmZC25ZDQyLTMaNDFlMzBwZWY1NwBGAAAAAAAR8UOK236vT6HjUnujBQGmBwBHzR1O8KNmTrjfGwlY0A56AAAAAAEKAABHzR1O8KNmTrjfGzlY2A75AAAAABxFAAA= -AttachmentFile 1.txt

ewsManage.exe -CerValidation No -ExchangeVersion Exchange2013_SP1 -u test1 -p test123! -ewsPath https://test.com/ews/Exchange.asmx -Mode DeleteAttachment -Id AAMzADFlMjRjMzM3LTgxZTUzNGRmZC25ZDQyLTMaNDFlMzBwZWY1NwBGAAAAAAAR8UOK236vT6HjUnujBQGmBwBHzR1O8KNmTrjfGwlY0A56AAAAAAEKAABHzR1O8KNmTrjfGzlY2A75AAAAABxFAAA= -AttachmentFile 1.txt

ewsManage.exe -CerValidation No -ExchangeVersion Exchange2013_SP1 -u test1 -p test123! -ewsPath https://test.com/ews/Exchange.asmx -Mode ClearAllAttachment -Id AAMzADFlMjRjMzM3LTgxZTUzNGRmZC25ZDQyLTMaNDFlMzBwZWY1NwBGAAAAAAAR8UOK236vT6HjUnujBQGmBwBHzR1O8KNmTrjfGwlY0A56AAAAAAEKAABHzR1O8KNmTrjfGzlY2A75AAAAABxFAAA=

ewsManage.exe -CerValidation No -ExchangeVersion Exchange2013_SP1 -u test1 -p test123! -ewsPath https://test.com/ews/Exchange.asmx -Mode SearchMail -String vpn

ewsManage.exe -CerValidation No -ExchangeVersion Exchange2013_SP1 -u test1 -p test123! -ewsPath https://test.com/ews/Exchange.asmx -Mode DeleteMail -Id AAMzADFlMjRjMzM3LTgxZTUzNGRmZC25ZDQyLTMaNDFlMzBwZWY1NwBGAAAAAAAR8UOK236vT6HjUnujBQGmBwBHzR1O8KNmTrjfGwlY0A56AAAAAAEKAABHzR1O8KNmTrjfGzlY2A75AAAAABxFAAA=

ewsManage.exe -CerValidation No -ExchangeVersion Exchange2013_SP1 -u test1 -p test123! -ewsPath https://test.com/ews/Exchange.asmx -Mode ViewMail -Id AAMzADFlMjRjMzM3LTgxZTUzNGRmZC25ZDQyLTMaNDFlMzBwZWY1NwBGAAAAAAAR8UOK236vT6HjUnujBQGmBwBHzR1O8KNmTrjfGwlY0A56AAAAAAEKAABHzR1O8KNmTrjfGzlY2A75AAAAABxFAAA=

ewsManage.exe -CerValidation No -ExchangeVersion Exchange2013_SP1 -u test1 -p test123! -ewsPath https://test.com/ews/Exchange.asmx -Mode ReadXML -Path ews.xml
