# ewsManage
My exercise of using Exchange Web Service(EWS)

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
