# outlookmail

This structure stores information about a mail message, which was retrieved from the Outlook folder with the `outlook.getfolder` command. It contains the following fields:

| Field         | Type                                                        | Description                              |
| ------------- | ----------------------------------------------------------- | ---------------------------------------- |
| `id`          | [text](G1ANT.Language/G1ANT.Language/Structures/TextStructure.md) | The identification number of the message |
| `from`        | [text](G1ANT.Language/G1ANT.Language/Structures/TextStructure.md) | The sender of the message                |
| `cc`        | [text](G1ANT.Language/G1ANT.Language/Structures/TextStructure.md) | The CC recipient(s) of the message                |
| `bcc`        | [text](G1ANT.Language/G1ANT.Language/Structures/TextStructure.md) | The BCC recipient(s) of the message                |
| `account`          | [text](G1ANT.Language/G1ANT.Language/Structures/TextStructure.md) | The recipient of the message             |
| `subject`     | [text](G1ANT.Language/G1ANT.Language/Structures/TextStructure.md) | The subject of the message               |
| `body`        | [text](G1ANT.Language/G1ANT.Language/Structures/TextStructure.md) | The contents of the message              |
| `htmlbody`    | [text](G1ANT.Language/G1ANT.Language/Structures/TextStructure.md) | The HTML contents of the message         |
| `attachments` | [list](G1ANT.Language/G1ANT.Language/Structures/ListStructure.md) | The list of attachments in the message   |

## Example

The script below retrieves unread emails from the Outlook Inbox folder, checks their subjects one by one, using the `♥email` variable, which is of the `outlookmail` structure, and if it finds any containing the word “invoice”, it writes a text file on the user’s Desktop under a filename composed from the sender’s address and a message subject with the message body (content) as the file’s content (be sure to provide the correct Outlook folder information in the `♥outlookInboxFolder` variable):

```G1ANT
♥outlookInboxFolder = john.doe@g1ant.com\Inbox

outlook.open display false
outlook.getfolder ♥outlookInboxFolder result ♥inboxFolder errormessage ‴Cannot find folder "♥outlookInboxFolder"‴
♥unreadEmails = ♥inboxFolder⟦unread⟧
foreach ♥email in ♥unreadEmails
  if ⊂♥email⟦subject⟧.Contains("invoice")⊃
    text.write ♥email⟦body⟧ filename ‴♥environment⟦USERPROFILE⟧\Desktop\♥email⟦from⟧ - ♥email⟦subject⟧.txt‴
  end
end
```

Note that another Outlook structure is used here as well: [outlookfolder](outlookfolderstructure.md) (for the `♥inboxFolder` variable).
