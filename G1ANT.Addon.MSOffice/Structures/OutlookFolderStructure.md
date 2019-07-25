# outlookfolder

This structure stores information about the Outlook folder, which was retrieved with the `outlook.getfolder` command. It contains the following fields:

| Field        | Type                                                        | Description                                     |
| ------------ | ----------------------------------------------------------- | ----------------------------------------------- |
| `name`       | [text](https://manual.g1ant.com/link/G1ANT.Language/G1ANT.Language/Structures/TextStructure.md) | The name of the folder                          |
| `folderpath` | [text](https://manual.g1ant.com/link/G1ANT.Language/G1ANT.Language/Structures/TextStructure.md) | The path to the folder                          |
| `folders`    | [list](https://manual.g1ant.com/link/G1ANT.Language/G1ANT.Language/Structures/ListStructure.md) | The list of subfolders                          |
| `mails`      | [list](https://manual.g1ant.com/link/G1ANT.Language/G1ANT.Language/Structures/ListStructure.md) | The list of email messages stored in the folder |
| `unread`   | [list](https://manual.g1ant.com/link/G1ANT.Language/G1ANT.Language/Structures/ListStructure.md) | The list of unread messages                     |

## Example

The script below retrieves unread emails from the Outlook Inbox folder, using the `♥inboxFolder` variable, which is of the `outlookfolder` structure (be sure to provide the correct Outlook folder information in the `♥outlookInboxFolder` variable):

```G1ANT
♥outlookInboxFolder = john.doe@g1ant.com\Inbox

outlook.open display false
outlook.getfolder ♥outlookInboxFolder result ♥inboxFolder errormessage ‴Cannot find folder "♥outlookInboxFolder"‴
♥unreademails = ♥inboxFolder⟦unread⟧
foreach ♥email in ♥unreademails
  dialog ‴New message from ♥email⟦from⟧ with subject: "♥email⟦subject⟧"‴
end
```

Note that another Outlook structure is used here as well: [outlookmail](https://manual.g1ant.com/link/G1ANT.Addon.MSOffice/G1ANT.Addon.MSOffice/Structures/outlookmailstructure.md) (for the `♥email` variable).
