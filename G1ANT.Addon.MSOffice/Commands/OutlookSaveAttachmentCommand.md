# outlook.saveattachment

## Syntax

```G1ANT
outlook.saveattachment attachment ⟦outlookattachment⟧ path ⟦path⟧
```

## Description

This command saves an attachment to a file.

| Argument       | Type                                                         | Required | Default Value                                                | Description                                                  |
| -------------- | ------------------------------------------------------------ | -------- | ------------------------------------------------------------ | ------------------------------------------------------------ |
| `attachment`   | [outlookattachment](G1ANT.Addon/G1ANT.Addon.MSOffice/G1ANT.Addon.MSOffice/Structures/OutlookAttachmentStructure.md) | yes      |                                                              | Email attachment to be saved                                 |
| `path`         | [path](G1ANT.Language/G1ANT.Language/Structures/PathStructure.md) | yes      |                                                              | Path to the saved file                                       |
| `if`           | [bool](G1ANT.Language/G1ANT.Language/Structures/BooleanStructure.md) | no       | true                                                         | Executes the command only if a specified condition is true   |
| `timeout`      | [timespan](G1ANT.Language/G1ANT.Language/Structures/TimeSpanStructure.md) | no       | [♥timeoutcommand](G1ANT.Language/G1ANT.Addon.Core/Variables/TimeoutCommandVariable.md) | Specifies time in milliseconds for G1ANT.Robot to wait for the command to be executed |
| `errorcall`    | [procedure](G1ANT.Language/G1ANT.Language/Structures/ProcedureStructure.md) | no       |                                                              | Name of a procedure to call when the command throws an exception or when a given `timeout` expires |
| `errorjump`    | [label](G1ANT.Language/G1ANT.Language/Structures/LabelStructure.md) | no       |                                                              | Name of the label to jump to when the command throws an exception or when a given `timeout` expires |
| `errormessage` | [text](G1ANT.Language/G1ANT.Language/Structures/TextStructure.md) | no       |                                                              | A message that will be shown in case the command throws an exception or when a given `timeout` expires, and no `errorjump` argument is specified |
| `errorresult`  | [variable](G1ANT.Language/G1ANT.Language/Structures/VariableStructure.md) | no       |                                                              | Name of a variable that will store the returned exception. The variable will be of [error](G1ANT.Language/G1ANT.Language/Structures/ErrorStructure.md) structure |

For more information about `if`, `timeout`, `errorcall`, `errorjump`, `errormessage` and `errorresult` arguments, see [Common Arguments](G1ANT.Manual/appendices/common-arguments.md) page.

## Example

In this example, the robot reads emails from the Outlook Inbox folder, then processes them one by one and saves all attachments to files located in Attachments folder on the user’s Desktop, using their original names (be sure to provide the correct Outlook folder information in the `♥outlookInboxFolder` variable):

```G1ANT
♥outlookInboxFolder = john.doe@g1ant.com\Inbox

outlook.open display false
outlook.getfolder ♥outlookInboxFolder result ♥inboxFolder errormessage ‴Cannot find folder "♥outlookInboxFolder"‴
♥emails = ♥inboxFolder⟦mails⟧
foreach ♥email in ♥emails
  ♥attachments = ♥email⟦attachments⟧
  foreach ♥attachment in ♥attachments
    outlook.saveattachment ♥attachment path ♥environment⟦USERPROFILE⟧\Desktop\Attachments\♥attachment⟦filename⟧
  end
end
```