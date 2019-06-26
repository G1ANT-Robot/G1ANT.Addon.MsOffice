# outlook.selectitem

## Syntax

```G1ANT
outlook.selectitem item ⟦outlookmail/outlookfolder⟧
```

## Description

This command selects a mail or a folder element in Outlook’s user interface.

| Argument       | Type                                                         | Required | Default Value                                                | Description                                                  |
| -------------- | ------------------------------------------------------------ | -------- | ------------------------------------------------------------ | ------------------------------------------------------------ |
| `item`         | [outlookmail](G1ANT.Addon/G1ANT.Addon.MSOffice/G1ANT.Addon.MSOffice/Structures/OutlookMailStructure.md) or [outlookfolder](G1ANT.Addon/G1ANT.Addon.MSOffice/G1ANT.Addon.MSOffice/Structures/OutlookFolderStructure.md) | yes      |                                                              | Item — a mail message or a folder — to be selected           |
| `result`       | [variable](G1ANT.Language/G1ANT.Language/Structures/VariableStructure.md) | no       | `♥result`                                                    | Name of a variable where the command's result will be stored |
| `if`           | [bool](G1ANT.Language/G1ANT.Language/Structures/BooleanStructure.md) | no       | true                                                         | Executes the command only if a specified condition is true   |
| `timeout`      | [timespan](G1ANT.Language/G1ANT.Language/Structures/TimeSpanStructure.md) | no       | [♥timeoutcommand](G1ANT.Language/G1ANT.Addon.Core/Variables/TimeoutCommandVariable.md) | Specifies time in milliseconds for G1ANT.Robot to wait for the command to be executed |
| `errorcall`    | [procedure](G1ANT.Language/G1ANT.Language/Structures/ProcedureStructure.md) | no       |                                                              | Name of a procedure to call when the command throws an exception or when a given `timeout` expires |
| `errorjump`    | [label](G1ANT.Language/G1ANT.Language/Structures/LabelStructure.md) | no       |                                                              | Name of the label to jump to when the command throws an exception or when a given `timeout` expires |
| `errormessage` | [text](G1ANT.Language/G1ANT.Language/Structures/TextStructure.md) | no       |                                                              | A message that will be shown in case the command throws an exception or when a given `timeout` expires, and no `errorjump` argument is specified |
| `errorresult`  | [variable](G1ANT.Language/G1ANT.Language/Structures/VariableStructure.md) | no       |                                                              | Name of a variable that will store the returned exception. The variable will be of [error](G1ANT.Language/G1ANT.Language/Structures/ErrorStructure.md) structure |

For more information about `if`, `timeout`, `errorcall`, `errorjump`, `errormessage` and `errorresult` arguments, see [Common Arguments](G1ANT.Manual/appendices/common-arguments.md) page.

## Example

In this example, the robot selects the first (the oldest) unread email message in the Inbox folder (be sure to provide the correct Outlook folder information in the `♥outlookInboxFolder` variable):

```G1ANT
♥outlookInboxFolder = john.doe@g1ant.com\Inbox

outlook.open
window ✱Outlook
outlook.getfolder ♥outlookInboxFolder result ♥inboxFolder errormessage ‴Cannot find folder "♥outlookInboxFolder"‴
♥unreademails = ♥inboxFolder⟦unread⟧
outlook.selectitem ♥unreademails⟦1⟧
```