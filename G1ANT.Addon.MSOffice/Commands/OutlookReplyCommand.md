# outlook.reply

## Syntax

```G1ANT
outlook.reply item ⟦outlookmail/outlookfolder⟧
```

## Description

This command creates a new variable of [outlookmail](G1ANT.Addon/G1ANT.Addon.MSOffice/G1ANT.Addon.MSOffice/Structures/OutlookMailStructure.md) structure which is a reply to a specified mail.

| Argument       | Type                                                         | Required | Default Value                                                | Description                                                  |
| -------------- | ------------------------------------------------------------ | -------- | ------------------------------------------------------------ | ------------------------------------------------------------ |
| `mail`         | [outlookmail](G1ANT.Addon/G1ANT.Addon.MSOffice/G1ANT.Addon.MSOffice/Structures/OutlookMailStructure.md) | yes      |                                                              | Mail message to be replied to                                |
| `result`       | [variable](G1ANT.Language/G1ANT.Language/Structures/VariableStructure.md) | no       | `♥result`                                                    | Name of a variable where the command's result will be stored. The variable will be of [outlookmail](G1ANT.Addon/G1ANT.Addon.MSOffice/G1ANT.Addon.MSOffice/Structures/OutlookMailStructure.md) structure |
| `if`           | [bool](G1ANT.Language/G1ANT.Language/Structures/BooleanStructure.md) | no       | true                                                         | Executes the command only if a specified condition is true   |
| `timeout`      | [timespan](G1ANT.Language/G1ANT.Language/Structures/TimeSpanStructure.md) | no       | [♥timeoutcommand](G1ANT.Language/G1ANT.Addon.Core/Variables/TimeoutCommandVariable.md) | Specifies time in milliseconds for G1ANT.Robot to wait for the command to be executed |
| `errorcall`    | [procedure](G1ANT.Language/G1ANT.Language/Structures/ProcedureStructure.md) | no       |                                                              | Name of a procedure to call when the command throws an exception or when a given `timeout` expires |
| `errorjump`    | [label](G1ANT.Language/G1ANT.Language/Structures/LabelStructure.md) | no       |                                                              | Name of the label to jump to when the command throws an exception or when a given `timeout` expires |
| `errormessage` | [text](G1ANT.Language/G1ANT.Language/Structures/TextStructure.md) | no       |                                                              | A message that will be shown in case the command throws an exception or when a given `timeout` expires, and no `errorjump` argument is specified |
| `errorresult`  | [variable](G1ANT.Language/G1ANT.Language/Structures/VariableStructure.md) | no       |                                                              | Name of a variable that will store the returned exception. The variable will be of [error](G1ANT.Language/G1ANT.Language/Structures/ErrorStructure.md) structure |

For more information about `if`, `timeout`, `errorcall`, `errorjump`, `errormessage` and `errorresult` arguments, see [Common Arguments](G1ANT.Manual/appendices/common-arguments.md) page.

## Example

In this example, the robot gets the content of the Outlook Inbox folder, assigns mails to the `♥mails` variable, then prepares a reply to the first unread email message (starting from the oldest unread element in the Inbox). The resulting draft reply is assigned to the `♥replyMail` variable. Since it is of `outlookmail` structure, you can retrieve the message body (the conversation history, because it’s a reply) and add a reply text to it: here, it’s done by adding the `♥replyText` to the `♥replyMailBody` variables. Then, the reply message stored in the `♥replyMail` variable is sent with the `outlook.send` command, using the `mail` argument:

```G1ANT
♥outlookInboxFolder = john.doe@g1ant.com\Inbox
outlook.open
outlook.getfolder ♥outlookInboxFolder result ♥inboxFolder
♥mails = ♥inboxFolder⟦unread⟧
outlook.reply ♥mails⟦1⟧ result ♥replyMail
♥replyMailBody = ♥replyMail⟦body⟧
♥replyText = ⊂"Hi,\r\nThanks for your email. We will contact you shortly.\r\nRegards,\r\nG1ANT\r\n\r\n"⊃
♥replyMail⟦body⟧ = ♥replyText + ♥replyMailBody
outlook.send mail ♥replyMail
outlook.close
```
