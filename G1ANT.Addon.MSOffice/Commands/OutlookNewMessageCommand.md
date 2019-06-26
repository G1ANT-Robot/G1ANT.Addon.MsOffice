# outlook.newmessage

## Syntax

```G1ANT
outlook.newmessage to ⟦text⟧ subject ⟦text⟧ body ⟦text⟧ attachments ⟦list⟧ isbodyhtml ⟦bool⟧
```

## Description

This command opens a new message window and fills it up with provided information.

| Argument | Type | Required | Default Value | Description |
| -------- | ---- | -------- | ------------- | ----------- |
|`to`| [text](G1ANT.Language/G1ANT.Language/Structures/TextStructure.md) | no |  | Mail recipients |
|`subject`| [text](G1ANT.Language/G1ANT.Language/Structures/TextStructure.md) | no |  | Mail subject |
|`body`| [text](G1ANT.Language/G1ANT.Language/Structures/TextStructure.md) | no |  | Mail body |
|`attachments`| [list](G1ANT.Language/G1ANT.Language/Structures/ListStructure.md) | no |  | List of attachments (as their filepaths) to be included in a mail message. Elements should be separated with ❚ character (**Ctrl+\\**) |
|`isbodyhtml`| [bool](G1ANT.Language/G1ANT.Language/Structures/BooleanStructure.md) | no | false | If set to `true`, indicates that the mail message body is in HTML |
| `result`       | [variable](G1ANT.Language/G1ANT.Language/Structures/VariableStructure.md) | no       | `♥result`                                                   | Name of a variable where the command's result will be stored |
| `if`           | [bool](G1ANT.Language/G1ANT.Language/Structures/BooleanStructure.md) | no       | true                                                        | Executes the command only if a specified condition is true   |
| `timeout`      | [timespan](G1ANT.Language/G1ANT.Language/Structures/TimeSpanStructure.md) | no       | [♥timeoutcommand](G1ANT.Language/G1ANT.Addon.Core/Variables/TimeoutCommandVariable.md) | Specifies time in milliseconds for G1ANT.Robot to wait for the command to be executed |
| `errorcall`    | [procedure](G1ANT.Language/G1ANT.Language/Structures/ProcedureStructure.md) | no       |                                                             | Name of a procedure to call when the command throws an exception or when a given `timeout` expires |
| `errorjump`    | [label](G1ANT.Language/G1ANT.Language/Structures/LabelStructure.md) | no       |                                                             | Name of the label to jump to when the command throws an exception or when a given `timeout` expires |
| `errormessage` | [text](G1ANT.Language/G1ANT.Language/Structures/TextStructure.md) | no       |                                                             | A message that will be shown in case the command throws an exception or when a given `timeout` expires, and no `errorjump` argument is specified |
| `errorresult`  | [variable](G1ANT.Language/G1ANT.Language/Structures/VariableStructure.md) | no       |                                                             | Name of a variable that will store the returned exception. The variable will be of [error](G1ANT.Language/G1ANT.Language/Structures/ErrorStructure.md) structure  |

For more information about `if`, `timeout`, `errorcall`, `errorjump`, `errormessage` and `errorresult` arguments, see [Common Arguments](G1ANT.Manual/appendices/common-arguments.md) page.

## Example

In the following script a new message with two attachments (be sure to provide real filepaths there) is created. To send it, use the [`outlook.send`](G1ANT.Addon/G1ANT.Addon.MSOffice/G1ANT.Addon.MSOffice/Commands/OutlookSendCommand.md) command.

```G1ANT
outlook.open
outlook.newmessage to hi@g1ant.com subject Great! body ‴Your robot rules!‴ attachments ‴D:\Files\Text.txt❚D:\Files\sheet.xlsx‴
outlook.send
outlook.close
```


