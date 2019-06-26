# outlook.send

## Syntax

```G1ANT
outlook.send
```

## Description

This command sends an email drafted with the `outlook.newmessage` command or other mail (such as a reply drafted with the `outlook.reply` command) stored in a variable of [outlookmail](G1ANT.Addon/G1ANT.Addon.MSOffice/G1ANT.Addon.MSOffice/Structures/OutlookMailStructure.md) structure.

| Argument | Type | Required | Default Value | Description |
| -------- | ---- | -------- | ------------- | ----------- |
| `mail`       | [outlookmail](G1ANT.Addon/G1ANT.Addon.MSOffice/G1ANT.Addon.MSOffice/Structures/OutlookMailStructure.md) | no       |                                                    | Name of an `outlookmail` variable where a mail to be sent is stored |
| `result`       | [variable](G1ANT.Language/G1ANT.Language/Structures/VariableStructure.md) | no       | `♥result`                                                   | Name of a variable where the command's result will be stored |
| `if`           | [bool](G1ANT.Language/G1ANT.Language/Structures/BooleanStructure.md) | no       | true                                                        | Executes the command only if a specified condition is true   |
| `timeout`      | [timespan](G1ANT.Language/G1ANT.Language/Structures/TimeSpanStructure.md) | no       | [♥timeoutcommand](G1ANT.Language/G1ANT.Addon.Core/Variables/TimeoutCommandVariable.md) | Specifies time in milliseconds for G1ANT.Robot to wait for the command to be executed |
| `errorcall`    | [procedure](G1ANT.Language/G1ANT.Language/Structures/ProcedureStructure.md) | no       |                                                             | Name of a procedure to call when the command throws an exception or when a given `timeout` expires |
| `errorjump`    | [label](G1ANT.Language/G1ANT.Language/Structures/LabelStructure.md) | no       |                                                             | Name of the label to jump to when the command throws an exception or when a given `timeout` expires |
| `errormessage` | [text](G1ANT.Language/G1ANT.Language/Structures/TextStructure.md) | no       |                                                             | A message that will be shown in case the command throws an exception or when a given `timeout` expires, and no `errorjump` argument is specified |
| `errorresult`  | [variable](G1ANT.Language/G1ANT.Language/Structures/VariableStructure.md) | no       |                                                             | Name of a variable that will store the returned exception. The variable will be of [error](G1ANT.Language/G1ANT.Language/Structures/ErrorStructure.md) structure  |

For more information about `if`, `timeout`, `errorcall`, `errorjump`, `errormessage` and `errorresult` arguments, see [Common Arguments](G1ANT.Manual/appendices/common-arguments.md) page.

## Example

In this example, the robot sends a simple message created with the `outlook.newmessage` command:

```G1ANT
outlook.open
outlook.newmessage to hi@g1ant.com subject ‴Hi, Robot!‴  body ‴Impressed with your work!‴
outlook.send
outlook.close
```
