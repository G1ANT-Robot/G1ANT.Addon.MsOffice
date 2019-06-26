# excel.switch

## Syntax

```G1ANT
excel.switch id ⟦integer⟧
```

## Description

This command switches to another Excel instance opened by G1ANT.Robot.

| Argument | Type | Required | Default Value | Description |
| -------- | ---- | -------- | ------------- | ----------- |
|`id`| [integer](G1ANT.Language/G1ANT.Language/Structures/IntegerStructure.md) | yes |  | ID number of an Excel instance that will be activated |
| `if`           | [bool](G1ANT.Language/G1ANT.Language/Structures/BooleanStructure.md) | no       | true                                                        | Executes the command only if a specified condition is true   |
| `timeout`      | [timespan](G1ANT.Language/G1ANT.Language/Structures/TimeSpanStructure.md) | no       | [♥timeoutcommand](G1ANT.Language/G1ANT.Addon.Core/Variables/TimeoutCommandVariable.md) | Specifies time in milliseconds for G1ANT.Robot to wait for the command to be executed |
| `errorcall`    | [procedure](G1ANT.Language/G1ANT.Language/Structures/ProcedureStructure.md) | no       |                                                             | Name of a procedure to call when the command throws an exception or when a given `timeout` expires |
| `errorjump`    | [label](G1ANT.Language/G1ANT.Language/Structures/LabelStructure.md) | no       |                                                             | Name of the label to jump to when the command throws an exception or when a given `timeout` expires |
| `errormessage` | [text](G1ANT.Language/G1ANT.Language/Structures/TextStructure.md) | no       |                                                             | A message that will be shown in case the command throws an exception or when a given `timeout` expires, and no `errorjump` argument is specified |
| `errorresult`  | [variable](G1ANT.Language/G1ANT.Language/Structures/VariableStructure.md) | no       |                                                             | Name of a variable that will store the returned exception. The variable will be of [error](G1ANT.Language/G1ANT.Language/Structures/ErrorStructure.md) structure  |

For more information about `if`, `timeout`, `errorcall`, `errorjump`, `errormessage` and `errorresult` arguments, see [Common Arguments](G1ANT.Manual/appendices/common-arguments.md) page.

## Example

In order to use the `id` argument for the `excel.switch` command, you need to store an Excel instance ID in a variable while using the `excel.open` command. The following example opens an empty Excel sheet, then an Excel file located on your Desktop. Both instances have their IDs stored in their respective `♥excel1` and `♥excel2` variables. Finally, the first ID is used to switch to the first Excel instance:

```G1ANT
excel.open result ♥excel1
excel.open ♥environment⟦USERPROFILE⟧\Desktop\test.xlsx result ♥excel2
excel.switch id ♥excel1
```

ID numbers start with 0 (zero), so switching to the Excel instance with `id 0` means that the robot activates the first Excel instance opened by a script.

