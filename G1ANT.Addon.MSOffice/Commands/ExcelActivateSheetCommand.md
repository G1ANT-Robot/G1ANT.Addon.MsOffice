# excel.activatesheet

## Syntax

```G1ANT
excel.activatesheet name ⟦text⟧
```

## Description

This command activates a specified sheet in the currently active Excel instance.

| Argument | Type | Required | Default Value | Description |
| -------- | ---- | -------- | ------------- | ----------- |
|`name`| [text](G1ANT.Language/G1ANT.Language/Structures/TextStructure.md) | yes |  | Name of a sheet to be activated |
| `if`           | [bool](G1ANT.Language/G1ANT.Language/Structures/BooleanStructure.md) | no       | true                                                        | Executes the command only if a specified condition is true   |
| `timeout`      | [timespan](G1ANT.Language/G1ANT.Language/Structures/TimeSpanStructure.md) | no       | [♥timeoutcommand](G1ANT.Language/G1ANT.Addon.Core/Variables/TimeoutCommandVariable.md) | Specifies time in milliseconds for G1ANT.Robot to wait for the command to be executed |
| `errorcall`    | [procedure](G1ANT.Language/G1ANT.Language/Structures/ProcedureStructure.md) | no       |                                                             | Name of a procedure to call when the command throws an exception or when a given `timeout` expires |
| `errorjump`    | [label](G1ANT.Language/G1ANT.Language/Structures/LabelStructure.md) | no       |                                                             | Name of the label to jump to when the command throws an exception or when a given `timeout` expires |
| `errormessage` | [text](G1ANT.Language/G1ANT.Language/Structures/TextStructure.md) | no       |                                                             | A message that will be shown in case the command throws an exception or when a given `timeout` expires, and no `errorjump` argument is specified |
| `errorresult`  | [variable](G1ANT.Language/G1ANT.Language/Structures/VariableStructure.md) | no       |                                                             | Name of a variable that will store the returned exception. The variable will be of [error](G1ANT.Language/G1ANT.Language/Structures/ErrorStructure.md) structure  |

For more information about `if`, `timeout`, `errorcall`, `errorjump`, `errormessage` and `errorresult` arguments, see [Common Arguments](G1ANT.Manual/appendices/common-arguments.md) page.

## Example

This example shows how the Excel commands work together: the `excel.open` command opens a new instance of Excel, then the `excel.addsheet` command adds a sheet named *New Sheet*, and the `excel.setvalue` command inserts *some entry* text into the cell at row 1, column 1. Then, similar actions are performed with a new sheet named *Another Sheet*. Finally, the `excel.activatesheet` command allows returning to the first sheet:

```G1ANT
excel.open
excel.addsheet ‴New Sheet‴
excel.setvalue ‴some entry‴ row 1 colindex 1
excel.addsheet ‴Another Sheet‴
excel.setvalue 132 row 1 colname ‴A‴
excel.activatesheet ‴New Sheet‴
```
