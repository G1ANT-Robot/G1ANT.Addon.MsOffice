# excel.getvalue

## Syntax

```G1ANT
excel.getvalue row ⟦integer⟧ colindex ⟦integer⟧
```

or

```G1ANT
excel.getvalue row ⟦integer⟧ colname ⟦text⟧
```

## Description

This command gets a value from a specified cell.

| Argument | Type | Required | Default Value | Description |
| -------- | ---- | -------- | ------------- | ----------- |
|`row`| [integer](G1ANT.Language/G1ANT.Language/Structures/IntegerStructure.md) | yes |  | Cell's row number |
|`colindex` or `colname`| [integer](G1ANT.Language/G1ANT.Language/Structures/IntegerStructure.md)  or [text](G1ANT.Language/G1ANT.Language/Structures/TextStructure.md) | yes |  | `colindex`: cell's column number; `colname`: cell's column name |
| `result`       | [variable](G1ANT.Language/G1ANT.Language/Structures/VariableStructure.md) | no       | `♥result`                                                   | Name of a variable where the command's result will be stored |
| `if`           | [bool](G1ANT.Language/G1ANT.Language/Structures/BooleanStructure.md) | no       | true                                                        | Executes the command only if a specified condition is true   |
| `timeout`      | [timespan](G1ANT.Language/G1ANT.Language/Structures/TimeSpanStructure.md) | no       | [♥timeoutcommand](G1ANT.Language/G1ANT.Addon.Core/Variables/TimeoutCommandVariable.md) | Specifies time in milliseconds for G1ANT.Robot to wait for the command to be executed |
| `errorcall`    | [procedure](G1ANT.Language/G1ANT.Language/Structures/ProcedureStructure.md) | no       |                                                             | Name of a procedure to call when the command throws an exception or when a given `timeout` expires |
| `errorjump`    | [label](G1ANT.Language/G1ANT.Language/Structures/LabelStructure.md) | no       |                                                             | Name of the label to jump to when the command throws an exception or when a given `timeout` expires |
| `errormessage` | [text](G1ANT.Language/G1ANT.Language/Structures/TextStructure.md) | no       |                                                             | A message that will be shown in case the command throws an exception or when a given `timeout` expires, and no `errorjump` argument is specified |
| `errorresult`  | [variable](G1ANT.Language/G1ANT.Language/Structures/VariableStructure.md) | no       |                                                             | Name of a variable that will store the returned exception. The variable will be of [error](G1ANT.Language/G1ANT.Language/Structures/ErrorStructure.md) structure  |

For more information about `if`, `timeout`, `errorcall`, `errorjump`, `errormessage` and `errorresult` arguments, see [Common Arguments](G1ANT.Manual/appendices/common-arguments.md) page.

## Example

The script below demonstrates three different methods of entering data into an Excel sheet and reading the value of a specified cell. First, the robot opens an empty Excel sheet, fills the first cell with text sent with the `keyboard` command, then sets two cells’ values in the same column specified both by the `colindex` and the `colname` arguments. Finally, the robot reads the value from the very first cell and displays it in a dialog box:

```G1ANT
excel.open
window ✱Excel
keyboard ‴Remember, remember!‴
excel.setvalue value ‴The fifth of November‴ row 2 colindex 1
excel.setvalue value ‴The Gunpowder treason and plot‴ row 3 colname A
excel.getvalue row 1 colname A result ♥guy
dialog ♥guy
```
