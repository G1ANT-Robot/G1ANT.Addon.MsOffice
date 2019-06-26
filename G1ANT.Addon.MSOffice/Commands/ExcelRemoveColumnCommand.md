# excel.removecolumn

## Syntax

```G1ANT
excel.removecolumn colindex ⟦integer⟧
```

or

```G1ANT
excel.removecolumn colname ⟦text⟧
```

## Description

This command removes the specified column.

| Argument                | Type                                                         | Required | Default Value                                                | Description                                                  |
| ----------------------- | ------------------------------------------------------------ | -------- | ------------------------------------------------------------ | ------------------------------------------------------------ |
| `colindex` or `colname` | [integer](G1ANT.Language/G1ANT.Language/Structures/IntegerStructure.md)  or [text](G1ANT.Language/G1ANT.Language/Structures/TextStructure.md) | yes      |                                                              | `colindex`: cell's column number; `colname`: cell's column name |
| `if`                    | [bool](G1ANT.Language/G1ANT.Language/Structures/BooleanStructure.md) | no       | true                                                         | Executes the command only if a specified condition is true   |
| `timeout`               | [timespan](G1ANT.Language/G1ANT.Language/Structures/TimeSpanStructure.md) | no       | [♥timeoutcommand](G1ANT.Language/G1ANT.Addon.Core/Variables/TimeoutCommandVariable.md) | Specifies time in milliseconds for G1ANT.Robot to wait for the command to be executed |
| `errorcall`             | [procedure](G1ANT.Language/G1ANT.Language/Structures/ProcedureStructure.md) | no       |                                                              | Name of a procedure to call when the command throws an exception or when a given `timeout` expires |
| `errorjump`             | [label](G1ANT.Language/G1ANT.Language/Structures/LabelStructure.md) | no       |                                                              | Name of the label to jump to when the command throws an exception or when a given `timeout` expires |
| `errormessage`          | [text](G1ANT.Language/G1ANT.Language/Structures/TextStructure.md) | no       |                                                              | A message that will be shown in case the command throws an exception or when a given `timeout` expires, and no `errorjump` argument is specified |
| `errorresult`           | [variable](G1ANT.Language/G1ANT.Language/Structures/VariableStructure.md) | no       |                                                              | Name of a variable that will store the returned exception. The variable will be of [error](G1ANT.Language/G1ANT.Language/Structures/ErrorStructure.md) structure |

For more information about `if`, `timeout`, `errorcall`, `errorjump`, `errormessage` and `errorresult` arguments, see [Common Arguments](G1ANT.Manual/appendices/common-arguments.md) page.https://github.com/G1ANT-Robot/G1ANT.Addon.MSOffice)

## Example

The following script open an empty Excel sheet, fills three cells in a row with text, then inserts a new column as column B and removes the first column (A):

```G1ANT
excel.open
window ✱Excel
keyboard one⋘RIGHT⋙two⋘RIGHT⋙three⋘RIGHT⋙
excel.insertcolumn colname A where after
excel.removecolumn colname A
```
