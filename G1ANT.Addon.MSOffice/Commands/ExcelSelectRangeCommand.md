# excel.selectrange

## Syntax

```G1ANT
excel.selectrange colindex1 ⟦integer⟧ row1 ⟦integer⟧ colindex2 ⟦integer⟧ row2 ⟦integer⟧
```

or

```G1ANT
excel.selectrange colname1 ⟦text⟧ row1 ⟦integer⟧ colname2 ⟦text⟧ row2 ⟦integer⟧
```

> **Note:** Mixing of `colname` and `colindex` arguments for different cells is allowed, e.g.: 
>
> ```G1ANT
> excel.selectrange colindex1 ⟦integer⟧ row1 ⟦integer⟧ colname2 ⟦text⟧ row2 ⟦integer⟧
> ```

## Description

This command selects a range of cells in the currently active Excel instance.

| Argument | Type | Required | Default Value | Description |
| -------- | ---- | -------- | ------------- | ----------- |
| `colindex1` or `colname1` | [integer](https://manual.g1ant.com/link/G1ANT.Language/G1ANT.Language/Structures/IntegerStructure.md)  or [text](https://manual.g1ant.com/link/G1ANT.Language/G1ANT.Language/Structures/TextStructure.md) | yes      |                                                              | `colindex1`: starting cell's column number; `colname1`: starting cell's column name |
|`row1`| [integer](https://manual.g1ant.com/link/G1ANT.Language/G1ANT.Language/Structures/IntegerStructure.md) | yes |  | Starting cell's row number |
| `colindex2` or `colname2` | [integer](https://manual.g1ant.com/link/G1ANT.Language/G1ANT.Language/Structures/IntegerStructure.md)  or [text](https://manual.g1ant.com/link/G1ANT.Language/G1ANT.Language/Structures/TextStructure.md) | yes      |                                                              | `colindex2`: ending cell's column number; `colname2`: ending cell's column name |
|`row2`| [integer](https://manual.g1ant.com/link/G1ANT.Language/G1ANT.Language/Structures/IntegerStructure.md) | yes |  | Ending ell's row number |
| `if`           | [bool](https://manual.g1ant.com/link/G1ANT.Language/G1ANT.Language/Structures/BooleanStructure.md) | no       | true                                                        | Executes the command only if a specified condition is true   |
| `timeout`      | [timespan](https://manual.g1ant.com/link/G1ANT.Language/G1ANT.Language/Structures/TimeSpanStructure.md) | no       | [♥timeoutcommand](https://manual.g1ant.com/link/G1ANT.Addon.Core/G1ANT.Addon.Core/Variables/TimeoutCommandVariable.md) | Specifies time in milliseconds for G1ANT.Robot to wait for the command to be executed |
| `errorcall`    | [procedure](https://manual.g1ant.com/link/G1ANT.Language/G1ANT.Language/Structures/ProcedureStructure.md) | no       |                                                             | Name of a procedure to call when the command throws an exception or when a given `timeout` expires |
| `errorjump`    | [label](https://manual.g1ant.com/link/G1ANT.Language/G1ANT.Language/Structures/LabelStructure.md) | no       |                                                             | Name of the label to jump to when the command throws an exception or when a given `timeout` expires |
| `errormessage` | [text](https://manual.g1ant.com/link/G1ANT.Language/G1ANT.Language/Structures/TextStructure.md) | no       |                                                             | A message that will be shown in case the command throws an exception or when a given `timeout` expires, and no `errorjump` argument is specified |
| `errorresult`  | [variable](https://manual.g1ant.com/link/G1ANT.Language/G1ANT.Language/Structures/VariableStructure.md) | no       |                                                             | Name of a variable that will store the returned exception. The variable will be of [error](https://manual.g1ant.com/link/G1ANT.Language/G1ANT.Language/Structures/ErrorStructure.md) structure  |

For more information about `if`, `timeout`, `errorcall`, `errorjump`, `errormessage` and `errorresult` arguments, see [Common Arguments](https://manual.g1ant.com/link/G1ANT.Manual/appendices/common-arguments.md) page.

## Example

In this example the robot opens Excel, focuses on its Window in order to fill three cells in separate rows of column A with some text, then selects first two of these cells, copies their content and pastes it into two lower cells in column B:

```G1ANT
excel.open
window ✱Excel
keyboard ‴Remember, remember!⋘DOWN⋙the fifth of November⋘DOWN⋙The Gunpowder treason and plot‴
excel.selectrange colindex1 1 row1 1 colname2 A row2 2
excel.copy
excel.selectrange colname1 B row1 2 colindex2 2 row2 3
excel.paste
```

