# excel.getrow

## Syntax

```G1ANT
excel.getrow row ⟦integer⟧
```

## Description

This command gets all used cells of the specified row.

| Argument | Type | Required | Default Value | Description |
| -------- | ---- | -------- | ------------- | ----------- |
|`row`| [integer](G1ANT.Language/G1ANT.Language/Structures/IntegerStructure.md) | yes |  | Row number |
| `result`       | [variable](G1ANT.Language/G1ANT.Language/Structures/VariableStructure.md) | no       | `♥result`                                                   | Name of a variable where the command's result will be stored |
| `if`           | [bool](G1ANT.Language/G1ANT.Language/Structures/BooleanStructure.md) | no       | true                                                        | Executes the command only if a specified condition is true   |
| `timeout`      | [timespan](G1ANT.Language/G1ANT.Language/Structures/TimeSpanStructure.md) | no       | [♥timeoutcommand](G1ANT.Language/G1ANT.Addon.Core/Variables/TimeoutCommandVariable.md) | Specifies time in milliseconds for G1ANT.Robot to wait for the command to be executed |
| `errorcall`    | [procedure](G1ANT.Language/G1ANT.Language/Structures/ProcedureStructure.md) | no       |                                                             | Name of a procedure to call when the command throws an exception or when a given `timeout` expires |
| `errorjump`    | [label](G1ANT.Language/G1ANT.Language/Structures/LabelStructure.md) | no       |                                                             | Name of the label to jump to when the command throws an exception or when a given `timeout` expires |
| `errormessage` | [text](G1ANT.Language/G1ANT.Language/Structures/TextStructure.md) | no       |                                                             | A message that will be shown in case the command throws an exception or when a given `timeout` expires, and no `errorjump` argument is specified |
| `errorresult`  | [variable](G1ANT.Language/G1ANT.Language/Structures/VariableStructure.md) | no       |                                                             | Name of a variable that will store the returned exception. The variable will be of [error](G1ANT.Language/G1ANT.Language/Structures/ErrorStructure.md) structure  |

For more information about `if`, `timeout`, `errorcall`, `errorjump`, `errormessage` and `errorresult` arguments, see [Common Arguments](G1ANT.Manual/appendices/common-arguments.md) page.

## Example

The following script opens Excel with an empty sheet, then fills first cells in consecutive rows with text. Finally it copies the content of the third row and displays it in a dialog box:

```G1ANT
excel.open
window ✱Excel
keyboard ‴THE wild bee reels from bough to bough ⋘DOWN⋙With his furry coat and his gauzy wing. ⋘DOWN⋙Now in a lily-cup, and now THE wild bee reels from bough to bough ⋘DOWN⋙Setting a jacinth bell a-swing,  ⋘DOWN⋙In his wandering; ⋘DOWN⋙Sit closer love: it was here I trow ⋘DOWN⋙I made that vow⋘DOWN⋙With his furry coat and his gauzy wing.‴
keyboard ⋘ENTER⋙
excel.getrow 3 result ♥rowInput
dialog ♥rowInput
```

