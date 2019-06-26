# excel.runvbcode

## Syntax

```G1ANT
excel.runvbcode code ⟦text⟧
```

## Description

This command runs a Visual Basic macro code in the currently active Excel instance. The code can contain only procedures (Sub blocks), not functions.

| Argument       | Type                                                         | Required | Default Value                                                | Description                                                  |
| -------------- | ------------------------------------------------------------ | -------- | ------------------------------------------------------------ | ------------------------------------------------------------ |
| `code`         | [text](G1ANT.Language/G1ANT.Language/Structures/TextStructure.md) | yes      |                                                              | Visual Basic code of a macro that will be run                |
| `if`           | [bool](G1ANT.Language/G1ANT.Language/Structures/BooleanStructure.md) | no       | true                                                         | Executes the command only if a specified condition is true   |
| `timeout`      | [timespan](G1ANT.Language/G1ANT.Language/Structures/TimeSpanStructure.md) | no       | [♥timeoutcommand](G1ANT.Language/G1ANT.Addon.Core/Variables/TimeoutCommandVariable.md) | Specifies time in milliseconds for G1ANT.Robot to wait for the command to be executed |
| `errorcall`    | [procedure](G1ANT.Language/G1ANT.Language/Structures/ProcedureStructure.md) | no       |                                                              | Name of a procedure to call when the command throws an exception or when a given `timeout` expires |
| `errorjump`    | [label](G1ANT.Language/G1ANT.Language/Structures/LabelStructure.md) | no       |                                                              | Name of the label to jump to when the command throws an exception or when a given `timeout` expires |
| `errormessage` | [text](G1ANT.Language/G1ANT.Language/Structures/TextStructure.md) | no       |                                                              | A message that will be shown in case the command throws an exception or when a given `timeout` expires, and no `errorjump` argument is specified |
| `errorresult`  | [variable](G1ANT.Language/G1ANT.Language/Structures/VariableStructure.md) | no       |                                                              | Name of a variable that will store the returned exception. The variable will be of [error](G1ANT.Language/G1ANT.Language/Structures/ErrorStructure.md) structure |

For more information about `if`, `timeout`, `errorcall`, `errorjump`, `errormessage` and `errorresult` arguments, see [Common Arguments](G1ANT.Manual/appendices/common-arguments.md) page.

## Example

Suppose you want to run this Visual Basic macro code in Excel:

```visual basic
Sub Multiplication()
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]*2"
    Range("B2").Select
    Selection.AutoFill Destination:=Range("B2:B6"), Type:=xlFillDefault
    Range("B2:B6").Select
End Sub 
```

To use it with the `excel.runvbcode` command, you just skip the first and the last lines of the code (`Sub…` and `End Sub`) and separate remaining lines with this C# snippet that recreates the new line character: `⊂"\r\n"⊃`. The resulting script should look like this:

```G1ANT
excel.runvbcode ‴Range("B2").Select⊂"\r\n"⊃ActiveCell.FormulaR1C1 = "=RC[-1]*2"⊂"\r\n"⊃Range("B2").Select⊂"\r\n"⊃Selection.AutoFill Destination:=Range("B2:B6"), Type:=xlFillDefault⊂"\r\n"⊃Range("B2:B6").Select‴
```

> **Note:** In order to use this command, an access to VBA project object model must be granted in Excel. For more details, click [here](https://www.spreadsheet1.com/trust-access-to-the-vba-project-object-model.html).
