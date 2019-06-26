# excel.open

## Syntax

```G1ANT
excel.open path ⟦text⟧ inbackground ⟦bool⟧ sheet ⟦text⟧
```

## Description

This command opens a new Excel instance.

| Argument | Type | Required | Default Value | Description |
| -------- | ---- | -------- | ------------- | ----------- |
|`path`| [text](G1ANT.Language/G1ANT.Language/Structures/TextStructure.md) | no |  | Path of a file that has to be opened in Excel; if not specified, Excel will be opened with an empty sheet |
|`inbackground`| [bool](G1ANT.Language/G1ANT.Language/Structures/BooleanStructure.md) | no | false | Specifies whether Excel should be opened in the background |
|`sheet`| [text](G1ANT.Language/G1ANT.Language/Structures/TextStructure.md) | no |  | Name of a sheet to be activated |
|`result`| [variable](https://github.com/G1ANT-Robot/G1ANT.Manual/blob/master/G1ANT-Language/Special-Characters/variable.md) | no | ♥result  | Name of a variable where a currently opened Excel process number is stored. It can be used in the `excel.switch` command |
| `if`           | [bool](G1ANT.Language/G1ANT.Language/Structures/BooleanStructure.md) | no       | true                                                        | Executes the command only if a specified condition is true   |
| `timeout`      | [timespan](G1ANT.Language/G1ANT.Language/Structures/TimeSpanStructure.md) | no       | [♥timeoutcommand](G1ANT.Language/G1ANT.Addon.Core/Variables/TimeoutCommandVariable.md) | Specifies time in milliseconds for G1ANT.Robot to wait for the command to be executed |
| `errorcall`    | [procedure](G1ANT.Language/G1ANT.Language/Structures/ProcedureStructure.md) | no       |                                                             | Name of a procedure to call when the command throws an exception or when a given `timeout` expires |
| `errorjump`    | [label](G1ANT.Language/G1ANT.Language/Structures/LabelStructure.md) | no       |                                                             | Name of the label to jump to when the command throws an exception or when a given `timeout` expires |
| `errormessage` | [text](G1ANT.Language/G1ANT.Language/Structures/TextStructure.md) | no       |                                                             | A message that will be shown in case the command throws an exception or when a given `timeout` expires, and no `errorjump` argument is specified |
| `errorresult`  | [variable](G1ANT.Language/G1ANT.Language/Structures/VariableStructure.md) | no       |                                                             | Name of a variable that will store the returned exception. The variable will be of [error](G1ANT.Language/G1ANT.Language/Structures/ErrorStructure.md) structure  |

For more information about `if`, `timeout`, `errorcall`, `errorjump`, `errormessage` and `errorresult` arguments, see [Common Arguments](G1ANT.Manual/appendices/common-arguments.md) page.

## Example

This example shows how Excel can be opened in the background, so that you will not notice any action, but G1ANT.Robot will execute the script anyway. You can see the results in the `test.xlsx` file on your Desktop:

```G1ANT
excel.open inbackground true
excel.setvalue ‴Random Text‴ row 1 colname A
excel.save ♥environment⟦USERPROFILE⟧\Desktop\test.xlsx
excel.close
```

