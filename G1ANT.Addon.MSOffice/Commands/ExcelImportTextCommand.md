# excel.importtext

## Syntax

```G1ANT
excel.importtext path ⟦text⟧ destination ⟦text⟧ delimiter ⟦text⟧ name ⟦text⟧ resultrows ⟦variable⟧ resultcolumns ⟦variable⟧
```

## Description

This command establishes a data connection between a text file and the specified destination in an active sheet and imports data into it.

| Argument | Type | Required | Default Value | Description |
| -------- | ---- | -------- | ------------- | ----------- |
|`path`| [text](G1ANT.Language/G1ANT.Language/Structures/TextStructure.md) | yes | | Path to a text file to be imported (csv data format is supported) |
|`destination`| [text](G1ANT.Language/G1ANT.Language/Structures/TextStructure.md) | no | A1 | Starting cell (top left) for the imported data, specified either as text or a point |
|`delimiter`| [text](G1ANT.Language/G1ANT.Language/Structures/TextStructure.md) | no | semicolon | Delimiter (data separator) to be used while importing data: `tab`, `semicolon`, `comma`, `space` or any other character |
|`name`| [text](G1ANT.Language/G1ANT.Language/Structures/TextStructure.md) | no|  | Name of a range where data will be placed |
|`resultrows`| [variable](G1ANT.Language/G1ANT.Language/Structures/VariableStructure.md) | no | ♥resultrows | Name of a variable that will store the total number of rows of the imported data |
|`resultcolumns`| [variable](G1ANT.Language/G1ANT.Language/Structures/VariableStructure.md) | no | ♥resultcolumns | Name of a variable that will store the total number of columns of the imported data |
| `if`           | [bool](G1ANT.Language/G1ANT.Language/Structures/BooleanStructure.md) | no       | true                                                        | Executes the command only if a specified condition is true   |
| `timeout`      | [timespan](G1ANT.Language/G1ANT.Language/Structures/TimeSpanStructure.md) | no       | [♥timeoutcommand](G1ANT.Language/G1ANT.Addon.Core/Variables/TimeoutCommandVariable.md) | Specifies time in milliseconds for G1ANT.Robot to wait for the command to be executed |
| `errorcall`    | [procedure](G1ANT.Language/G1ANT.Language/Structures/ProcedureStructure.md) | no       |                                                             | Name of a procedure to call when the command throws an exception or when a given `timeout` expires |
| `errorjump`    | [label](G1ANT.Language/G1ANT.Language/Structures/LabelStructure.md) | no       |                                                             | Name of the label to jump to when the command throws an exception or when a given `timeout` expires |
| `errormessage` | [text](G1ANT.Language/G1ANT.Language/Structures/TextStructure.md) | no       |                                                             | A message that will be shown in case the command throws an exception or when a given `timeout` expires, and no `errorjump` argument is specified |
| `errorresult`  | [variable](G1ANT.Language/G1ANT.Language/Structures/VariableStructure.md) | no       |                                                             | Name of a variable that will store the returned exception. The variable will be of [error](G1ANT.Language/G1ANT.Language/Structures/ErrorStructure.md) structure  |

For more information about `if`, `timeout`, `errorcall`, `errorjump`, `errormessage` and `errorresult` arguments, see [Common Arguments](G1ANT.Manual/appendices/common-arguments.md) page.

## Example

In the following example you have to prepare a sample, comma-delimited `data.csv` file located on your Desktop. The data will be imported and inserted in the area starting with the B1 cell. The total number of imported rows and columns is then displayed in a dialog box, and the resulting Excel sheet is saved to a file on your Desktop:

```G1ANT
excel.open
excel.importtext path ♥environment⟦USERPROFILE⟧\Desktop\data.csv destination B1 delimiter comma resultrows ♥importedRows resultcolumns ♥importedColumns
dialog ‴Rows imported: ♥importedRows; Columns imported: ♥importedColumns‴
excel.save ♥environment⟦USERPROFILE⟧\Desktop\imported_data.xlsx
excel.close
```

