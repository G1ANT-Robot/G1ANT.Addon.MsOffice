# word.export

## Syntax

```G1ANT
word.export path ⟦text⟧ type ⟦text⟧
```

## Description

This command exports a document from the currently active Word instance to a specified file in either .pdf or .xps format.

| Argument | Type | Required | Default Value | Description |
| -------- | ---- | -------- | ------------- | ----------- |
|`path`| [text](G1ANT.Language/G1ANT.Language/Structures/TextStructure.md) | no |  | Path to the exported file; if not specified, the file will be saved in the location of the source file |
|`type`| [text](G1ANT.Language/G1ANT.Language/Structures/TextStructure.md) | no | | Type of the exported file: `pdf` or `xps`); if not specified, the type will be defined by the extension of the exported filename |
| `if`           | [bool](G1ANT.Language/G1ANT.Language/Structures/BooleanStructure.md) | no       | true                                                        | Executes the command only if a specified condition is true   |
| `timeout`      | [timespan](G1ANT.Language/G1ANT.Language/Structures/TimeSpanStructure.md) | no       | [♥timeoutcommand](G1ANT.Language/G1ANT.Addon.Core/Variables/TimeoutCommandVariable.md) | Specifies time in milliseconds for G1ANT.Robot to wait for the command to be executed |
| `errorcall`    | [procedure](G1ANT.Language/G1ANT.Language/Structures/ProcedureStructure.md) | no       |                                                             | Name of a procedure to call when the command throws an exception or when a given `timeout` expires |
| `errorjump`    | [label](G1ANT.Language/G1ANT.Language/Structures/LabelStructure.md) | no       |                                                             | Name of the label to jump to when the command throws an exception or when a given `timeout` expires |
| `errormessage` | [text](G1ANT.Language/G1ANT.Language/Structures/TextStructure.md) | no       |                                                             | A message that will be shown in case the command throws an exception or when a given `timeout` expires, and no `errorjump` argument is specified |
| `errorresult`  | [variable](G1ANT.Language/G1ANT.Language/Structures/VariableStructure.md) | no       |                                                             | Name of a variable that will store the returned exception. The variable will be of [error](G1ANT.Language/G1ANT.Language/Structures/ErrorStructure.md) structure  |

For more information about `if`, `timeout`, `errorcall`, `errorjump`, `errormessage` and `errorresult` arguments, see [Common Arguments](G1ANT.Manual/appendices/common-arguments.md) page.

## Example

In this example the robot opens a Word instance with a sample file declared in the `♥sourceFile` variable (be sure to provide a real filepath there). Then, the document is exported to a pdf file on a desktop (the filepath is declared in the `♥destinationFile` variable).

```G1ANT
♥sourceFile = ♥environment⟦USERPROFILE⟧\Documents\test.docx
♥destinationFile = ♥environment⟦USERPROFILE⟧\Documents\test.pdf
word.open ♥sourceFile
word.export ♥destinationFile
```


