# word.replace

## Syntax

```G1ANT
word.replace from ⟦text⟧ to ⟦text⟧ matchcase ⟦bool⟧ wholewords ⟦bool⟧
```

## Description

This command replaces specified text in a document.

| Argument | Type | Required | Default Value | Description |
| -------- | ---- | -------- | ------------- | ----------- |
|`from`| [text](G1ANT.Language/G1ANT.Language/Structures/TextStructure.md) | yes |  |Text to be found in a document|
|`to`| [text](G1ANT.Language/G1ANT.Language/Structures/TextStructure.md) | yes |  | Text to be replaced in a document |
|`matchcase`| [bool](G1ANT.Language/G1ANT.Language/Structures/BooleanStructure.md) | no |false | If set to `true`, then the search is case sensitive |
|`wholewords`| [bool](G1ANT.Language/G1ANT.Language/Structures/BooleanStructure.md) | no | false | If set to `true`, only whole words are replaced |
| `if`           | [bool](G1ANT.Language/G1ANT.Language/Structures/BooleanStructure.md) | no       | true                                                        | Executes the command only if a specified condition is true   |
| `timeout`      | [timespan](G1ANT.Language/G1ANT.Language/Structures/TimeSpanStructure.md) | no       | [♥timeoutcommand](G1ANT.Language/G1ANT.Addon.Core/Variables/TimeoutCommandVariable.md) | Specifies time in milliseconds for G1ANT.Robot to wait for the command to be executed |
| `errorcall`    | [procedure](G1ANT.Language/G1ANT.Language/Structures/ProcedureStructure.md) | no       |                                                             | Name of a procedure to call when the command throws an exception or when a given `timeout` expires |
| `errorjump`    | [label](G1ANT.Language/G1ANT.Language/Structures/LabelStructure.md) | no       |                                                             | Name of the label to jump to when the command throws an exception or when a given `timeout` expires |
| `errormessage` | [text](G1ANT.Language/G1ANT.Language/Structures/TextStructure.md) | no       |                                                             | A message that will be shown in case the command throws an exception or when a given `timeout` expires, and no `errorjump` argument is specified |
| `errorresult`  | [variable](G1ANT.Language/G1ANT.Language/Structures/VariableStructure.md) | no       |                                                             | Name of a variable that will store the returned exception. The variable will be of [error](G1ANT.Language/G1ANT.Language/Structures/ErrorStructure.md) structure  |

For more information about `if`, `timeout`, `errorcall`, `errorjump`, `errormessage` and `errorresult` arguments, see [Common Arguments](G1ANT.Manual/appendices/common-arguments.md) page.

## Example

```G1ANT
♥toInsert = ‴I hate yogurt. It's just stuff with bits in. All I've got to do is pass as an ordinary human being. Simple. What could possibly go wrong? Saving the world with meals on wheels.‴
word.open
word.inserttext ♥toInsert
word.replace Saving to Killing wholewords true
```

The script above opens a blank Word document, inserts text to it and then replaces a word in this text, matching the whole word “saving” and changing it into “killing”.

