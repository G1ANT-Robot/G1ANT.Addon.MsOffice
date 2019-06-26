/**
*    Copyright(C) G1ANT Ltd, All rights reserved
*    Solution G1ANT.Addon, Project G1ANT.Addon.MSOffice
*    www.g1ant.com
*
*    Licensed under the G1ANT license.
*    See License.txt file in the project root for full license information.
*
*/

using G1ANT.Language;



using System;
using System.Collections.Generic;

namespace G1ANT.Addon.MSOffice
{
    [Command(Name = "excel.runvbcode", Tooltip = "This command runs a Visual Basic macro code in the currently active Excel instance")]
    public class ExcelRunVBCodeCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Required = true, Tooltip = "Visual Basic code of a macro that will be run")]
            public TextStructure Code { get; set; }

            [Argument]
            public VariableStructure Result { get; set; } = new VariableStructure("result");
        }

        public ExcelRunVBCodeCommand(AbstractScripter scripter) : base(scripter)
        {
        }

        public void Execute(Arguments arguments)
        {
            try
            {

                var result = ExcelManager.CurrentExcel.RunMacroCode(arguments.Code.Value);
                try
                {
                    Scripter.Variables.SetVariableValue(arguments.Result.Value, Scripter.Structures.CreateStructure(result));
                }
                catch
                {
                    Scripter.Variables.SetVariableValue(arguments.Result.Value, Scripter.Structures.CreateStructure(result.ToString()));
                }
            }
            catch (Exception ex)
            {
                throw new ApplicationException($"Problem occured while running excel macro code. Code: '{arguments.Code.Value}'", ex);
            }
        }
    }
}
