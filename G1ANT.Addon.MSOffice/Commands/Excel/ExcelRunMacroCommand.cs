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
    [Command(Name = "excel.runmacro", Tooltip = "This command runs a macro in the currently active Excel instance")]
    public class ExcelRunMacroCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Required = true, Tooltip = "Name of a macro that is defined in a workbook")]
            public TextStructure Name { get; set; }

            [Argument(Tooltip = "Comma-separated arguments that will be passed to a macro")]
            public ListStructure Args { get; set; }

            [Argument(Tooltip = "Name of a variable where the macro's return value will be stored")]
            public VariableStructure Result { get; set; } = new VariableStructure("result");
        }

        public ExcelRunMacroCommand(AbstractScripter scripter) : base(scripter)
        {
        }

        public void Execute(Arguments arguments)
        {
            try
            {
                List<object> args = new List<object>();

                if (arguments.Args?.Value != null)
                {
                    foreach (var arg in arguments.Args?.Value)
                    {
                        if (arg is TextStructure textStruc)
                        {
                            args.Add(textStruc.Value);
                        }
                        else if (arg is string str)
                        {
                            args.Add(str);
                        }
                    }
                }
                //else
                //{
                //    args.Add(string.Empty);
                //}

                var result = ExcelManager.CurrentExcel.RunMacro(arguments.Name.Value, args);
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
                throw new ApplicationException($"Problem occured while running excel macro. Path: '{arguments.Name.Value}', Arguments count: '{arguments.Args?.Value?.Count ?? 0}'", ex);
            }
        }
    }
}
