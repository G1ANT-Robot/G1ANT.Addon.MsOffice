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
using System.Runtime.InteropServices;

namespace G1ANT.Addon.MSOffice
{
    [Command(Name = "excel.open", Tooltip = "This command opens a new Excel instance")]
    public class ExcelOpenCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Tooltip = "Path of a file that has to be opened in Excel; if not specified, Excel will be opened with an empty sheet")]
            public TextStructure Path { get; set; }

            [Argument(Tooltip = "Specifies whether Excel should be opened in the background")]
            public BooleanStructure InBackground { get; set; } = new BooleanStructure(false);

            [Argument(Tooltip = "Name of a sheet to be activated")]
            public TextStructure Sheet { get; set; }

            [Argument(Tooltip = "Name of a variable where a currently opened Excel process number is stored. It can be used in the `excel.switch` command")]
            public VariableStructure Result { get; set; } = new VariableStructure("result");
        }
        public ExcelOpenCommand(AbstractScripter scripter) : base(scripter)
        {
        }

        public void Execute(Arguments arguments)
        {
            try
            {
                ExcelWrapper excelWrapper = ExcelManager.CreateInstance();
                excelWrapper.Open(arguments.Path?.Value, arguments.Sheet?.Value, !arguments.InBackground.Value);
                Scripter.Variables.SetVariableValue(arguments.Result.Value, new IntegerStructure(excelWrapper.Id));
            }
            catch (Exception ex)
            {
                if (ex.GetType() == typeof(COMException) && ex.Message.Contains("80040154"))
                    throw new Exception("Could not find Microsoft Office on computer. Please make sure it is installed and try again.");
                throw new ApplicationException($"Problem occured while opening excel instance. Path: '{arguments.Path?.Value}', Sheet: '{arguments.Sheet?.Value}', InBackground: '{arguments.InBackground.Value}'", ex);
            }           
        }
    }
}