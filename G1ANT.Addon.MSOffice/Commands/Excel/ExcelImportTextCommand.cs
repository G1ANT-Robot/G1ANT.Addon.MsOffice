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

namespace G1ANT.Addon.MSOffice
{
    [Command(Name = "excel.importtext", Tooltip = "This command establishes a data connection between a text file and the specified destination in an active sheet and imports data into it")]
    public class ExcelImportTextCommand : Command
	{
        public class Arguments : CommandArguments
        {
            [Argument(Required = true, Tooltip = "Path to a text file to be imported (csv data format is supported)")]
            public TextStructure Path { get; set; }

            [Argument(Tooltip = "Starting cell (top left) for the imported data, specified either as text or a point")]
            public Structure Destination { get; set; } = new TextStructure("A1");

            [Argument(Tooltip = "Delimiter (data separator) to be used while importing data: `tab`, `semicolon`, `comma`, `space` or any other character")]
            public TextStructure Delimiter { get; set; } = new TextStructure("semicolon");

            [Argument(Tooltip = "Name of a range where data will be placed")]
            public TextStructure Name { get; set; }

            [Argument(Tooltip = "Name of a variable that will store the total number of rows of the imported data")]
            public TextStructure ResultRows { get; set; } = new TextStructure("resultrows");

            [Argument(Tooltip = "Name of a variable that will store the total number of columns of the imported data")]
            public TextStructure ResultColumns { get; set; } = new TextStructure("resultcolumns");
        }
        public ExcelImportTextCommand(AbstractScripter scripter) : base(scripter)
        {
        }

        public void Execute(Arguments arguments)
        {
            try
            {
                object destination = null;
                if (arguments.Destination is TextStructure)
                {
                    destination = (arguments.Destination as TextStructure).Value;
                }
                else if (arguments.Destination is PointStructure)
                {
                    destination = (arguments.Destination as PointStructure).Value;
                }
                else
                {
                    throw new ArgumentException("Wrong destination argument. It accepts either String or Point value.");
                }
                int columnsCount = 0;
                int rowsCount = 0;
                ExcelManager.CurrentExcel.ImportTextFile(arguments.Path.Value, destination, arguments?.Name?.Value, arguments.Delimiter.Value, out rowsCount, out columnsCount);
                Scripter.Variables.SetVariableValue(arguments.ResultColumns.Value, new Language.IntegerStructure(columnsCount));
                Scripter.Variables.SetVariableValue(arguments.ResultRows.Value, new Language.IntegerStructure(rowsCount));
            }
            catch (Exception ex)
            {
                throw new ApplicationException($"Problem occured while importing text data. Path: '{arguments.Path.Value}'. Message: '{ex.Message}'", ex);
            }
        }
    }
}
