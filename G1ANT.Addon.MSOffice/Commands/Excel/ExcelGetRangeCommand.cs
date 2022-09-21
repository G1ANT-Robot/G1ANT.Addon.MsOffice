/**
*    Copyright(C) G1ANT Ltd, All rights reserved
*    Solution G1ANT.Addon, Project G1ANT.Addon.MSOffice
*    www.g1ant.com
*
*    Licensed under the G1ANT license.
*    See License.txt file in the project root for full license information.
*
*/
using System;

using G1ANT.Language;

namespace G1ANT.Addon.MSOffice.Commands.Excel
{
    [Command(Name = "excel.getrange", Tooltip = "This command gets a value from a specified range")]
    public class ExcelGetRangeCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Required = true, Tooltip = "Starting column's index")]
            public IntegerStructure ColStartIndex { get; set; }

            [Argument(Required = true, Tooltip = "Ending column's index")]
            public IntegerStructure ColEndIndex { get; set; }

            [Argument(Required = true, Tooltip = "Starting row's index")]
            public IntegerStructure RowStartIndex { get; set; }

            [Argument(Required = true, Tooltip = "Ending row's index")]
            public IntegerStructure RowEndIndex { get; set; }

            [Argument(Tooltip = "Name of a variable where the command's result will be stored")]
            public VariableStructure Result { get; set; } = new VariableStructure("result");

        }

        public ExcelGetRangeCommand(AbstractScripter scripter) : base(scripter)
        {
        }

        public void Execute(Arguments arguments)
        {
            if (arguments.ColStartIndex.Value < 1 || arguments.ColEndIndex.Value < 1 || arguments.RowStartIndex.Value < 1 || arguments.RowEndIndex.Value < 1)
                throw new Exception("Index cannot be smaller than 1");

            if (arguments.ColStartIndex.Value > arguments.ColEndIndex.Value || arguments.RowStartIndex.Value > arguments.RowEndIndex.Value)
                throw new Exception("Starting index cannot be bigger than ending index");

            var result = ExcelManager.CurrentExcel.GetRangeValue(arguments.ColStartIndex.Value, arguments.ColEndIndex.Value,
                             arguments.RowStartIndex.Value, arguments.RowEndIndex.Value);
            Scripter.Variables.SetVariableValue(arguments.Result.Value, new DataTableStructure(result, null, Scripter));
        }

    }
}
