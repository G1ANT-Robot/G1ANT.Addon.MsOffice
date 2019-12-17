/**
*    Copyright(C) G1ANT Ltd, All rights reserved
*    Solution G1ANT.Addon, Project G1ANT.Addon.MSOffice
*    www.g1ant.com
*
*    Licensed under the G1ANT license.
*    See License.txt file in the project root for full license information.
*
*/

using G1ANT.Addon.MSOffice.Helpers;
using G1ANT.Language;
using System;

namespace G1ANT.Addon.MSOffice
{
    [Command(Name = "excel.insertformula", Tooltip = "This command inserts formula to a specified cell")]
    public class ExcelInsertFormulaCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Required = true, Tooltip = "Formula to be inserted")]
            public TextStructure Formula { get; set; }

            [Argument(Required = true, Tooltip = "Cell's row number")]
            public IntegerStructure Row { get; set; }

            [Argument(Tooltip = "Cell's column index")]
            public IntegerStructure ColIndex { get; set; }

            [Argument(Tooltip = "Cell's column name")]
            public TextStructure ColName { get; set; }
        }

        public ExcelInsertFormulaCommand(AbstractScripter scripter) : base(scripter)
        {
        }

        public void Execute(Arguments arguments)
        {
            var column = ExcelHelper.GetColumn(arguments.ColIndex, arguments.ColName);

            try
            {
                ExcelManager.CurrentExcel.SetFormula(arguments.Row.Value, column, arguments.Formula.Value);
            }
            catch (Exception ex)
            {
                throw new ApplicationException($"Problem occured while inserting formula. Col: '{column}', Row: '{arguments.Row.Value}'. Message: {ex.Message}", ex);
            }
        }
    }
}
