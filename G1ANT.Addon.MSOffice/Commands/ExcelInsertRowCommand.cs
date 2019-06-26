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



namespace G1ANT.Addon.MSOffice
{
    [Command(Name = "excel.insertrow", Tooltip = "This command inserts an empty row into a specified place")]
    public class ExcelInsertRowCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Required = true, Tooltip = "Row number")]
            public IntegerStructure Row { get; set; }

            [Argument(Tooltip = "Specifies where to insert a row: `above` or `below` a specified row")]
            public TextStructure Where { get; set; } = new TextStructure("below");
        }
        public ExcelInsertRowCommand(AbstractScripter scripter) : base(scripter)
        {
        }
        public void Execute(Arguments arguments)
        {
            try
            {
                ExcelManager.CurrentExcel.InsertRow(arguments.Row.Value, arguments.Where.Value);
            }
            catch (Exception ex)
            {
                throw new ApplicationException($"Error occured while trying to insert row. Row: '{arguments.Row.Value}', where: '{arguments.Where.Value}'. Message: {ex.Message}", ex);
            }            
        }
    }
}
