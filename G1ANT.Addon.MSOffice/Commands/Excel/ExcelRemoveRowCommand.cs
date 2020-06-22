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
    [Command(Name = "excel.removerow", Tooltip = "This command deletes the specified row")]
    public class ExcelRemoveRowCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Required = true, Tooltip = "Number of a row to be deleted")]
            public IntegerStructure Row { get; set; }
        }
        public ExcelRemoveRowCommand(AbstractScripter scripter) : base(scripter)
        {
        }
        public void Execute(Arguments arguments)
        {
            try
            {
                ExcelManager.CurrentExcel.RemoveRow(arguments.Row.Value);
            }
            catch (Exception ex)
            {
                throw new ApplicationException($"Error occured while trying to remove row. Row: '{arguments.Row.Value}'. Message: {ex.Message}", ex);
            }            
        }
    }
}
