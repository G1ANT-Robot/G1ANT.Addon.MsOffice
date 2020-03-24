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
    [Command(Name = "excel.duplicaterow", Tooltip = "This command copies a specified row to a specified place")]
    public class ExcelDuplicateRowCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Required = true, Tooltip = "Source row number")]
            public IntegerStructure Source { get; set; }

            [Argument(Required = true, Tooltip = "Destination row number")]
            public IntegerStructure Destination { get; set; }
        }
        public ExcelDuplicateRowCommand(AbstractScripter scripter) : base(scripter)
        {
        }
        public void Execute(Arguments arguments)
        {
            try
            {
                ExcelManager.CurrentExcel.DuplicateRow(arguments.Source.Value, arguments.Destination.Value);
            }
            catch (Exception ex)
            {
                throw new ApplicationException($"Error occured while trying to duplicate row. Source: '{arguments.Source.Value}', destination: '{arguments.Destination.Value}'. Message: {ex.Message}", ex);
            }            
        }
    }
}
