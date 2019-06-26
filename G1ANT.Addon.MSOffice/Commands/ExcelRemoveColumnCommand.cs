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
    [Command(Name = "excel.removecolumn", Tooltip = "This command removes the specified column")]
    public class ExcelRemoveColumnCommand : Command
    {
        public class Arguments : CommandArguments
        {

            [Argument(Tooltip = "Cell's column index")]
            public IntegerStructure ColIndex { get; set; }

            [Argument(Tooltip = "Cell's column name")]
            public TextStructure ColName { get; set; }
        }
        public ExcelRemoveColumnCommand(AbstractScripter scripter) : base(scripter)
        {
        }
        public void Execute(Arguments arguments)
        {
            object col = null;
            try
            {
                if (arguments.ColIndex != null)
                    col = arguments.ColIndex.Value;
                else if (arguments.ColName != null)
                    col = arguments.ColName.Value;
                else
                    throw new ArgumentException("One of the ColIndex or ColName arguments have to be set up.");
                ExcelManager.CurrentExcel.RemoveColumn(col);
            }
            catch (Exception ex)
            {
                throw new ApplicationException($"Error occured while trying to remove column. Column: '{col}'. Message: {ex.Message}", ex);
            }            
        }
    }
}
