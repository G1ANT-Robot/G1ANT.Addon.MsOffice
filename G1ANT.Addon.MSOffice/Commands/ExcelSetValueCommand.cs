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
using System.Linq;
using G1ANT.Language;

namespace G1ANT.Addon.MSOffice
{


    [Command(Name = "excel.setvalue", Tooltip = "This command enters a value into a specified cell")]
    public class ExcelSetValueCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Required = true, Tooltip = "Any value to be set in a specified cell")]
            public TextStructure Value { get; set; }

            [Argument(Required = true, Tooltip = "Cell's row number")]
            public IntegerStructure Row { get; set; }

            [Argument(Tooltip = "Cell's column index")]
            public IntegerStructure ColIndex { get; set; }

            [Argument(Tooltip = "Cell's column name")]
            public TextStructure ColName { get; set; }
        }


        public ExcelSetValueCommand(AbstractScripter scripter) : base(scripter)
        { }

        public void Execute(Arguments arguments)
        {
            object col = null;

            try
            {
                if (HasCorrectColumnArguments(arguments))
                {
                    if (arguments.ColName != null)
                    {
                        if (arguments.ColName.Value.All(x => x.IsLetter()))
                        {
                            col = arguments.ColName.Value;
                        }
                        else
                        {
                            throw new ArgumentException("ColName should not contain any special characters or digits.");
                        }
                    }
                    else
                    {
                        col = arguments.ColIndex.Value;
                    }

                    ExcelManager.CurrentExcel.SetCellValue(arguments.Row.Value, col, arguments.Value.Value);
                }
            }
            catch (Exception ex)
            {
                throw new ApplicationException(
                    $"Problem occured while setting value. Row: '{arguments.Row.Value}', Col: '{col}', Val: '{arguments.Value.Value}'",
                    ex);
            }
        }

        private static bool HasCorrectColumnArguments(Arguments arguments)
        {
            if (arguments.ColIndex == null && arguments.ColName == null)
            {
                throw new ArgumentException("One of the ColIndex or ColName arguments have to be set up.");
            }

            if (arguments.ColName != null && arguments.ColIndex != null)
            {
                throw new ArgumentException("Only of one the ColIndex or ColName arguments should be set up.");
            }

            return true;
        }
    }
}
