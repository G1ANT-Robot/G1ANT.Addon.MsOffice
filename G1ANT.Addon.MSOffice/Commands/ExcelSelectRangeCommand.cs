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
using Org.BouncyCastle.Crypto.Agreement.JPake;

namespace G1ANT.Addon.MSOffice
{
    [Command(Name = "excel.selectrange", Tooltip = "This command selects a range in the currently active Excel instance")]
    public class ExcelSelectRangeCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Tooltip = "Starting Cell's column index")]
            public IntegerStructure ColIndex1 { get; set; }

            [Argument(Tooltip = "Starting Cell's column name")]
            public TextStructure ColName1 { get; set; } = new TextStructure(string.Empty);

            [Argument(Tooltip = "Starting Cell's row index", Required = true)]
            public IntegerStructure Row1 { get; set; }

            [Argument(Tooltip = "Ending Cell's column index")]
            public IntegerStructure ColIndex2 { get; set; }

            [Argument(Tooltip = "Ending Cell's column name")]
            public TextStructure ColName2 { get; set; } = new TextStructure(string.Empty);

            [Argument(Tooltip = "Ending Cell's row index", Required = true)]
            public IntegerStructure Row2 { get; set; }
        }

        public ExcelSelectRangeCommand(AbstractScripter scripter) : base(scripter)
        {
        }
        

        public void Execute(Arguments arguments)
        {
            try
            {
                object col1, col2;
                if (arguments.ColIndex1 != null)
                    col1 = arguments.ColIndex1.Value;
                else if (arguments.ColName1 != null)
                    col1 = arguments.ColName1.Value;
                else
                    throw new ArgumentException("One of the ColIndex1 or ColName1 arguments have to be set up.");

                if (arguments.ColIndex2 != null)
                    col2 = arguments.ColIndex2.Value;
                else if (arguments.ColName2 != null)
                    col2 = arguments.ColName2.Value;
                else
                    throw new ArgumentException("One of the ColIndex2 or ColName2 arguments have to be set up.");


                ExcelManager.CurrentExcel.SelectRange(col1, arguments.Row1.Value, col2, arguments.Row2.Value);
            }
            catch (Exception ex)
            {
                throw new ApplicationException($"Error occured while selecting range in current excel instance. Message: '{ex.Message}'", ex);
            }
        }
    }
}
