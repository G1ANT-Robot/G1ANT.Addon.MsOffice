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
    [Command(Name = "excel.save", Tooltip = "This command saves the currently active Excel workbook")]
    public class ExcelSaveCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Tooltip = "Full path to a file to be saved. If not specified, G1ANT.Robot will try to save the file using the path it was loaded from. If the current Excel instance was opened with no path specified, error handling will be applied")]
            public TextStructure Path { get; set; }


        }
        public ExcelSaveCommand(AbstractScripter scripter) : base(scripter)
        {
        }
        public void Execute(Arguments arguments)
        {
            try
            {
                ExcelManager.CurrentExcel.Save(arguments.Path?.Value);

            }
            catch (Exception ex)
            {
                throw new ApplicationException($"Unable to save file: '{arguments.Path.Value}'", ex);
            }
        }
    }
}
