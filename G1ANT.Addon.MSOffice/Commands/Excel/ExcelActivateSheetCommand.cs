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
using System.Windows.Forms;

namespace G1ANT.Addon.MSOffice
{
    [Command(Name = "excel.activatesheet", Tooltip = "This command activates a specified sheet in the currently active Excel instance")]
    public class ExcelActivateSheetCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Required = true, Tooltip = "Name of a sheet to be activated")]
            public TextStructure Name { get; set; }
        }
        public ExcelActivateSheetCommand(AbstractScripter scripter) : base(scripter)
        {
        }

        public void Execute(Arguments arguments)
        {
            try
            {
                ExcelManager.CurrentExcel.ActivateSheet(arguments.Name.Value);
            }
            catch (Exception ex)
            {
                throw new ApplicationException($"Error occured while activating sheet. Sheet name: '{arguments.Name.Value}' Message: '{ex.Message}'", ex);
            }
        }
    }
}
