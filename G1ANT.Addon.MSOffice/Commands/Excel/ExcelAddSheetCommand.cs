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
    [Command(Name = "excel.addsheet", Tooltip = "This command adds a new sheet to the currently active Excel instance")]
    public class ExcelAddSheetCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Tooltip = "Name of a sheet to be added")]
            public TextStructure Name { get; set; } = new TextStructure(string.Empty);
        }
        public ExcelAddSheetCommand(AbstractScripter scripter) : base(scripter)
        {
        }

        public void Execute(Arguments arguments)
        {
            try
            {
                ExcelManager.CurrentExcel.AddSheet(arguments.Name.Value);
            }
            catch (Exception ex)
            {
                throw new ApplicationException($"Error occured while adding sheet to current excel instance. Name: '{arguments.Name.Value}'. Message: '{ex.Message}'", ex);
            }
        }
    }
}
