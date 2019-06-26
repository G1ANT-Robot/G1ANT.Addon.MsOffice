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


using System.Linq;
using System;

namespace G1ANT.Addon.MSOffice
{
    [Command(Name = "excel.switch", Tooltip = "This command switches to another Excel instance opened by G1ANT.Robot")]
    public class ExcelSwitchCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Required = true, Tooltip = "ID number of an Excel instance that will be activated")]
            public IntegerStructure Id { get; set; }

        }

        public ExcelSwitchCommand(AbstractScripter scripter) : base(scripter)
        {
        }
        public void Execute(Arguments arguments)
        {
            try
            {
                ExcelManager.SwitchExcel(arguments.Id.Value);
            }
            catch (Exception ex)
            {
                throw new ApplicationException($"Problem occured while switching to another excel instance. Id: '{arguments.Id.Value}'. Message: '{ex.Message}'");
            }
        }
    }
}