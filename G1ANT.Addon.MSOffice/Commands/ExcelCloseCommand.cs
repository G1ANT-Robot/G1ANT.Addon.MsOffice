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
    [Command(Name = "excel.close", Tooltip = "This command closes the currently active Excel instance")]
    public class ExcelCloseCommand : Command
    {
        public class Arguments : CommandArguments
        {
        }

        public ExcelCloseCommand(AbstractScripter scripter) : base(scripter)
        {
        }

        public void Execute(Arguments arguments)
        {
            try
            {
                ExcelManager.RemoveInstance();
            }
            catch (Exception ex)
            {
                throw new ApplicationException($"Error occured while closing current excel instance. Message: '{ex.Message}'", ex);
            }
        }
    }
}
