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

namespace G1ANT.Addon.MSOffice.Access
{
    [Command(Name = "access.close", Tooltip = "This command closes Access database leaving the Access instance open")]
    public class AccessCloseCommand : Command
    {
        public class Arguments : CommandArguments
        {
        }

        public AccessCloseCommand(AbstractScripter scripter) : base(scripter)
        { }

        public void Execute(Arguments arguments)
        {
            AccessManager.CurrentAccess.CloseDatabase();
        }
    }
}