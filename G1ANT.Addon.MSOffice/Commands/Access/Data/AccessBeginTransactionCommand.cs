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

namespace G1ANT.Addon.MSOffice.Commands.Access.Data
{
    [Command(Name = "access.begintransaction", Tooltip = "This command starts new database transaction")]
    public class AccessBeginTransactionCommand : Command
    {
        public class Arguments : CommandArguments
        {
        }

        public AccessBeginTransactionCommand(AbstractScripter scripter) : base(scripter)
        { }

        public void Execute(Arguments arguments)
        {
            AccessManager.CurrentAccess.BeginTransaction();
        }
    }
}