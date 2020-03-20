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

namespace G1ANT.Addon.MSOffice.Commands.Access.Application
{
    [Command(Name = "access.close", Tooltip = "This command closes Access database instance. Consider usage of `access.quit` in order to close Access process")]
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