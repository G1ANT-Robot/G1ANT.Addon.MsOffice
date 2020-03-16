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

namespace G1ANT.Addon.MSOffice.Commands.Access
{
    [Command(Name = "access.killorphanedprocesses", Tooltip = "This command quits all Access proceses that are not registered in Running Object Table")]
    public class AccessKillOrphanedProcessesCommand : Command
    {
        public class Arguments : CommandArguments
        {
        }

        public AccessKillOrphanedProcessesCommand(AbstractScripter scripter) : base(scripter)
        { }

        public void Execute(Arguments arguments)
        {
            AccessManager.KillOrphanedAccessProcesses();
        }
    }
}