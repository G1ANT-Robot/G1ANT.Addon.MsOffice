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

namespace G1ANT.Addon.MSOffice.Access
{
    [Command(Name = "access.quit", Tooltip = "This command quits Access instance")]
    public class AccessQuitCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Tooltip = "Set to false to discard any changes. True by default")]
            public BooleanStructure SaveAllChanges { get; set; } = new BooleanStructure(true);
        }

        public AccessQuitCommand(AbstractScripter scripter) : base(scripter)
        { }

        public void Execute(Arguments arguments)
        {
            AccessManager.CurrentAccess.Quit(arguments.SaveAllChanges.Value);
        }
    }
}