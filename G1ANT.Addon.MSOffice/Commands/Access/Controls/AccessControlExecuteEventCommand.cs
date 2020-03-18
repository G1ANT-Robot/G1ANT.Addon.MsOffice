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

namespace G1ANT.Addon.MSOffice.Commands.Access.Controls
{
    [Command(Name = "access.executecontrolevent", Tooltip = "Executes action assigned to a event at control selected by path. This command has several limitations")]
    public class AccessControlExecuteEventCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Required = true, Tooltip = "Path to the control. See `access.getcontrol` tooltip for path examples")]
            public TextStructure Path { get; set; }

            [Argument(Required = true, Tooltip = "Event name, i.e. OnClick, OnEnter")]
            public TextStructure EventName { get; set; }
        }

        public AccessControlExecuteEventCommand(AbstractScripter scripter) : base(scripter)
        { }

        public void Execute(Arguments arguments)
        {
            AccessManager.CurrentAccess.ExecuteEvents(arguments.Path.Value, arguments.EventName.Value);
        }
    }
}
