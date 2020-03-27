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

namespace G1ANT.Addon.MSOffice.Commands.Access.Controls
{
    [Command(Name = "access.setfocus", Tooltip = "Set focus on control selected by path")]
    public class AccessSetFocusCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Required = true, Tooltip = "Path to the control. See `access.getcontrol` tooltip for path examples")]
            public TextStructure Path { get; set; }
        }

        public AccessSetFocusCommand(AbstractScripter scripter) : base(scripter)
        { }

        public void Execute(Arguments arguments)
        {
            var control = AccessManager.CurrentAccess.GetControlByPath(arguments.Path.Value);

            var handle = (IntPtr)control.GetFormHwnd();
            RobotWin32.BringWindowToFront(handle);
            RobotWin32.ShowWindow(handle, RobotWin32.ShowWindowEnum.Maximize);
            control.SetFocus();
        }
    }
}
