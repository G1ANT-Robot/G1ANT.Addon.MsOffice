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

namespace G1ANT.Addon.MSOffice.Commands.Access.Controls.Items
{
    [Command(Name = "access.selectcontrolitem", Tooltip = "Mark item of Access control selected by path as selected")]
    public class AccessSelectItemCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Required = true, Tooltip = "Path to the control. See `access.getcontrol` tooltip for path examples")]
            public TextStructure Path { get; set; }

            [Argument(Required = true, Tooltip = "Index of item (starting from 0)")]
            public IntegerStructure Index { get; set; }
       }

        public AccessSelectItemCommand(AbstractScripter scripter) : base(scripter)
        { }

        public void Execute(Arguments arguments)
        {
            var control = AccessManager.CurrentAccess.GetControlByPath(arguments.Path.Value);
            control.SetItemSelected(arguments.Index.Value, true);
        }
    }
}
