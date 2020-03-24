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
    [Command(Name = "access.setnewpassword", Tooltip = "This command sets new password for Access database")]
    public class AccessSetNewPasswordCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Tooltip = "Old password", Required = true)]
            public TextStructure OldPassword { get; set; }

            [Argument(Tooltip = "New password", Required = true)]
            public TextStructure NewPassword { get; set; }
        }

        public AccessSetNewPasswordCommand(AbstractScripter scripter) : base(scripter)
        { }

        public void Execute(Arguments arguments)
        {
            AccessManager.CurrentAccess.SetNewPassword(arguments.OldPassword.Value, arguments.NewPassword.Value);
        }
    }
}