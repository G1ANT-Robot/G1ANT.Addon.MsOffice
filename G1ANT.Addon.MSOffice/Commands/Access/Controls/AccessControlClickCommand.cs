///**
//*    Copyright(C) G1ANT Ltd, All rights reserved
//*    Solution G1ANT.Addon, Project G1ANT.Addon.MSOffice
//*    www.g1ant.com
//*
//*    Licensed under the G1ANT license.
//*    See License.txt file in the project root for full license information.
//*
//*/
//using G1ANT.Language;

//namespace G1ANT.Addon.MSOffice.Commands.Access.Controls
//{
//    [Command(Name = "access.controls.click", Tooltip = "Perform a mouse click at control selected by path")]
//    public class AccessControlClickCommand : Command
//    {
//        public class Arguments : CommandArguments
//        {
//            [Argument(Required = true, Tooltip = "Path to the control. See `access.control.find` tooltip for path examples")]
//            public TextStructure Path { get; set; }
//        }

//        public AccessControlClickCommand(AbstractScripter scripter) : base(scripter)
//        { }

//        public void Execute(Arguments arguments)
//        {
//            var control = AccessManager.CurrentAccess.GetAccessControlByPath(arguments.Path.Value);
//            AccessManager.CurrentAccess.Click(control);
//        }
//    }
//}
