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

//namespace G1ANT.Addon.MSOffice.Access
//{
//    [Command(Name = "access.getselectedindex", Tooltip = "Get index of selected item of an Access control selected by path")]
//    public class AccessGetSelectedIndexCommand : Command
//    {
//        public class Arguments : CommandArguments
//        {
//            [Argument(Required = true, Tooltip = "Path to the control. See `access.findcontrol` tooltip for path examples")]
//            public TextStructure Path { get; set; }
//        }

//        public AccessGetSelectedIndexCommand(AbstractScripter scripter) : base(scripter)
//        { }

//        public void Execute(Arguments arguments)
//        {
//            var control = AccessManager.CurrentAccess.GetAccessControlByPath(arguments.Path.Value);
//            control.SetFocus();
//        }
//    }
//}
