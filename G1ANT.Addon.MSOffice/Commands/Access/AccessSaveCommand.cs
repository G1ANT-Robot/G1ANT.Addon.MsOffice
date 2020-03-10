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
    [Command(Name = "access.save", Tooltip = "This command save changes in Access objects")]
    public class AccessSaveCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Required = true, Tooltip = "Type of object to save (object must be open, type must be one of values from `AcObjectType`: acDefault, acTable, acQuery, acForm, acReport, acMacro, acModule, acDataAccessPage, acServerView, acDiagram, acStoredProcedure, acFunction0, acDatabaseProperties, acTableDataMacro)")]
            public TextStructure ObjectType { get; set; }

            [Argument(Required = true, Tooltip = "Name for the object")]
            public TextStructure ObjectName { get; set; }
        }

        public AccessSaveCommand(AbstractScripter scripter) : base(scripter)
        { }

        public void Execute(Arguments arguments)
        {
            AccessManager.CurrentAccess.Save(arguments.ObjectType.Value, arguments.ObjectName.Value);
        }
    }
}