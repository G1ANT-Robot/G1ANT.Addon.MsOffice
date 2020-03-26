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
    [Command(Name = "access.setactiverecord", Tooltip = "This command makes the specified record the current record in an open table, form or query result set")]
    public class AccessGotoRecordCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Tooltip = "Specifies the type of object that contains the record that you want to make current. Valid values are `ActiveDataObject`, `DataTable`, `DataQuery`, `DataForm`, `DataReport`, DataServerView`, `DataStoredProcedure`, `DataFunction`. The default value is `ActiveDataObject`")]
            public TextStructure ObjectType { get; set; } = new TextStructure("ActiveDataObject");

            [Argument(Tooltip = "Name of an object of the type selected by the `ObjectType` argument. This value has to be set if `ObjectType` parameter differs from ActiveDataObject")]
            public TextStructure ObjectName { get; set; }

            [Argument(Tooltip = "Specifies the record to make the current record. Valid values are `Previous`, `Next`, `First`, `Last`, `GoTo`, `NewRec`. The default value is Next.")]
            public TextStructure Operation { get; set; } = new TextStructure("Next");

            [Argument(Tooltip = "Number of records to move forward or backward if you specify Next or Previous for the `Operation` argument or the record to move to if you specify `GoTo` for the `Operation` argument. The value must be a valid record number. Default is 0.")]
            public IntegerStructure Offset { get; set; } = new IntegerStructure(0);
        }

        public AccessGotoRecordCommand(AbstractScripter scripter) : base(scripter)
        { }

        public void Execute(Arguments arguments)
        {
            AccessManager.CurrentAccess.GoToRecord(
                arguments.ObjectType.Value,
                arguments.ObjectName?.Value,
                arguments.Operation.Value,
                arguments.Offset.Value
            );
        }
    }
}