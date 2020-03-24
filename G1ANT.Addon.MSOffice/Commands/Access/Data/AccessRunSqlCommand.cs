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

namespace G1ANT.Addon.MSOffice.Commands.Access.Data
{
    [Command(Name = "access.runsql", Tooltip = "This command runs a non-selective (insert, update, alter, create, ...) SQL query in open Access database. If you need to select some data use command `access.executesql`")]
    public class AccessRunSqlCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Tooltip = "SQL query to run in open Access database", Required = true)]
            public TextStructure SQL{ get; set; }

            [Argument(Tooltip = "Run query within transaction. Default is false")]
            public BooleanStructure UseTransaction { get; set; } = new BooleanStructure(false);
        }

        public AccessRunSqlCommand(AbstractScripter scripter) : base(scripter)
        { }

        public void Execute(Arguments arguments)
        {
            AccessManager.CurrentAccess.RunSql(arguments.SQL.Value, arguments.UseTransaction.Value);
        }
    }
}