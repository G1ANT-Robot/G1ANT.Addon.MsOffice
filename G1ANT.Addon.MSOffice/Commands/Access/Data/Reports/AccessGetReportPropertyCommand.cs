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
using System.Linq;

namespace G1ANT.Addon.MSOffice.Commands.Access.Data.Reports
{
    [Command(Name = "access.getreportproperty", Tooltip = "This command get a value of property of Access Report")]
    public class AccessGetReportPropertyCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Tooltip = "Name of the report", Required = true)]
            public TextStructure ReportName { get; set; }

            [Argument(Tooltip = "Name of the property", Required = true)]
            public TextStructure PropertyName { get; set; }

            [Argument(Tooltip = "Name of a variable where the command's result will be stored")]
            public VariableStructure Result { get; set; } = new VariableStructure("result");
        }

        public AccessGetReportPropertyCommand(AbstractScripter scripter) : base(scripter)
        { }

        public void Execute(Arguments arguments)
        {
            var report = AccessManager.CurrentAccess.GetReportDetails(arguments.ReportName.Value);
            var property = report.Properties.Value.FirstOrDefault(rp => rp.Name == arguments.PropertyName.Value);

            if (property == null)
                throw new ApplicationException($"Property {arguments.PropertyName.Value} not found in report {arguments.ReportName.Value}");

            Scripter.Variables.SetVariableValue(arguments.Result.Value, new TextStructure(property.Value));
        }
    }
}