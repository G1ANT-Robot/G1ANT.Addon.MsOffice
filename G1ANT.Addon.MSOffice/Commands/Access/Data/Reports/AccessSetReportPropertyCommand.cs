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

namespace G1ANT.Addon.MSOffice.Commands.Access.Data.Reports
{
    [Command(Name = "access.setreportproperty", Tooltip = "This command sets a property of an Access Report")]
    public class AccessSetReportPropertyCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Tooltip = "Name of the report", Required = true)]
            public TextStructure ReportName { get; set; }

            [Argument(Tooltip = "Name of the property", Required = true)]
            public TextStructure PropertyName { get; set; }

            [Argument(Tooltip = "Value to set", Required = true)]
            public TextStructure Value { get; set; } = new TextStructure("Normal");
        }

        public AccessSetReportPropertyCommand(AbstractScripter scripter) : base(scripter)
        { }

        public void Execute(Arguments arguments)
        {
            var report = AccessManager.CurrentAccess.GetReportDetails(arguments.ReportName.Value);
            var value = ParseStringValueToProperType(arguments.Value.Value);

            report.SetProperty(arguments.PropertyName.Value, value);
        }

        private static object ParseStringValueToProperType(string value)
        {

            if (int.TryParse(value, out int intValue))
                return intValue;

            if (float.TryParse(value, out float floatValue))
                return floatValue;

            if (bool.TryParse(value, out bool boolValue))
                return boolValue;

            if (DateTime.TryParse(value, out DateTime dateTimeValue))
                return dateTimeValue;

            if (TimeSpan.TryParse(value, out TimeSpan timeSpanValue))
                return timeSpanValue;

            return value;
        }
    }
}