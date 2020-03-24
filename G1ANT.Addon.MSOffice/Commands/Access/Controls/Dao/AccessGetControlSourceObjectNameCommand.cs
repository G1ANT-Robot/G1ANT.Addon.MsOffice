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
    [Command(Name = "access.getcontrolsourceobjectname", Tooltip = "Get name of table that is data source for Access control (`SourceObject` property, equivalent of `control.getcontrolproperty nameofproperty SourceObject`)")]
    public class AccessGetControlSourceObjectNameCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Required = true, Tooltip = "Path to the control. See `access.getcontrol` tooltip for path examples")]
            public TextStructure Path { get; set; }

            [Argument(Tooltip = "Name of a variable where the command's result will be stored")]
            public VariableStructure Result { get; set; } = new VariableStructure("result");
        }

        public AccessGetControlSourceObjectNameCommand(AbstractScripter scripter) : base(scripter)
        { }

        public void Execute(Arguments arguments)
        {
            var control = AccessManager.CurrentAccess.GetControlByPath(arguments.Path.Value);
            var sourceObjectName = control.GetPropertyValue(AccessGetControlSourceObjectDetailsCommand.PropertyName);

            Scripter.Variables.SetVariableValue(arguments.Result.Value, new TextStructure(sourceObjectName));
        }
    }
}