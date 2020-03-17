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

namespace G1ANT.Addon.MSOffice.Commands.Access.Data
{
    [Command(Name = "access.opendatabasediagram", Tooltip = "This command opens an Access Database Diagram")]
    public class AccessOpenDatabaseDiagramCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Tooltip = "Name of the diagram to open", Required = true)]
            public TextStructure Name { get; set; }
        }

        public AccessOpenDatabaseDiagramCommand(AbstractScripter scripter) : base(scripter)
        { }

        public void Execute(Arguments arguments)
        {
            AccessManager.CurrentAccess.OpenDiagram(arguments.Name.Value);
        }
    }
}