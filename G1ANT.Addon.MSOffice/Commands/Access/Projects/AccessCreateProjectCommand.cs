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

namespace G1ANT.Addon.MSOffice.Commands.Access.Projects
{
    [Command(Name = "access.createproject", Tooltip = "This command creates a new Access project")]
    public class AccessCreateProjectCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Tooltip = "Name of the new Access project, including the path name and the file name extension", Required = true)]
            public TextStructure Path { get; set; }

            [Argument(Tooltip = "Optional valid connection string for the Access project")]
            public TextStructure ConnectionString { get; set; }

        }

        public AccessCreateProjectCommand(AbstractScripter scripter) : base(scripter)
        { }

        public void Execute(Arguments arguments)
        {
            AccessManager.CurrentAccess.CreateAccessProject(arguments.Path.Value, arguments.ConnectionString?.Value);
        }
    }
}