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

namespace G1ANT.Addon.MSOffice.Commands.Access
{
    [Command(Name = "access.open", Tooltip = "This command opens a new Access instance")]
    public class AccessOpenCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Tooltip = "Path of a file that has to be opened in Access", Required = true)]
            public TextStructure Path { get; set; }

            [Argument(Tooltip = "Password required to open Access database")]
            public TextStructure Password { get; set; } = new TextStructure("");

            [Argument(Tooltip = "Set to true to open excusively. False by default")]
            public BooleanStructure OpenExclusive { get; set; } = new BooleanStructure(false);

            [Argument(Tooltip = "Set to fale to hide the application. True by default")]
            public BooleanStructure Show { get; set; } = new BooleanStructure(true);

            [Argument(Tooltip = "Name of a variable where a currently opened Access process number is stored. It can be used in the `access.switch` command")]
            public VariableStructure Result { get; set; } = new VariableStructure("result");
        }

        public AccessOpenCommand(AbstractScripter scripter) : base(scripter)
        { }

        public void Execute(Arguments arguments)
        {
            try
            {
                var access = AccessManager.AddAccess();
                access.Open(arguments.Path.Value, arguments.Password.Value, arguments.OpenExclusive.Value, arguments.Show.Value);

                Scripter.Variables.SetVariableValue(arguments.Result.Value, new IntegerStructure(access.Id));
            }
            catch (Exception ex)
            {
                //if (ex.GetType() == typeof(COMException) && ex.Message.Contains("80040154"))
                //    throw new Exception("Could not find Microsoft Office on computer. Please make sure it is installed and try again.");
                throw new ApplicationException($"Problem occured while opening access instance. Path: '{arguments.Path?.Value}'", ex);
            }
        }
    }
}