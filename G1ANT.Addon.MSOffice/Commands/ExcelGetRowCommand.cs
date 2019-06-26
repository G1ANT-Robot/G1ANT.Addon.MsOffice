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

namespace G1ANT.Addon.MSOffice
{
    [Command(Name = "excel.getrow", Tooltip = "This command gets all used cells of the specified row")]
    public class ExcelGetRowCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Required = true, Tooltip = "Row number")]
            public IntegerStructure Row { get; set; }

            [Argument(Tooltip = "Name of a variable where the command's result will be stored")]
            public VariableStructure Result { get; set; } = new VariableStructure("result");
        }
        public ExcelGetRowCommand(AbstractScripter scripter) : base(scripter)
        {
        }
        public void Execute(Arguments arguments)
        {
            try
            {
                var val = ExcelManager.CurrentExcel.GetRow(arguments.Row.Value);
                var structureDictionary = val.ToDictionary(x => x.Key.ToLower(), x => x.Value);
                Scripter.Variables.SetVariableValue(arguments.Result.Value, new DictionaryStructure(structureDictionary,"",Scripter)); 
            }
            catch (Exception ex)
            {
                throw new ApplicationException($"Problem occured while getting row: '{arguments.Row.Value}'. Message: {ex.Message}", ex);
            }
        }
    }
}
