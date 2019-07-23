using System;
using G1ANT.Addon.MSOffice.Helpers;
using G1ANT.Language;

namespace G1ANT.Addon.MSOffice.Commands
{
    [Command(Name = "excel.getcolor", Tooltip = "This command gets color of cell in current excel worksheet.")]
    public class ExcelGetColorCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Required = true, Tooltip = "Cell's row number")]
            public IntegerStructure Row { get; set; }

            [Argument(Tooltip = "Cell's column index")]
            public IntegerStructure ColIndex { get; set; }

            [Argument(Tooltip = "Cell's column name")]
            public TextStructure ColName { get; set; }

            [Argument(Tooltip = "Name of a variable where the font color's result will be stored")]
            public VariableStructure FontColorResult { get; set; } = new VariableStructure("fontcolorresult");

            [Argument(Tooltip = "Name of a variable where the background color's result will be stored")]
            public VariableStructure BackgroundResult { get; set; } = new VariableStructure("backgroundcolorresult");

        }

        public ExcelGetColorCommand(AbstractScripter scripter) : base(scripter) { }

        public void Execute(Arguments arguments)
        {
            object col = null;
            try
            {
                col = ExcelHelper.GetColumn(arguments.ColIndex, arguments.ColName, true);

                var colors = ExcelManager.CurrentExcel.GetColor(arguments.Row.Value, col);

                Scripter.Variables.SetVariableValue(arguments.BackgroundResult.Value, new ColorStructure(colors.Item1));
                Scripter.Variables.SetVariableValue(arguments.FontColorResult.Value,  new ColorStructure(colors.Item2));
            }
            catch (Exception ex)
            {
                throw new ApplicationException($"Problem occured getting color of the cell. Col: '{col}', Row: '{arguments.Row.Value}'. Message: {ex.Message}", ex);
            }
        }
    }
}
