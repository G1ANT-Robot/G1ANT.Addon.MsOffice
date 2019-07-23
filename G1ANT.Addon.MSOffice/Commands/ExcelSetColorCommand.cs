using System;
using G1ANT.Addon.MSOffice.Helpers;
using G1ANT.Language;

namespace G1ANT.Addon.MSOffice.Commands
{
    [Command(Name = "excel.setcolor", Tooltip = "This command sets color of cell in current excel worksheet.")]
    public class ExcelSetColorCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Required = true, Tooltip = "Cell's row number")]
            public IntegerStructure Row { get; set; }

            [Argument(Tooltip = "Cell's column index")]
            public IntegerStructure ColIndex { get; set; }

            [Argument(Tooltip = "Cell's column name")]
            public TextStructure ColName { get; set; }

            [Argument(Tooltip = "Font color to be set")]
            public ColorStructure FontColor { get; set; }

            [Argument(Tooltip = "Background color to be set")]
            public ColorStructure BackgroundColor { get; set; }

        }

        public ExcelSetColorCommand(AbstractScripter scripter) : base(scripter) { }

        public void Execute(Arguments arguments)
        {
            object col = null;
            try
            {
                col = ExcelHelper.GetColumn(arguments.ColIndex, arguments.ColName, true);
                ExcelManager.CurrentExcel.SetColor(arguments.Row.Value, col, arguments.BackgroundColor?.Value, arguments.FontColor.Value);
            }
            catch (Exception ex)
            {
                throw new ApplicationException($"Problem occured setting color of the cell. Col: '{col}', Row: '{arguments.Row.Value}'. Message: {ex.Message}", ex);
            }
        }
    }
}
