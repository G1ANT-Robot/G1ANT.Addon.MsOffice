﻿/**
*    Copyright(C) G1ANT Ltd, All rights reserved
*    Solution G1ANT.Addon, Project G1ANT.Addon.MSOffice
*    www.g1ant.com
*
*    Licensed under the G1ANT license.
*    See License.txt file in the project root for full license information.
*
*/
using Microsoft.Office.Interop.Excel;
using Microsoft.Vbe.Interop;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;

using G1ANT.Language;
using System.Globalization;
using System.Threading;

namespace G1ANT.Addon.MSOffice
{
    public class ExcelWrapper
    {
        private Microsoft.Office.Interop.Excel.Application application = null;
        private Workbook workbook = null;
        private _Worksheet sheet = null;
        private string path = null;

        public int Id { get; set; }

        public ExcelWrapper(int id)
        {
            this.Id = id;
        }

        public void SetCellValue(int rowNumber, object columnNumber, string value)
        {
            sheet.Cells[rowNumber, columnNumber].Select();
            sheet.Cells[rowNumber, columnNumber] = value;
        }

        public string GetCellValue(int rowNumber, object columnAddress)
        {
            try
            {
                var range = sheet.Cells[rowNumber, columnAddress] as Range;
                return range.Text?.ToString() ?? string.Empty;
            }
            catch
            {
                throw new ArgumentException($"Wrong row number: '{rowNumber}' or column address: '{columnAddress}'.");
            }
        }

        public System.Data.DataTable GetRangeValue(int startingColIndex, int endingColIndex, int startingRowIndex, int endingRowIndex, bool hasHeaders)
        {
            var rangeTable = new System.Data.DataTable("RangeTable");

            var nbOfColumns = endingColIndex - startingColIndex + 1;
            var nbOfRows = endingRowIndex - startingRowIndex + 1;

            var startingCell = sheet.Cells[startingRowIndex, startingColIndex];
            var endingCell = sheet.Cells[endingRowIndex, endingColIndex];
            var range = sheet.Range[startingCell, endingCell].Cells.Value;

            for (int i = 0; i < nbOfColumns; i++)
                rangeTable.Columns.Add("Column" + i.ToString());

            if (hasHeaders)
            {
                for (int i = 1; i < nbOfColumns + 1; i++)
                {
                    // nie wiem co w przypadku gdy ktos zaznacza has headers a wartosci sa puste
                    // nazwy kolumn nie moga byc nullowe
                    // moze rzucac jakis exception?
                    if (range[1, i] != null)
                        rangeTable.Columns[i - 1].ColumnName = Convert.ToString(range[1, i]);
                }
                nbOfRows += -1;
            }

            for (int i = 1; i < nbOfRows + 1; i++)
            {
                var dataRow = rangeTable.NewRow();

                for (int j = 1; j < nbOfColumns + 1; j++)
                    dataRow[j - 1] = Convert.ToString(range[i, j]);

                rangeTable.Rows.Add(dataRow);
            }

            return rangeTable;
        }

        public Dictionary<string, object> GetRow(int rowNumber)
        {
            if (rowNumber < 1)
            {
                throw new ArgumentException("Row number must be a positive integer");
            }
            Dictionary<string, object> ret = new Dictionary<string, object>();
            Range usedRange = sheet.UsedRange.Rows[rowNumber] as Range;
            if (usedRange != null)
            {
                for (int i = 1; i <= usedRange.Columns.Count; i++)
                {
                    var currentColumn = usedRange.Columns[i];
                    if (currentColumn != null)
                    {
                        Regex reg = new Regex(@"[A-Za-z]+");
                        if (reg.IsMatch(currentColumn.Address))
                        {
                            Match match = reg.Match(currentColumn.Address);
                            string columnName = match.Value;
                            string colValue = currentColumn.Text.ToString();
                            ret.Add(columnName, colValue);
                        }
                    }
                }
            }
            return ret;
        }

        public Tuple<Color, Color> GetColor(int row, object column)
        {
            try
            {
                var cell = sheet.Cells[row, column];
                var fontColor = ColorTranslator.FromOle((int)cell.Font.Color);
                var backgroundColor = ColorTranslator.FromOle((int)cell.Interior.Color);
                return new Tuple<Color, Color>(backgroundColor, fontColor);
            }
            catch
            {
                throw new ArgumentException("Wrong cells position arguments. Row must be a positive integer and column must be either positive integer or alphanumeric address.");
            }
        }

        public void SetColor(int row, object column, Color? backgroundColor, Color? fontColor)
        {
            try
            {
                var cell = sheet.Cells[row, column];
                if (backgroundColor.HasValue)
                    cell.Interior.Color = ColorTranslator.ToOle(backgroundColor.Value);
                if (fontColor.HasValue)
                    cell.Font.Color = ColorTranslator.ToOle(fontColor.Value);
            }
            catch
            {
                throw new ArgumentException("Wrong cells position arguments. Row must be a positive integer and column must be either positive integer or alphanumeric address.");
            }
        }

        public string GetFormula(int rowNumber, object columnNumber)
        {
            Range range = null;
            try
            {
                range = sheet.Cells[rowNumber, columnNumber];
            }
            catch
            {
                throw new ArgumentException("Wrong cells position arguments. Row must be a positive integer and column must be either positive integer or alphanumeric address.");
            }
            return range.Formula.ToString();
        }

        public void SetFormula(int rowNumber, object columnNumber, string formula)
        {
            try
            {
                sheet.Cells[rowNumber, columnNumber].Formula = formula;
            }
            catch
            {
                throw new ArgumentException("Wrong cells position arguments. Row must be a positive integer and column must be either positive integer or alphanumeric address.");
            }
        }

        public object RunMacro(string macroName, List<object> args)
        {
            List<object> arguments = new List<object> { macroName };
            object result = null;
            arguments.AddRange(args);
            result = application.GetType().InvokeMember("Run", BindingFlags.InvokeMethod, null, this.application, arguments.ToArray());
            return result;
        }

        public object RunMacroCode(string macroCode, string resultVariableName)
        {
            object result = null;
            VBComponent component = null;
            try
            {
                component = workbook.VBProject.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule);
                string macroName = $"G1ANT{Guid.NewGuid().ToString("N")}";
                if (!string.IsNullOrEmpty(resultVariableName))
                    macroCode = macroCode.Replace(resultVariableName, macroName);
                component.CodeModule.AddFromString($"Function {macroName}()\r\n{macroCode}\r\nEnd Function\r\n");
                result = RunMacro(macroName, new List<object>());
            }
            catch
            {
                throw;
            }
            finally
            {
                if (component != null)
                    workbook.VBProject.VBComponents.Remove(component);
            }

            return result;
        }

        public void CopySelectedCellsToClipboard()
        {
            if (!sheet.Application.Selection.Copy(Type.Missing))
                throw new ApplicationException("Sheet.Application.Selection.Copy returned false");
        }

        public void ActivateSheet(string name)
        {
            sheet = GetSheetByName(name);
            if (sheet == null && !string.IsNullOrEmpty(name))
            {
                throw new ArgumentException($"Could not find sheet name with specified name '{name}'.");
            }
            sheet = sheet ?? workbook.ActiveSheet;
            sheet.Activate();
        }

        public void AddSheet(string name)
        {
            _Worksheet currentActiveSheet = workbook.ActiveSheet;
            if (currentActiveSheet?.Name?.ToLower() == name?.ToLower())
            {
                throw new ArgumentException($"Can not add sheet because a sheet with the same name: '{name}' already exists");
            }
            else
            {
                currentActiveSheet = workbook.Sheets.Add();
                currentActiveSheet.Name = name;
                sheet = currentActiveSheet;
            }
        }

        private void InitialiseNewinstance(bool visibile)
        {
            application = new Microsoft.Office.Interop.Excel.Application();
            application.DisplayAlerts = false;
            application.Visible = visibile;
        }

        public void Open(string path, string sheetName, bool visibile = true)
        {
            InitialiseNewinstance(visibile);
            workbook = OpenWorkbook(path);
            ActivateSheet(sheetName);
            workbook.BeforeClose += new WorkbookEvents_BeforeCloseEventHandler(WorkbookBeforeClose);
        }

        private Workbook OpenWorkbook(string path)
        {
            Workbook workbook = null;
            try
            {
                workbook = string.IsNullOrEmpty(path) ? application.Workbooks.Add(Missing.Value) : application.Workbooks.Open(path);
            }
            catch (Exception ex)
            {
                throw new ArgumentException($"Unable to open excel file. Specified path: '{path}'. Message: {ex.Message}", ex);
            }
            return workbook;
        }

        public void Save(string filePath = null)
        {
            if (filePath == path || string.IsNullOrEmpty(filePath))
                workbook.Save();
            else
            {
                string savingPath = string.IsNullOrEmpty(filePath) ? filePath : path;
                path = savingPath;

                if (File.Exists(filePath))
                {
                    File.Delete(filePath);
                }

                workbook.SaveAs(filePath, XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, false, false,
                                 XlSaveAsAccessMode.xlNoChange, XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            }
        }

        public void Export(string filePath)
        {
            try
            {
                string type = Path.GetExtension(filePath)?.Replace(".", "")?.ToLower();
                XlFixedFormatType trueType;
                switch (type.ToString().ToLower())
                {
                    case "pdf":
                        trueType = XlFixedFormatType.xlTypePDF;
                        break;
                    case "xps":
                        trueType = XlFixedFormatType.xlTypeXPS;
                        break;
                    default:
                        throw new ArgumentException("Unsupported format type");
                }
                workbook.ExportAsFixedFormat(trueType, filePath);
            }
            catch (Exception ex)
            {
                throw new ArgumentException($"Problem occured while exporting currently active workbook. Path: '{filePath}', Message: {ex.Message}", ex);
            }
        }

        public void Paste()
        {
            sheet.GetType().InvokeMember("Paste", BindingFlags.InvokeMethod, null, sheet, null, Thread.CurrentThread.CurrentCulture);
        }

        public void Close()
        {
            object misValue = System.Reflection.Missing.Value;
            workbook.Close(false, misValue, misValue);
        }

        public void SelectRange(object startColumn, int startRow, object endColumn, int endRow)
        {
            if (startColumn == null || endColumn == null)
            {
                throw new ApplicationException("Starting cell's column and ending cell's column need to be specified.");
            }

            var startingCell = sheet.Cells[startRow, startColumn];
            var endingCell = sheet.Cells[endRow, endColumn];
            var range = sheet.Range[startingCell, endingCell];
            range.Select();
        }

        public void InsertColumn(object column, string where)
        {
            if (string.IsNullOrEmpty(where))
            {
                throw new ArgumentException("Argument 'where' can not be empty.");
            }
            if (column == null || string.IsNullOrEmpty(column?.ToString()))
            {
                throw new ArgumentException("Argument 'column' can not be empty.");
            }
            where = where.ToLower();
            if (where != "before" && where != "after")
            {
                throw new ArgumentException("Wrong 'where' argument. It accepts either 'before' or 'after' values.");
            }
            Range range = null;
            Range columnRange = null;
            try
            {
                range = sheet.Columns[column];
                columnRange = (where == "before") ? range.EntireColumn : range.EntireColumn.Next.EntireColumn;
            }
            catch (Exception)
            {
                throw new ArgumentException("Wrong 'column' argument.");
            }
            columnRange.Insert(where == "before" ? XlInsertFormatOrigin.xlFormatFromRightOrBelow : XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
        }

        public void RemoveColumn(object column)
        {
            if (string.IsNullOrEmpty(column?.ToString()) || column == null)
            {
                throw new ArgumentException("Argument 'column' can not be empty.");
            }
            Range range = null;
            Range columnRange = null;
            try
            {
                range = sheet.Columns[column];
                columnRange = range.EntireColumn;
            }
            catch (Exception)
            {
                throw new ArgumentException("Wrong 'column' argument.");
            }
            columnRange.Delete();
        }

        public void InsertRow(int row, string where)
        {
            if (string.IsNullOrEmpty(where))
            {
                throw new ArgumentException("Argument 'where' can not be empty.");
            }
            where = where.ToLower();
            if (where != "below" && where != "above")
            {
                throw new ArgumentException("Wrong 'where' argument. It accepts either 'below' or 'above' value.");
            }
            if (row < 1)
            {
                throw new ArgumentException("Row's number not correct.");
            }
            row = (where == "below") ? row + 1 : row;
            Range line = (Range)sheet.Rows[row];
            line.Insert();
        }

        public void RemoveRow(int row)
        {
            if (row < 1)
            {
                throw new ArgumentException("Row's number not correct.");
            }
            Range rangeRow = (Range)sheet.Rows[row];
            rangeRow.Delete();
        }

        public void DuplicateRow(int rowSource, int rowDestination)
        {
            if (rowSource < 1)
            {
                throw new ArgumentException("Can not get row from that place. Source row position can't be less than 1.");
            }
            if (rowDestination < 1)
            {
                throw new ArgumentException("Can not put row into that place. Destination row position can't be less than 1.");
            }
            Range source = sheet.Rows[rowSource, Missing.Value];
            source.Copy(sheet.Rows[rowDestination, Missing.Value]);
        }

        public void ImportTextFile(string path, object destination, string tableName, string delimiter, out int rowsCount, out int columnsCount)
        {
            string extension = string.Empty;
            try
            {
                string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(path);
                extension = Path.GetExtension(path).ToLower().Replace(".", string.Empty);
                if (string.IsNullOrEmpty(tableName))
                {
                    tableName = fileNameWithoutExtension;
                }
            }
            catch
            {
                throw new ArgumentException($"Wrong path: '{path} argument.'");
            }
            if (extension.ToLower() != "csv")
            {
                throw new ArgumentException($"Wrong path: '{path} argument.' Csv format is supported only.");
            }
            string connection = $"TEXT;{path}";
            Range target = null;
            if (destination is string && Regex.Matches((string)destination, @"[a-zA-Z]").Count > 0)
            {
                target = sheet.Range[destination];
            }
            else if (destination is System.Drawing.Point)
            {
                System.Drawing.Point p = (System.Drawing.Point)destination;
                target = sheet.Cells[p.X, p.Y] as Range;
            }
            if (target == null)
            {
                target = sheet.Cells[1, 1] as Range;
            }
            QueryTable qTable = sheet.QueryTables.Add(connection, target);
            SetDelimiter(qTable, delimiter);
            qTable.Name = tableName;
            qTable.Refresh();
            columnsCount = qTable.ResultRange.Columns.Count;
            rowsCount = qTable.ResultRange.Rows.Count;
        }

        private void SetDelimiter(QueryTable queryTable, string delimiter)
        {
            switch (delimiter.ToLower())
            {
                case "comma":
                case ",":
                    queryTable.TextFileCommaDelimiter = true;
                    break;
                case "semicolon":
                case ";":
                    queryTable.TextFileSemicolonDelimiter = true;
                    break;
                case "space":
                case " ":
                    queryTable.TextFileSpaceDelimiter = true;
                    break;
                case "tab":
                case "\t":
                    queryTable.TextFileTabDelimiter = true;
                    break;
                default:
                    queryTable.TextFileOtherDelimiter = delimiter.ToLower();
                    break;
            }
        }

        private _Worksheet GetSheetByName(string name)
        {
            if (!string.IsNullOrEmpty(name))
            {
                foreach (_Worksheet sheet in workbook.Sheets)
                {
                    if (sheet?.Name?.ToLower() == name.ToLower())
                    {
                        return sheet;
                    }
                }
            }
            return null;
        }

        private void WorkbookBeforeClose(ref bool Cancel)
        {
            ExcelManager.RemoveInstance(Id);
        }
    }
}
