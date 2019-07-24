using System;
using G1ANT.Language;

namespace G1ANT.Addon.MSOffice.Helpers
{
    public static class ExcelHelper
    {
        public static object GetColumn(IntegerStructure columnIndex, TextStructure columnName, bool isColumnRequired = false)
        {
            if (columnIndex != null)
                return columnIndex.Value;
            if (columnName != null)
                return columnName.Value;
            if(isColumnRequired)
                throw new ArgumentException("One of the ColIndex or ColName arguments have to be set up.");

            return null;
        }
    }
}
