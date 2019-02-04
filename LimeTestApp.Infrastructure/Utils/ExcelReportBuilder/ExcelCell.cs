using DocumentFormat.OpenXml.Spreadsheet;

namespace LimeTestApp.Infrastructure.Utils.ExcelReportBuilder
{
    public enum ExcelCellType { Number, Bool, String, Date, Formula }

    public class ExcelCell
    {
        public object Value { get; private set; }
        public ExcelCellType ValueType { get; private set; }
        public int? StyleIndex { get; set; }


	     public ExcelCell(object val, ExcelCellType type = ExcelCellType.String, int? styleIndex = null)
	     {
		    Value = val; ValueType = type; StyleIndex = styleIndex; 
	     }

        public void SetValue(object val, ExcelCellType type = ExcelCellType.String) { Value = val; ValueType = type; }
        public void SetStyle(int styleIndex) { StyleIndex = styleIndex; }

        internal Cell Get(int? style = null)
        {
            if (Value != null) {
                if (ValueType == ExcelCellType.Number)
                    return Value.GenerateNumberCell(style ?? StyleIndex);
                if (ValueType == ExcelCellType.Bool)
                    return Value.GenerateBoolCell(style ?? StyleIndex);
                if (ValueType == ExcelCellType.Date)
                    return Value.GenerateDateTimeCell(style ?? StyleIndex);
                if (ValueType == ExcelCellType.Formula)
                    return Value.GenerateFormulaCell(style ?? StyleIndex);
            }

            return Value.GenerateStringCell(style ?? StyleIndex);
        }
    }
}
