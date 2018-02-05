using DocumentFormat.OpenXml.Spreadsheet;

namespace LimeTestApp.Infrastructure.Utils.ExcelReportBuilder
{
    public enum ExcelCellType { _number, _bool, _string, _date, _formula }

    public class ExcelCell
    {
        public object Value { get; private set; }
        public ExcelCellType ValueType { get; private set; }
        public int? StyleIndex { get; set; }


        public ExcelCell (object _val, ExcelCellType _type = ExcelCellType._string, int? _styleIndex = null) { Value = _val; ValueType = _type; StyleIndex = _styleIndex; }

        public void SetValue(object _val, ExcelCellType _type = ExcelCellType._string) { Value = _val; ValueType = _type; }
        public void SetStyle(int _styleIndex) { StyleIndex = _styleIndex; }

        internal Cell get(int? _style = null)
        {
            if (Value != null) {
                if (ValueType == ExcelCellType._number)
                    return Value.generateNumberCell(_style ?? StyleIndex);
                if (ValueType == ExcelCellType._bool)
                    return Value.generateBoolCell(_style ?? StyleIndex);
                if (ValueType == ExcelCellType._date)
                    return Value.generateDateTimeCell(_style ?? StyleIndex);
                if (ValueType == ExcelCellType._formula)
                    return Value.generateFormulaCell(_style ?? StyleIndex);
            }

            return Value.generateStringCell(_style ?? StyleIndex);
        }
    }
}
