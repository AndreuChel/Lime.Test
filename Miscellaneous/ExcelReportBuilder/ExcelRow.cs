using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;
using System.Linq;

namespace ExcelReportBuilder
{
    public class ExcelRow
    {
        public ExcelRow () { }
        public ExcelRow(params ExcelCell[] newCells) { Append(newCells); }
        public ExcelRow(IEnumerable<ExcelCell> newCells) { Append(newCells); }

        public int? StyleIndex { get; set; }
        public bool UseStyleForAllCells { get; private set; } = false;

        public List<ExcelCell> Cells { get; private set; } = new List<ExcelCell>();
        public void SetStyle(int _styleIndex, bool _useForAll = false) { StyleIndex = _styleIndex; UseStyleForAllCells = _useForAll; }

        public void Append(params ExcelCell[] newCells) { Cells.AddRange(newCells); }
        public void Append(IEnumerable<ExcelCell> newCells) { Cells.AddRange(newCells); }

        internal Row get(int? _style = null) {
            return new Row(Cells.Select(c => c.get(_style ?? (UseStyleForAllCells? StyleIndex: null))));
        }
    }
}
