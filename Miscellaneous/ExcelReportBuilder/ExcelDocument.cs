using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

using System.Collections.Generic;
using System.Linq;

namespace ExcelReportBuilder
{
    public class ExcelDocument
    {
        public string Name { get; private set; }
        public int? StyleIndex { get; set; }
        public bool UseStyleForAllRows { get; private set; } = false;
        public List<ExcelRow> Rows { get; private set; } = new List<ExcelRow>();
        private List<double> ColumnsWidth { get; set; } = new List<double>();

        public Pane FixedArea { get; set; }

        public ExcelDocument(string _name) { Name = _name; }

        public void SetStyle(int _styleIndex, bool _useForAll = false) { StyleIndex = _styleIndex; UseStyleForAllRows = _useForAll; }

        public void Append(params ExcelRow[] newRows) { Rows.AddRange(newRows); }
        public void Append(IEnumerable<ExcelRow> newRows) { Rows.AddRange(newRows); }

        public void SetColumnsWidth(params double[] widths) { ColumnsWidth.AddRange(widths); }
        internal SheetData get()
        {
            return new SheetData(Rows.Select(r=>r.get(UseStyleForAllRows? StyleIndex : null)));
        }
        internal Columns ColumnsPart
        {
            get
            {
                return new Columns(ColumnsWidth.Select((cw, _ind) => new Column { Min = (uint)_ind + 1, Max = (uint)_ind + 1, Width = cw, CustomWidth = true })); 
            }
        }

    }
}
