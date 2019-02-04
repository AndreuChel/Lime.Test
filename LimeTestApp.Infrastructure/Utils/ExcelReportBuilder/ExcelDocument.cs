using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

using System.Collections.Generic;
using System.Linq;

namespace LimeTestApp.Infrastructure.Utils.ExcelReportBuilder
{
    public class ExcelDocument
    {
        public string Name { get; private set; }

        public int? StyleIndex { get; set; }

        public bool UseStyleForAllRows { get; private set; }

        public List<ExcelRow> Rows { get; private set; } = new List<ExcelRow>();

        private List<double> ColumnsWidth { get; set; } = new List<double>();

        public Pane FixedArea { get; set; }

        public ExcelDocument(string name) { Name = name; }

        public void SetStyle(int styleIndex, bool useForAll = false) { StyleIndex = styleIndex; UseStyleForAllRows = useForAll; }

        public void Append(params ExcelRow[] newRows) { Rows.AddRange(newRows); }
        public void Append(IEnumerable<ExcelRow> newRows) { Rows.AddRange(newRows); }

        public void SetColumnsWidth(params double[] widths) { ColumnsWidth.AddRange(widths); }

        internal SheetData Get()
            => new SheetData(Rows.Select(r=>r.get(UseStyleForAllRows? StyleIndex : null)));

        internal Columns ColumnsPart
	         => new Columns(ColumnsWidth.Select((cw, ind) => new Column { Min = (uint)ind + 1, Max = (uint)ind + 1, Width = cw, CustomWidth = true })); 

    }
}
