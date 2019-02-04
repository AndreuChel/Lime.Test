using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace LimeTestApp.Infrastructure.Utils.ExcelReportBuilder
{
    /*
     * Для построения отчетов используется библиотека "DocumentFormat.OpenXml".
     * Т.е. чтобы добавлять закрепленные области для документов, или переопределять ExcelReportUtils.DefaultStylesheet в других проектах
     * надо ее подключать.
     * 
     * для задания стилей ячеек необходимо переопределить ReportStylesheet (объекта ExcelReportBuilder)
     * или модифицировать ExcelReportUtils.DefaultStylesheet (чтобы другие отчеты не пострадали, можно только добавлять объекты). 
     
     */

    public class ExcelReport
    {
        public List<ExcelDocument> Documents { get; private set; } = new List<ExcelDocument>();
        public Stylesheet ReportStylesheet { get; set; } = ExcelReportUtils.DefaultStylesheet;
        public bool IsBuilded { get; private set; }

        public ExcelDocument AddDocument(string name)
        {
            var doc = new ExcelDocument(name);
            Documents.Add(doc);
            return doc;
        }

        public void Save(Stream stream)
        {
            SpreadsheetDocument spreadsheetDocument;
            spreadsheetDocument = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook,true);

            // Add a WorkbookPart to the document.
            WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();

            if (ReportStylesheet != null)
            {
                WorkbookStylesPart wbsp = workbookpart.AddNewPart<WorkbookStylesPart>();
                wbsp.Stylesheet = ReportStylesheet;
                wbsp.Stylesheet.Save();
            }

            workbookpart.Workbook = new Workbook();

            var worksheetPartDict = Documents.ToDictionary(d => d, d=>
            {
                WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet();

                //рисуем закрепленные области
                if (d.FixedArea != null)
                {
                    var sheetViews = new SheetViews();
                    var sheetView = new SheetView { TabSelected = false, WorkbookViewId = 0U };
                    sheetView.Append(d.FixedArea);
                    sheetViews.Append(sheetView);
                    worksheetPart.Worksheet.Append(sheetViews);
                }

                var cols = d.ColumnsPart;
                if (cols.Any()) worksheetPart.Worksheet.AppendChild(cols);

                worksheetPart.Worksheet.AppendChild(d.Get());

                worksheetPart.Worksheet.Save();
                return worksheetPart;
            });

            // Add Sheets to the Workbook.
            var sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild(new Sheets());

            uint cx = 0;
            sheets.Append(
                worksheetPartDict.Select(w => new Sheet() {
                    Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(w.Value), SheetId = new UInt32Value(++cx), Name = w.Key.Name
                }).ToList()
            );

            workbookpart.Workbook.Save();
            spreadsheetDocument.Close();

            IsBuilded = true;
        }
    }
}
