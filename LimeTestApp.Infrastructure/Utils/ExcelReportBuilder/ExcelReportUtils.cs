using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Globalization;

namespace LimeTestApp.Infrastructure.Utils.ExcelReportBuilder
{
    internal static class ExcelReportUtils
    {
        internal static Cell GenerateStringCell(this object obj, int? stileindex = null)
        {
            return new Cell()
            {
                CellValue = new CellValue(obj as string),
                DataType = CellValues.String,
                StyleIndex = UInt32Value.FromUInt32((uint)stileindex.GetValueOrDefault(0))
            };
        }

        internal static Cell GenerateFormulaCell(this object obj, int? stileindex = null)
        {
            return new Cell()
            {
                CellValue = new CellValue("0"),
                CellFormula = new CellFormula() { Text = obj as string },
                StyleIndex = UInt32Value.FromUInt32((uint)stileindex.GetValueOrDefault(0))
            };
        }

        internal static Cell GenerateNumberCell(this object obj, int? stileindex = null)
        {
            return new Cell
            {
                DataType = new EnumValue<CellValues>(CellValues.Number),
                CellValue = new CellValue(Convert.ToString(obj).Replace(",", ".")),
                StyleIndex = UInt32Value.FromUInt32((uint)stileindex.GetValueOrDefault(0))
            };
        }

        internal static Cell GenerateBoolCell(this object obj, int? stileindex = null)
        {
            if (!(obj is bool)) return "".GenerateStringCell();
            return (((bool)obj) ? 1 : 0).GenerateNumberCell(stileindex);
        }

        internal static Cell GenerateDateTimeCell(this object obj,int? stileindex = null)
        {
            if (obj == null || !(obj is DateTime)) return "".GenerateStringCell();
            var date = (DateTime)obj;
            return new Cell {
                CellValue = new CellValue((date.ToOADate()).ToString(CultureInfo.InvariantCulture)),
                DataType = new EnumValue<CellValues>(CellValues.Number),
                StyleIndex = 7 //спец. стиль для формат ячейки Дата
            };
        }


        internal static Stylesheet DefaultStylesheet
        {
            get
            {
                /*
                 * Здесь создается XML.
                 * Необходимо добавить нужные объекты в соответствующие массивы (например fonts1)
                 * Сама таблица стилей - cellFormats1
                */

                Stylesheet stylesheet1 = new Stylesheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
                stylesheet1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
                stylesheet1.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");

                Fonts fonts1 = new Fonts() { Count = (UInt32Value)1U, KnownFonts = true };

                Font font1 = new Font(
                    new FontSize() { Val = 11 },
                    //new Color() { Rgb = new HexBinaryValue() { Value = "000000" } },
                    new FontName() { Val = "Calibri" });
                Font font2 = new Font(
                    new Bold(),
                    new FontSize() { Val = 11 },
                    //new Color() { Rgb = new HexBinaryValue() { Value = "000000" } },
                    new FontName() { Val = "Calibri" });

                fonts1.Append(font1);
                fonts1.Append(font2);

                Fills fills1 = new Fills() { Count = (UInt32Value)7U };

                // FillId = 0
                Fill fill1 = new Fill();
                PatternFill patternFill1 = new PatternFill() { PatternType = PatternValues.None };
                fill1.Append(patternFill1);

                // FillId = 1
                Fill fill2 = new Fill();
                PatternFill patternFill2 = new PatternFill() { PatternType = PatternValues.Gray125 };
                fill2.Append(patternFill2);

                // FillId = 2,RED
                Fill fill3 = new Fill();
                PatternFill patternFill3 = new PatternFill() { PatternType = PatternValues.Solid };
                ForegroundColor foregroundColor1 = new ForegroundColor() { Rgb = "C0C0C0" };
                BackgroundColor backgroundColor1 = new BackgroundColor() { Indexed = (UInt32Value)64U };
                patternFill3.Append(foregroundColor1);
                patternFill3.Append(backgroundColor1);
                fill3.Append(patternFill3);

                // FillId = 3,BLUE
                Fill fill4 = new Fill();
                PatternFill patternFill4 = new PatternFill() { PatternType = PatternValues.Solid };
                ForegroundColor foregroundColor2 = new ForegroundColor() { Rgb = "FF0070C0" };
                BackgroundColor backgroundColor2 = new BackgroundColor() { Rgb = "FFFFFF00"/*, Indexed = (UInt32Value)64U*/ };
                patternFill4.Append(foregroundColor2);
                patternFill4.Append(backgroundColor2);
                fill4.Append(patternFill4);

                // FillId = 4,YELLO
                Fill fill5 = new Fill();
                PatternFill patternFill5 = new PatternFill() { PatternType = PatternValues.Solid };
                ForegroundColor foregroundColor3 = new ForegroundColor() { Rgb = "FFFFFF00" };
                BackgroundColor backgroundColor3 = new BackgroundColor() { Indexed = (UInt32Value)64U };
                patternFill5.Append(foregroundColor3);
                patternFill5.Append(backgroundColor3);
                fill5.Append(patternFill5);

                // FillId = 5,lightcyan
                Fill fill6 = new Fill();
                PatternFill patternFill6 = new PatternFill() { PatternType = PatternValues.Solid };
                ForegroundColor foregroundColor4 = new ForegroundColor() { Rgb = "FFe0ffff" };
                BackgroundColor backgroundColor4 = new BackgroundColor() { Indexed = (UInt32Value)64U };
                patternFill6.Append(foregroundColor4);
                patternFill6.Append(backgroundColor4);
                fill6.Append(patternFill6);

                // FillId = 6,lightpink
                Fill fill7 = new Fill();
                PatternFill patternFill7 = new PatternFill() { PatternType = PatternValues.Solid };
                ForegroundColor foregroundColor5 = new ForegroundColor() { Rgb = "ffffb6c1" };
                BackgroundColor backgroundColor5 = new BackgroundColor() { Indexed = (UInt32Value)64U };
                patternFill7.Append(foregroundColor5);
                patternFill7.Append(backgroundColor5);
                fill7.Append(patternFill7);

                fills1.Append(fill1);
                fills1.Append(fill2);
                fills1.Append(fill3);
                fills1.Append(fill4);
                fills1.Append(fill5);
                fills1.Append(fill6);
                fills1.Append(fill7);

                Borders borders1 = new Borders() { Count = (UInt32Value)1U };

                Border border1 = new Border();
                LeftBorder leftBorder1 = new LeftBorder();
                RightBorder rightBorder1 = new RightBorder();
                TopBorder topBorder1 = new TopBorder();
                BottomBorder bottomBorder1 = new BottomBorder();
                DiagonalBorder diagonalBorder1 = new DiagonalBorder();

                border1.Append(leftBorder1);
                border1.Append(rightBorder1);
                border1.Append(topBorder1);
                border1.Append(bottomBorder1);
                border1.Append(diagonalBorder1);

                borders1.Append(border1);

                CellStyleFormats cellStyleFormats1 = new CellStyleFormats() { Count = (UInt32Value)1U };
                CellFormat cellFormat1 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U };

                cellStyleFormats1.Append(cellFormat1);

                CellFormats cellFormats1 = new CellFormats() { Count = (UInt32Value)8U };
                CellFormat cellFormat2 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U };
                CellFormat cellFormat3 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)1U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFill = true };
                CellFormat cellFormat4 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)3U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFill = true };
                CellFormat cellFormat5 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFill = true };
                CellFormat cellFormat6 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)5U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFill = true };
                CellFormat cellFormat7 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)6U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFill = true };

                CellFormat cellFormat8 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFill = true, ApplyAlignment = true };
                cellFormat8.Append (new Alignment() { TextRotation = (UInt32Value)90U });
                CellFormat cellFormat9 = new CellFormat() { NumberFormatId = (UInt32Value)14U,FontId = (UInt32Value)0U,FillId = (UInt32Value)0U,BorderId = (UInt32Value)0U,FormatId = (UInt32Value)0U, ApplyNumberFormat = true }; //спец. стиль для формат ячейки Дата

                cellFormats1.Append(cellFormat2);
                cellFormats1.Append(cellFormat3);
                cellFormats1.Append(cellFormat4);
                cellFormats1.Append(cellFormat5);
                cellFormats1.Append(cellFormat6);
                cellFormats1.Append(cellFormat7);
                cellFormats1.Append(cellFormat8);
                cellFormats1.Append(cellFormat9);

                CellStyles cellStyles1 = new CellStyles() { Count = (UInt32Value)1U };
                CellStyle cellStyle1 = new CellStyle() { Name = "Normal", FormatId = (UInt32Value)0U, BuiltinId = (UInt32Value)0U };

                cellStyles1.Append(cellStyle1);
                DifferentialFormats differentialFormats1 = new DifferentialFormats() { Count = (UInt32Value)0U };
                TableStyles tableStyles1 = new TableStyles() { Count = (UInt32Value)0U, DefaultTableStyle = "TableStyleMedium2", DefaultPivotStyle = "PivotStyleMedium9" };

                StylesheetExtensionList stylesheetExtensionList1 = new StylesheetExtensionList();

                StylesheetExtension stylesheetExtension1 = new StylesheetExtension() { Uri = "{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" };
                stylesheetExtension1.AddNamespaceDeclaration("x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");

                stylesheetExtensionList1.Append(stylesheetExtension1);

                stylesheet1.Append(fonts1);
                stylesheet1.Append(fills1);
                stylesheet1.Append(borders1);
                stylesheet1.Append(cellStyleFormats1);
                stylesheet1.Append(cellFormats1);
                stylesheet1.Append(cellStyles1);
                stylesheet1.Append(differentialFormats1);
                stylesheet1.Append(tableStyles1);
                stylesheet1.Append(stylesheetExtensionList1);
                return stylesheet1;
            }
        }
    }
}
