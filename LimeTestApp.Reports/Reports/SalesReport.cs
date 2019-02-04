
using LimeTestApp.Data.NorthwindDb;
using LimeTestApp.Infrastructure.Utils.ExcelReportBuilder;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LimeTestApp.Reports.Reports
{
    /// <summary>
    /// Реализация отчета по продажам
    /// </summary>
    public class SalesReport : ReportBase
    {
        public SalesReport(INorthwindContext dbContext) : base(dbContext) { }

        public override string Title { get; } = "Отчет по продажам"; 
        public override string FileName { get; } = "SalesRep.xlsx";

        public override Stream BuildToStream(params object[] args)
        {
            //Проверяем аргументы
            if (args != null 
                && (args.Count() > 2 
                    || args.Where(a => a != null).Any( a => !(a is DateTime))))
                throw new ArgumentException("SalesReport.BuildToStream: invalid arguments");

            var sDate = args != null && args.Length >= 1 && args[0] != null 
                ? (args[0] as DateTime?).Value : DateTime.MinValue;

            var eDate = args != null && args.Length >= 2 && args[1] != null 
                ? (args[1] as DateTime?).Value : DateTime.MaxValue;

            if (sDate > eDate)
            {
                var tmpDate = eDate; eDate = sDate; sDate = tmpDate;
            }

            var result = new MemoryStream();

            //создаем объект отчета
            var report = new ExcelReport();
            //создаем вкладку
            var excelsheet = report.AddDocument("Продажи");
            //Задаем ширины колонок
            excelsheet.SetColumnsWidth(10, 25, 10, 40, 10, 10);

            //Шапка таблицы
            var rHead = new ExcelRow();
            rHead.SetStyle(1, true);
            rHead.Append( new ExcelCell("OrderId"), 
                            new ExcelCell("OrderDate"), 
                            new ExcelCell("ProductID"), 
                            new ExcelCell("ProductName"), 
                            new ExcelCell("Quantity"), 
                            new ExcelCell("UnitPrice"),
                            new ExcelCell("CalcField"));
            excelsheet.Append(rHead);

            //Построение тела таблицы
            var cx = 1;
            NorthwindContext.Order
                .Where(o => o.OrderDate.HasValue && o.OrderDate.Value >= sDate && o.OrderDate.Value <= eDate)
                .SelectMany(o => o.OrderDetail)
                .ToList()
                .ForEach(o => {
                    cx++;
                    excelsheet.Append(new ExcelRow(new ExcelCell(o.OrderID, ExcelCellType.Number),
                                                    new ExcelCell(o.Order.OrderDate, ExcelCellType.Date),
                                                    new ExcelCell(o.ProductID, ExcelCellType.Number),
                                                    new ExcelCell(o.Product.Name),
                                                    new ExcelCell(o.Quantity, ExcelCellType.Number),
                                                    new ExcelCell(o.UnitPrice, ExcelCellType.Number),
                                                    new ExcelCell($"E{cx}*F{cx}", ExcelCellType.Formula)));
                });

            report.Save(result);

            return result;
        }
    }
}
