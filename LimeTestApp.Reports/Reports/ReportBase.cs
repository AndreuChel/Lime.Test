using LimeTestApp.Data.NorthwindDb;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LimeTestApp.Reports.Reports
{
    /// <summary>
    /// Базовый класс отчетов
    /// </summary>
    public abstract class ReportBase
    {
        protected readonly INorthwindContext NorthwindContext;
        public ReportBase(INorthwindContext _dc) { NorthwindContext = _dc; }

        // Построение отчета в MemoryStream
        public abstract Stream BuildToStream(params object[] _args);
        
        public abstract string Title { get; }

        public abstract string FileName { get; }
    }
}
