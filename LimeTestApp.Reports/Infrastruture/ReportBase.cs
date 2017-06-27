using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LimeTestApp.Reports.Infrastruture
{
    /// <summary>
    /// Базовый класс отчетов
    /// </summary>
    public abstract class ReportBase
    {
        /// <summary>
        /// Построение отчета в MemoryStream
        /// </summary>
        /// <param name="_args"></param>
        /// <returns></returns>
        public abstract Stream BuildToStream(params object[] _args);
        
        public abstract string Title { get; }

        public abstract string FileName { get; }
    }
}
