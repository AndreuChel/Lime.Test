using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Ninject;
using LimeTestApp.Core.Injection;
using LimeTestApp.Reports.Infrastruture;
using LimeTestApp.Reports.Infrastruture.Reports;

namespace LimeTestApp.Reports
{
    /// <summary>
    /// Модуль для загрузки в IoC-контейнер Ninject
    /// </summary>
    public class ReportLoadModule : Ninject.Modules.NinjectModule
    {
        public override void Load()
        {
            NinjectResolver.Bind<ReportBase, SalesReport>().Named("SalesReport");
        }
    }
}
