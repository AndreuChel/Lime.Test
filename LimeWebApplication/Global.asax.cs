using LimeTestApp.Core.Injection;
using LimeTestApp.Data.NorthwindDataContext;
using LimeTestApp.Reports;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Optimization;
using System.Web.Routing;

namespace LimeTest
{
    public class MvcApplication : System.Web.HttpApplication
    {
        protected void Application_Start()
        {
            AreaRegistration.RegisterAllAreas();
            FilterConfig.RegisterGlobalFilters(GlobalFilters.Filters);
            RouteConfig.RegisterRoutes(RouteTable.Routes);
            BundleConfig.RegisterBundles(BundleTable.Bundles);

            //Загрузка модулей Ninject 
            NinjectResolver.LoadModule<ReportLoadModule>(); //Отчеты

            //Получаем строку подключения
            System.Configuration.Configuration rootWebConfig =
            System.Web.Configuration.WebConfigurationManager.OpenWebConfiguration("/MyWebSiteRoot");
            System.Configuration.ConnectionStringSettings connString = null;
            if (rootWebConfig.ConnectionStrings.ConnectionStrings.Count > 0)
                connString = rootWebConfig.ConnectionStrings.ConnectionStrings["NorthwindConnectionString"];
            
            //Создаем контекст подключения к БД (строка подключения из web.config)
            NinjectResolver.Bind<NorthwindModel>().To<NorthwindModel>()
                           .WithConstructorArgument("conn", connString?.ConnectionString ?? "");

        }
    }
}
