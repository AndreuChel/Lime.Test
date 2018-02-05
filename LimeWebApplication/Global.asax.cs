

using LimeTestApp.Data.NorthwindDb;
using LimeTestApp.Infrastructure.Utils.Mailer;
using LimeTestApp.Reports;
using Ninject;
using Ninject.Web.Mvc;
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

            // внедрение зависимостей
            var kernel = new StandardKernel();
            kernel.Bind<INorthwindContext>().To<NorthwindContext>();
            kernel.Bind<IMailSender>().To<MailSender>();
            DependencyResolver.SetResolver(new NinjectDependencyResolver(kernel));

        }
    }
}
