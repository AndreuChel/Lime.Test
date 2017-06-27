using LimeTestApp.Core.Injection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using System.Web;
using System.Web.Mvc;
using LimeTest.Models;
using LimeTestApp.Data.NorthwindDataContext;
using System.IO;

namespace LimeTest.Controllers
{
    public class HomeController : Controller
    {
        [HttpGet]
        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public ActionResult Index(HomeIndexFormModel model)
        {
            try
            {
                if (ModelState.IsValid) //Валидация формы
                {
                    //Получаем класс отчета SalesReport
                    var salesRep = NinjectResolver.Get<LimeTestApp.Reports.Infrastruture.ReportBase>("SalesReport");

                    //Строим отчет SalesReport в поток resultStream
                    using (var resultStream = salesRep.BuildToStream(model.StartDate, model.EndDate))
                    {
                        resultStream.Position = 0;
                        
                        //отправляем отчет по почте
                        MailMessage mail = new MailMessage();
                        mail.To.Add(model.Email);
                        mail.Subject = salesRep.Title;
                        mail.Attachments.Add(new Attachment(resultStream, salesRep.FileName, "text/csv"));

                        //Настройка smtp для отправки в web.config
                        (new SmtpClient()).Send(mail);
                    }

                    return RedirectToAction("Success");
                }
            }
            catch (Exception ex) { return View("Error", ex); }

            return View();
        }

        public ActionResult Success()
        {
            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }

    }
}