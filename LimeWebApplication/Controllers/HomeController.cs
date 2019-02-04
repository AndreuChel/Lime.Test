
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using System.Web;
using System.Web.Mvc;
using LimeTest.Models;
using System.IO;
using LimeTestApp.Data.NorthwindDb;
using LimeTestApp.Reports.Reports;
using LimeTestApp.Infrastructure.Utils.Mailer;

namespace LimeTest.Controllers
{
    public class HomeController : Controller
    {
        private readonly INorthwindContext _northwindDbContext;
        private readonly IMailSender _mailSender;

        public HomeController(INorthwindContext dbContext, IMailSender mailSender)
        {
            _northwindDbContext = dbContext;
            _mailSender = mailSender;
        }

        protected override void Dispose(bool disposing)
        {
            _northwindDbContext?.Dispose();
            base.Dispose(disposing);
        }

        [HttpGet]
        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public ActionResult Index(HomeIndexFormModel model)
        {
            if (!ModelState.IsValid) return View();
            try
            {
                //Получаем класс отчета SalesReport
                var salesRep = new SalesReport(_northwindDbContext);

                //Строим отчет SalesReport в поток resultStream
                using (var resultStream = salesRep.BuildToStream(model.StartDate, model.EndDate))
                {
                    resultStream.Position = 0;
                        
                    //отправляем отчет по почте
                    MailMessage mail = new MailMessage() {
                        To = { model.Email },
                        Subject = salesRep.Title,
                        Attachments = { new Attachment(resultStream, salesRep.FileName, "text/csv") }
                    };
                    _mailSender.Send(mail);
                }
                return RedirectToAction("Success");
            }
            catch (Exception ex) { return View("Error", ex); }
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