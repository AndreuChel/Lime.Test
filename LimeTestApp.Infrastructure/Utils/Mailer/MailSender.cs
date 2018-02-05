using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace LimeTestApp.Infrastructure.Utils.Mailer
{
    public class MailSender : IMailSender
    {
        public void Send(MailMessage message)
        {
            //Настройка smtp для отправки в web.config
            (new SmtpClient()).Send(message);
        }
    }
}
