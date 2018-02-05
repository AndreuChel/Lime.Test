using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace LimeTestApp.Infrastructure.Utils.Mailer
{
    public interface IMailSender
    {
        void Send(MailMessage message);
    }
}
