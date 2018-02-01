using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace OutlookMailService.Models
{
    public interface IMailService
    {
        void Start();
        void Stop();
    }
}
