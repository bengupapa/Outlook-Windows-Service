using System;
using Microsoft.Office.Interop.Outlook;
using static OutlookMailService.Constants.Constants;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Net;
using System.Net.Mail;

namespace OutlookMailService.Models
{
    public class MailService : IMailService
    {
        //installutil.exe "C:\Users\bengupj\Desktop\Pj Stuff\Pj Tools\Outluk\OutlookMailService\bin\Release\OutlookMailService.exe"
        private static _Application _outlookApp = null;
        private static IMailService _mailService = null;

        private MailService()
        {
            try
            {
                Process[] processes = Process.GetProcessesByName(OutlookProcess);
                _outlookApp = processes.Any() ? (Marshal.GetActiveObject(OutlookProgID) as Application) : new Application();
            }
            catch
            {
                _outlookApp = new Application();
            }
        }

        public static IMailService Instance
        {
            get
            {
                return _mailService ?? (_mailService = new MailService());
            }
        }

        public void Start()
        {
            NameSpace outlookNamespace = _outlookApp.GetNamespace(MailProtocol);
            MAPIFolder inbox = outlookNamespace.GetDefaultFolder(OlDefaultFolders.olFolderInbox);

            Items items = inbox.Items;
            items.ItemAdd += (object obj) =>
            {
                MailItem mailObj = (MailItem)obj;
                if (mailObj == null) return;


                MailItem message = (MailItem)_outlookApp.CreateItem(OlItemType.olMailItem);
                message.To = ToAddress;
                message.Subject = mailObj.Subject;
                message.Body = mailObj.Body;
                message.BodyFormat = mailObj.BodyFormat;
                message.Importance = mailObj.Importance;

                ((_MailItem)message).Send();
            };
        }

        public void Stop()
        {
            if (_mailService != null)
            {
                _outlookApp.Quit();
                _outlookApp = null;
                _mailService = null;
            }
        }
    }
}
