using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using outlook = Microsoft.Office.Interop.Outlook;
using OutlookMailService;
using OutlookMailService.Models;

namespace Outluk
{
    public class Program
    {
        //Reference : In COM => Microsoft Outlook Object Library
        static void Main(string[] args)
        {
            MailService.Instance.Start();
            //Temp();
            Console.Read();
        }

        public static void Boot()
        {
            MailService.Instance.Start();
        }

        static void Temp()
        {
            outlook._Application _app = new outlook.Application();
            //outlook.MailItem mail = (outlook.MailItem)_app.CreateItem(outlook.OlItemType.olMailItem);
            //mail.To = "papa.bengu@gmail.com";
            //mail.Subject = "Subject Tester";
            //mail.Body = "Body Tester";
            //mail.BodyFormat = outlook.OlBodyFormat.olFormatHTML;
            //mail.Importance = outlook.OlImportance.olImportanceNormal;

            //((outlook._MailItem)mail).Send();

            var outlookNamespace = _app.GetNamespace("MAPI");
            var inbox = outlookNamespace.GetDefaultFolder(outlook.OlDefaultFolders.olFolderInbox);

            outlook.Items items = inbox.Items;
            items.ItemAdd += (object obj) =>
            {
                //TODO: Forward mail to Gmail;
                if (obj != null)
                {

                    var mailObj = (obj as outlook.MailItem);
                    if (mailObj != null)
                    {
                        outlook.MailItem mailObj2 = (outlook.MailItem)_app.CreateItem(outlook.OlItemType.olMailItem);
                        mailObj2.To = "papa.bengu@gmail.com";
                        mailObj2.Subject = mailObj.Subject;
                        mailObj2.Body = mailObj.Body;
                        mailObj2.BodyFormat = mailObj.BodyFormat;
                        mailObj2.Importance = mailObj.Importance;
                        ((outlook._MailItem)mailObj2).Send();
                    }
                }
            };

        }
    }
}
