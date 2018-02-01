using OutlookMailService.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;

namespace OutlookMailService
{
    public partial class OutlookMailService : ServiceBase
    {
        public OutlookMailService()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            MailService.Instance.Start();
        }

        protected override void OnStop()
        {
            MailService.Instance.Stop();
        }
    }
}
