using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;

namespace OutlookND
{
    public partial class ThisAddIn
    {
        public const string CategoryTag = "Send Immediately";
        public readonly TimeSpan DefaultSendDelay = TimeSpan.FromMinutes(5);

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Globals.ThisAddIn.Application.ItemSend += Application_ItemSend;
        }

        private void Application_ItemSend(object itemObj, ref bool cancel)
        {
            _MailItem item = itemObj as _MailItem;
            if (item != null)
            {
                if (item.Categories == null || !item.Categories.Contains(CategoryTag))
                {
                    item.DeferredDeliveryTime = DateTime.Now.Add(DefaultSendDelay);
                }
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

      

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
