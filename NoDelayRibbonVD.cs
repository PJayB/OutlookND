using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Win32;

namespace OutlookND
{
    public partial class NoDelayRibbonVD
    {
        string delim = ",";

        private void NoDelayRibbonVD_Load(object sender, RibbonUIEventArgs e)
        {
            try
            {
                delim = Registry.GetValue("HKEY_CURRENT_USER\\Control Panel\\International", "sList", delim) as string;
            }
            catch (System.Exception) { }
        }

        private void sendNoDelay_Click(object sender, RibbonControlEventArgs e)
        {
            var inspector = Globals.ThisAddIn.Application.ActiveInspector();
            var mailItem = inspector.CurrentItem as _MailItem;
            if (!mailItem.Sent)
            {
                mailItem.Categories += delim + "Send Immediately";
                mailItem.DeferredDeliveryTime = DateTime.MinValue;
                mailItem.Send();
            }
        }
    }
}
