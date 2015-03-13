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

        private void SendMessageAt(DateTime when)
        {
            try
            {
                var inspector = Globals.ThisAddIn.Application.ActiveInspector();
                var mailItem = inspector.CurrentItem as _MailItem;
                if (!mailItem.Sent)
                {
                    mailItem.Categories += delim + ThisAddIn.CategoryTag;
                    mailItem.DeferredDeliveryTime = when;
                    mailItem.Send();
                }
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(
                    ex.Message,
                    "Microsoft Outlook",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Warning);
            }
        }

        private void sendNoDelay_Click(object sender, RibbonControlEventArgs e)
        {
            SendMessageAt(DateTime.MinValue);
        }

        private void btn1m_Click(object sender, RibbonControlEventArgs e)
        {
            SendMessageAt(DateTime.Now.Add(TimeSpan.FromMinutes(1)));
        }

        private void btn5m_Click(object sender, RibbonControlEventArgs e)
        {
            SendMessageAt(DateTime.Now.Add(TimeSpan.FromMinutes(5)));
        }
    }
}
