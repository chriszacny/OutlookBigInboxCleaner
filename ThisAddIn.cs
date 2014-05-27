using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Xml.Linq;
using OutlookAddIn1;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace OutlookInboxCleaner
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            System.Windows.Forms.Application.ThreadException +=
                new System.Threading.ThreadExceptionEventHandler(Application_ThreadException);
        }

        private void Application_ThreadException(object sender, ThreadExceptionEventArgs e)
        {
            EventLog.WriteEntry(Strings.AppName, e.Exception.GetType().ToString(), EventLogEntryType.Error);
            EventLog.WriteEntry(Strings.AppName, e.Exception.Message, EventLogEntryType.Error);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            System.Windows.Forms.Application.ThreadException -=
                new System.Threading.ThreadExceptionEventHandler(Application_ThreadException);
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
