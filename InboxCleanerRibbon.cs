using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;
using System.Threading;
using System.Runtime.InteropServices;
using System.ComponentModel;
using Exception = System.Exception;
using OutlookAddIn1;

namespace OutlookInboxCleaner
{
    public partial class InboxCleanerRibbon
    {
        private MAPIFolder _selectedFolder;
        private NameSpace _mapi;
        private MAPIFolder _deletedItems;
        private BackgroundWorker _worker;

        private void InboxCleanerRibbon_Load(object sender, RibbonUIEventArgs e)
        { }

        private void btnCleanup_Click(object sender, RibbonControlEventArgs e)
        {
            _selectedFolder = Globals.ThisAddIn.Application.ActiveExplorer().CurrentFolder;
            _mapi = Globals.ThisAddIn.Application.GetNamespace(Strings.MapiNamespace);
            _deletedItems = _mapi.GetDefaultFolder(OlDefaultFolders.olFolderDeletedItems);
            DialogResult result = MessageBox.Show(String.Format(Strings.AlertMessage, _selectedFolder.Name, batchSize.Text), Strings.Alert,
                            MessageBoxButtons.OKCancel);
            if (result == DialogResult.OK)
            {
                btnCleanup.Enabled = false;
                _worker = new BackgroundWorker();
                _worker.DoWork += t_Run;
                _worker.RunWorkerCompleted += t_RunCompleted;
                _worker.WorkerSupportsCancellation = true;
                _worker.RunWorkerAsync();
            }
        }

        private void t_RunCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            btnStop.Enabled = false;
            btnCleanup.Enabled = true;
            btnStop.Label = Strings.Stop;
            Marshal.ReleaseComObject(_mapi);
            Marshal.ReleaseComObject(_deletedItems);
            Marshal.ReleaseComObject(_selectedFolder);
            _selectedFolder = null;
        }

        private void t_Run(object sender, DoWorkEventArgs e)
        {
            btnStop.Enabled = true;
            var mailItem = _selectedFolder.Items.GetFirst() as MailItem;
            while (mailItem != null && !_worker.CancellationPending)
            {
                for (int i = 0; i < Int32.Parse(batchSize.Text); i++ )
                {
                    if (mailItem != null && !_worker.CancellationPending)
                    {
                        UserProperties properties = mailItem.UserProperties;
                        UserProperty cleanupid = properties.Add(Strings.CleanupIdProperty, OlUserPropertyType.olText);
                        UserDefinedProperty folderProperty = _deletedItems.UserDefinedProperties.Add(Strings.CleanupIdProperty, OlUserPropertyType.olText);
                        try
                        {
                            cleanupid.Value = Guid.NewGuid().ToString();
                            mailItem.Save();
                            mailItem.Delete();

                            if (deletePermanently.Checked)
                            {
                                var deletedMailItem = _deletedItems.Items.Find(String.Format("[{0}] = '{1}'", Strings.CleanupIdProperty, cleanupid.Value));
                                deletedMailItem.Delete();
                                Marshal.ReleaseComObject(deletedMailItem);
                            }
                        }
                        catch(COMException ex)
                        {
                            EventLog.WriteEntry(Strings.AppName, ex.GetType().ToString(), EventLogEntryType.Error);
                            EventLog.WriteEntry(Strings.AppName, ex.Message, EventLogEntryType.Error);
                        }
                        finally 
                        {
                            folderProperty.Delete();
                            Marshal.ReleaseComObject(properties);
                            Marshal.ReleaseComObject(cleanupid);
                            Marshal.ReleaseComObject(folderProperty);
                            Marshal.ReleaseComObject(mailItem);
                        }
                        mailItem = _selectedFolder.Items.GetFirst() as MailItem;
                    }
                    else
                    {
                        break;
                    }
                }
                Thread.Sleep(3000);
            }
        }

        private void btnStop_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                _worker.CancelAsync();
                btnStop.Label = Strings.Stopping;
                btnStop.Enabled = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}
