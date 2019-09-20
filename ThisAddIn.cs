using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;

namespace OutlookAddIn5
{
    public partial class ThisAddIn
    {
        private Outlook.Inspectors _appInspectors;
        ThisRibbonCollection ribbonCollection;
        private Outlook.Explorer explorer;
        private Outlook.Application app;
        private Outlook.MailItem mailItem;


        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            explorer = this.Application.ActiveExplorer();
            app = (Outlook.Application)explorer.Application;

            _appInspectors = app.Inspectors;
            _appInspectors.NewInspector += new Outlook.InspectorsEvents_NewInspectorEventHandler(Inspectors_NewInspector);

            this.Application.ItemSend += new Microsoft.Office.Interop.Outlook.ApplicationEvents_11_ItemSendEventHandler(Application_ItemSend);
        }

        void Application_ItemSend(object Item, ref bool Cancel)
        {
            ThisRibbonCollection ribbonCollection;
            ribbonCollection = Globals.Ribbons[Globals.ThisAddIn.Application.ActiveInspector()];
            if (ribbonCollection.Ribbon1.checkBox1.Checked)
            {
                Outlook.MailItem mi = Item as Outlook.MailItem;
                mi.FlagRequest = "This email is a resolution response for the case";
                mi.Save();
            }
        }

        void Inspectors_NewInspector(Microsoft.Office.Interop.Outlook.Inspector Inspector)
        {
            mailItem = Inspector.CurrentItem as Outlook.MailItem;
            ribbonCollection = Globals.Ribbons[Globals.ThisAddIn.Application.ActiveInspector()];
            if (!string.IsNullOrEmpty(mailItem.FlagRequest) && mailItem.FlagRequest.Contains("This email is a resolution response for the case"))
            {
                ribbonCollection.Ribbon1.checkBox1.Checked = true;
                ribbonCollection.Ribbon1.checkBox1.Enabled = false;
                ribbonCollection.Ribbon1.checkBox1.Visible = true;
                mailItem.Save();
            }
            else
            {
                //   Globals.Ribbons.Ribbon1.checkBox1.Visible = false;
                ribbonCollection.Ribbon1.checkBox1.Checked = false;
                ribbonCollection.Ribbon1.checkBox1.Enabled = true;
                ribbonCollection.Ribbon1.checkBox1.Visible = false;
                mailItem.Save();
            }
            GC.SuppressFinalize(this);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see http://go.microsoft.com/fwlink/?LinkId=506785
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
