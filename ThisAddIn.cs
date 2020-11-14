using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace CitiMeetTracker
{
    public partial class ThisAddIn
    {
        private UserControl1 taskPaneControl1;
        private Microsoft.Office.Tools.CustomTaskPane taskPaneValue;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            taskPaneControl1 = new UserControl1();
            taskPaneValue = this.CustomTaskPanes.Add(taskPaneControl1, "Citi Meet Tracker");
            taskPaneValue.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
            taskPaneValue.DockPositionRestrict = Office.MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange;
            taskPaneValue.Width = 300;
            //taskPaneValue.Height = 500;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        public Microsoft.Office.Tools.CustomTaskPane TaskPane
        {
            get { return taskPaneValue; }
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
