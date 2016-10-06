using log4net.Config;
using log4net;
using System;

namespace SE.UnityCommentsExcel
{
    public partial class ThisAddIn
    {
        private static ILog _Log = LogManager.GetLogger("Unity comment Excel Addins");
        /// <summary>
        /// Handles the Startup event of the ThisAddIn control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.EventArgs"/> instance containing the event data.</param>
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                
                XmlConfigurator.Configure();
                UnityCommentRibbon.Log = _Log;
                _Log.Info("Unity Comments Excel addins Loaded");
            }
            catch(Exception  ex)
            {
                _Log.Fatal(ex);
            }
        }

        /// <summary>
        /// Handles the Shutdown event of the ThisAddIn control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.EventArgs"/> instance containing the event data.</param>
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            _Log.Info("Unity Comments Excel unloaded");
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
