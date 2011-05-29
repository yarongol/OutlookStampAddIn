using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Configuration;

namespace StampAddIn
{
    public partial class ThisAddIn
    {
        public String strName;
        public String strColor;

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
             return new StampRibbon1();
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Configuration config =
                ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);

            AppSettingsSection appSettingSection =
                (AppSettingsSection)config.GetSection("appSettings");
            try
            {
                strName = appSettingSection.Settings["StampName"].Value;
                strColor = appSettingSection.Settings["StampColor"].Value;
            }
            catch (NullReferenceException) //Name setting does not exist, prompt a dialog
            {
                dlgUserName dlg = new dlgUserName();
                dlg.ShowDialog();
                strName = dlg.txtUserName.Text;
                System.Drawing.Color c = dlg.txtUserName.ForeColor;
                int i;

                // Unfortunately, .Net and Word use different Color representations.
                // .Net uses ARGB. Word seems to use BGR. Home made conversion:
                i = c.B;
                i = i * 256 + c.G;
                i = i * 256 + c.R;
                strColor = i.ToString();

                config.AppSettings.Settings.Add("StampName", strName);
                config.AppSettings.Settings.Add("StampColor", strColor);
                config.Save(ConfigurationSaveMode.Modified); // Save the configuration file.    
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
