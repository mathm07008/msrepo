using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using System.Windows.Forms;
using System.Data.OleDb;

namespace WordAddIn1
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e) {
            OleDbConnection cnn = new OleDbConnection();
            cnn.ConnectionString = "Provider = sqloledb; Data Source = 192.168.0.102; Initial Catalog = QDocsPelekis; User Id = sa; Password = sup3rn0v@;";
            string sql = "SELECT TOP 1 UserName FROM tblUser ";
            OleDbCommand command = new OleDbCommand(sql, cnn);
            cnn.Open();
            OleDbDataReader reader = command.ExecuteReader();
            while (reader.Read()) {
                MessageBox.Show(reader[0].ToString());
            }
        }
        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject() {
            return new Ribbon2();
        }

        
        


        private void ThisAddIn_Shutdown(object sender, System.EventArgs e) {

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
