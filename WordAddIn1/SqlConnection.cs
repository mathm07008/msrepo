using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordAddIn1 {
    class SqlConnection {
        OleDbConnection cnn = new OleDbConnection();
        public SqlConnection () {            
            cnn.ConnectionString = "Provider = sqloledb; Data Source = 192.168.0.102; Initial Catalog = QDocsPelekis; User Id = sa; Password = sup3rn0v@;"; 
        }
        
        public OleDbConnection getConn() {
            return cnn;
        }
    }
}
