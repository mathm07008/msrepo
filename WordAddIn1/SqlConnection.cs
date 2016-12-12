using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MySql.Data.MySqlClient;

namespace WordAddIn1 {
    class SqlConnection {
        MySqlConnection cnn;
        String ConnectionString = "server=94.177.246.40; database=OddsComparisonNew; User Id = am; Password = sup3rn0v@;";
        
        public SqlConnection()
        {
            cnn = new MySqlConnection(ConnectionString);            
        }

        public MySqlConnection getConnection()
        {
            return cnn;
        }
        public void openConnection()
        {
            cnn.Open();
        }
    }
}
