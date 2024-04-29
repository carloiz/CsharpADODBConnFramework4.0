using ADODB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CsharpADODBConnection
{
    internal class ADODBConn
    {
        //public static Connection connect;

        public static Connection conn()
        {
            Connection conn = new Connection();

            conn.Open("Driver=MySQL ODBC 8.0 ANSI driver;" +
                        "server=127.0.0.1;" +
                        "uid=root;" +
                        "password=;" +
                        "database=records;", null, null, 0);
            return conn;
        }

    }
}


