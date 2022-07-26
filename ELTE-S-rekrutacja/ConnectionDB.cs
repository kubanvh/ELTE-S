using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ELTE_S_rekrutacja
{
    class ConnectionDB
    {
        public string GetConnection()
        {
        string con = "Data Source=DESKTOP-BB82HMK;Initial Catalog=ELTE-S_db;Integrated Security=True";
        return con;
        }
    }
}
