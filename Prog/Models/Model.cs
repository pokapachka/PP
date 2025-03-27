using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Prog.Connection;

namespace Prog.Models
{
    public static class Model
    {

        public static string Username{ get; set; }
        public static string Password { get; set; }
        public static string DbName { get; set; }
        public static ConnectionDB CurrentConnection { get; set; }
        public static void CloseCurrentConnection()
        {
            if (CurrentConnection != null)
            {
                CurrentConnection.CloseConnection();
                CurrentConnection = null;
            }
            Username = null;
            Password = null;
            DbName = null;
        }
    }
}
