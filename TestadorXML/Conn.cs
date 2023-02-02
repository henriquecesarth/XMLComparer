using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace TestadorXML
{
    public class Conn
    {
        public static string Server { get; set; }
        public static string DataBase { get; set; }
        
        private static string User = "sa";
        public static string Password { get; set; }



        public static string StrCon
        {

            get {
                return $"Data Source={Server}; Integrated Security=False;Initial Catalog={DataBase}; " +
                    $"User ID={User}; Password={Password}";
            }

        }

        
    }


}
