using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;



namespace Optim8_Staffing_Sheets
{
    
    static class Program
    {
        public static bool parkServices;
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]        
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new O8SS());
            //Application.Run(new pleasestandby());
        }
    }
}
