using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Optim8_Staffing_Sheets
{
    class ps_ride
    {
        public string m_name = "";
        public ps_shift[] m_shift = new ps_shift[5];
        public int m_listUsed = 0;

        public ps_ride(string name)
        {
            for (int i = 0; i < 5; i++)
            {
                m_shift[i] = new ps_shift();
            }
            m_name = name;
        }
    }
}
