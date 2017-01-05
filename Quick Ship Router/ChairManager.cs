using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.Odbc;
using System.Diagnostics;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using Marshal = System.Runtime.InteropServices.Marshal;
using System.Drawing.Printing;

namespace Quick_Ship_Router
{
    class ChairManager
    {
        private Excel.Worksheet m_;
        private List<Chair> m_chairs = new List<Chair>();
        private OdbcConnection MAS = new OdbcConnection();
    }
}
