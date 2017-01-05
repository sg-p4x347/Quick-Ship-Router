using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Data.Odbc;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Marshal = System.Runtime.InteropServices.Marshal;

namespace Quick_Ship_Router
{
    class Table : Traveler
    {
        //===========================
        // PUBLIC
        //===========================

        public Table() : base() {
        }
        // Creates a traveler from a part number and quantity
        public Table(string partNo, int quantity) : base(partNo, quantity)
        {
            GetBlacklist();
            m_colorNo = Convert.ToInt32(m_partNo.Substring(m_partNo.Length - 2));
        }
        public Table(string partNo, int quantity, OdbcConnection MAS) : base(partNo,quantity,MAS)
        {
            GetBlacklist();
        }
        //===========================
        // Private
        //===========================
        private void GetBlacklist()
        {
            m_blacklist.Add(new BlacklistItem(Method.StartsWith, "88")); // Glue items
            m_blacklist.Add(new BlacklistItem(Method.StartsWith, "92")); // Foam items
            m_blacklist.Add(new BlacklistItem(Method.StartsWith, "/")); // Misc work items
        }
        private int m_colorNo { get; set; } = 0;
        private string m_shapeNo { get; set; } = "";
        // Blank information
        private string m_blankNo = "";
        private string m_blankSize = "";
        private int m_partsPerBlank = 0;
        private int m_blankQuantity = 0;
        private int m_leftoverParts = 0;
    }
}
