﻿using System;
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
    class Chair : Traveler
    {
        //===========================
        // PUBLIC
        //===========================

        // Doesn't do anything
        public Chair() : base() {}
        // Gets the base properties and orders of the traveler from a json string
        public Chair(string json) : base(json) {
            GetBlacklist();
        }
        // Creates a traveler from a part number and quantity
        public Chair(string partNo, int quantity) : base(partNo, quantity) {
            GetBlacklist();
        }
        // Creates a traveler from a part number and quantity, then loads the bill of materials
        public Chair(string partNo, int quantity, OdbcConnection MAS) : base(partNo, quantity, MAS)
        {
            GetBlacklist();
        }
        //===========================
        // Private
        //===========================
        private void GetBlacklist()
        {
            m_blacklist.Add(new BlacklistItem(Blacklist.StartsWith, "/")); // Misc work items
        }
    }
}
