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
    class Table : Traveler
    {
        //===========================
        // PUBLIC
        //===========================
        public Table() : base() { }
        public Table(string json) : base(json) {
            GetBlacklist();
            m_colorNo = Convert.ToInt32(m_partNo.Substring(m_partNo.Length - 2));
            m_shapeNo = m_partNo.Substring(0, m_partNo.Length - 3);
        }
        // Creates a traveler from a part number and quantity
        public Table(string partNo, int quantity) : base(partNo, quantity)
        {
            GetBlacklist();
            m_colorNo = Convert.ToInt32(m_partNo.Substring(m_partNo.Length - 2));
            m_shapeNo = m_partNo.Substring(0, m_partNo.Length - 3);
        }
        public Table(string partNo, int quantity, OdbcConnection MAS) : base(partNo,quantity,MAS)
        {
            GetBlacklist();
            m_colorNo = Convert.ToInt32(m_partNo.Substring(m_partNo.Length - 2));
            m_shapeNo = m_partNo.Substring(0, m_partNo.Length - 3);
        }
        //===========================
        // Private
        //===========================
        private void GetBlacklist()
        {
            m_blacklist.Add(new BlacklistItem("88")); // Glue items
            m_blacklist.Add(new BlacklistItem("92")); // Foam items
            m_blacklist.Add(new BlacklistItem("/")); // Misc work items
        }
        // part information
        private int m_colorNo = 0;
        private string m_shapeNo = "";
        // Blank information
        private string m_blankNo = "";
        private string m_blankColor = "";
        private string m_blankSize = "";
        private string m_sheetSize = "";
        private string m_blankComment = "";
        private int m_partsPerBlank = 0;
        private int m_blankQuantity = 0;
        private int m_leftoverParts = 0;
        // pallet
        private string m_palletSize = "";
        private int m_palletQty = 0;

        public int ColorNo
        {
            get
            {
                return m_colorNo;
            }

            set
            {
                m_colorNo = value;
            }
        }

        public string ShapeNo
        {
            get
            {
                return m_shapeNo;
            }

            set
            {
                m_shapeNo = value;
            }
        }

        public string BlankNo
        {
            get
            {
                return m_blankNo;
            }

            set
            {
                m_blankNo = value;
            }
        }

        public string BlankSize
        {
            get
            {
                return m_blankSize;
            }

            set
            {
                m_blankSize = value;
            }
        }

        public int PartsPerBlank
        {
            get
            {
                return m_partsPerBlank;
            }

            set
            {
                m_partsPerBlank = value;
            }
        }

        public int BlankQuantity
        {
            get
            {
                return m_blankQuantity;
            }

            set
            {
                m_blankQuantity = value;
            }
        }

        public int LeftoverParts
        {
            get
            {
                return m_leftoverParts;
            }

            set
            {
                m_leftoverParts = value;
            }
        }

        public string BlankColor
        {
            get
            {
                return m_blankColor;
            }

            set
            {
                m_blankColor = value;
            }
        }

        public string SheetSize
        {
            get
            {
                return m_sheetSize;
            }

            set
            {
                m_sheetSize = value;
            }
        }

        public string BlankComment
        {
            get
            {
                return m_blankComment;
            }

            set
            {
                m_blankComment = value;
            }
        }

        public string PalletSize
        {
            get
            {
                return m_palletSize;
            }

            set
            {
                m_palletSize = value;
            }
        }

        public int PalletQty
        {
            get
            {
                return m_palletQty;
            }

            set
            {
                m_palletQty = value;
            }
        }
    }
}
