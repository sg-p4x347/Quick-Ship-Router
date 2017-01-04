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
    class Router
    {
        // Interface
        public Router(string json)
        {
            try
            {
                bool readString = false;
                string stringToken = "";

                string memberName = "";
                bool readMember = false;

                string value = "";
                bool readValue = false;

                // SalesOrderNo
                for (int pos = 0; pos < json.Length; pos++)
                {
                    char ch = json[pos];
                    switch (ch)
                    {
                        case '"':
                            readString = !readString;
                            continue;
                        case ':':
                            memberName = stringToken; stringToken = "";
                            continue;
                        case '[':
                            while (json[pos] != ']')
                            {
                                if (json[pos] == '{')
                                {
                                    string orderJson = "";
                                    while (json[pos] != '}')
                                    {
                                        ch = json[pos];
                                        orderJson += ch;
                                        pos++;
                                    }
                                    m_orders.Add(new Order(orderJson + '}'));
                                }
                                pos++;
                            }
                            continue;
                        case ',':
                            value = stringToken; stringToken = "";
                            // set the corresponding member
                            if (memberName == "ID")
                            {
                                m_ID = Convert.ToInt32(value);
                            }
                            else if (memberName == "copy")
                            {
                                m_copy = Convert.ToBoolean(value);
                            }
                            else if (memberName == "itemCode")
                            {
                                m_partNo = value;
                            }
                            else if (memberName == "quantity")
                            {
                                m_quantity = Convert.ToInt32(value);
                            }
                            continue;
                        case '}': continue;
                    }
                    if (readString)
                    {
                        // read string character by character
                        stringToken += ch;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Problem reading in traveler from printed.json: " + ex.Message);
            }
            m_copy = true;
        }
        public Router(string partNo, int quantity, string shipVia, OdbcConnection MAS, Excel.Worksheet crossRef, Excel.Worksheet colorRef, Excel.Worksheet boxRef)
        {
            string exeDir = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            System.IO.StreamReader readID = new StreamReader(System.IO.Path.Combine(exeDir, "currentID.txt"));
            ID = Convert.ToInt32(readID.ReadLine());
            readID.Close();
            // increment the current ID
            File.WriteAllText(System.IO.Path.Combine(exeDir, "currentID.txt"), (ID + 1).ToString() + '\n');
            // set META information
            m_partNo = partNo;
            m_quantity = quantity;
            // populate everything else
            GatherInfo(MAS, crossRef, colorRef, boxRef);
        }
        public void GatherInfo (OdbcConnection MAS, Excel.Worksheet crossRef, Excel.Worksheet colorRef, Excel.Worksheet boxRef)
        {
            // Complete the creation of the Traveler
            m_item = new Bill(m_partNo, m_quantity, MAS);
            m_drawingNo = m_item.DrawingNo;
            m_colorNo = m_partNo.Substring(m_partNo.Length - 2);
            // Get the color from the color reference
            for (int row = 2; row < 27; row++)
            {
                var colorRefRange = colorRef.get_Range("A" + row, "B" + row);
                if (Convert.ToString(colorRefRange.Item[1].Value2) == m_colorNo)
                {
                    m_color = colorRefRange.Item[2].Value2;
                }
                if (colorRefRange != null) Marshal.ReleaseComObject(colorRefRange);
            }

            for (int i = m_partNo.Length - 1; i > 0; i--)
            {
                if (m_partNo[i] == '-')
                {
                    m_shapeNo = m_partNo.Substring(0, i);
                    break;
                }
            }
            // fill out work and material
            FindWork(m_item);

            // Get information from the cross reference sheet
            for (int row = 2; row < 78; row++)
            {
                var range = (Excel.Range)crossRef.get_Range("B" + row.ToString(), "F" + row.ToString());
                // find the correct model number in the spreadsheet
                if (range.Item[1].Value2 == m_shapeNo)
                {
                    if (range.Item[3].Value2 == "Yes")
                    {
                        string blankCode = "";
                        // find the Blank code in the color table
                        for (int crow = 2; crow < 19; crow++)
                        {
                            var colorRange = crossRef.get_Range("K" + crow, "M" + crow);
                            // find the correct color
                            if (Convert.ToString(colorRange.Item[1].Value2) == m_colorNo)
                            {
                                EBand = colorRange.Item[2].Value2;
                                blankCode = colorRange.Item[3].Value2;
                                if (colorRange != null) Marshal.ReleaseComObject(colorRange);
                                break;
                            }
                            if (colorRange != null) Marshal.ReleaseComObject(colorRange);
                        }
                        // check to see if there is a MAGR blank
                        if (blankCode == "MAGR" && range.Item[4].Value2 != null)
                        {

                            m_blankNo = range.Item[4].Value2;
                        }
                        // check to see if there is a CHOK blank
                        else if (blankCode == "CHOK" && range.Item[5].Value2 != null)
                        {

                            m_blankNo = range.Item[5].Value2;
                        }
                        // there are no available blanks
                        else
                        {
                            m_blankNo = "";
                        }
                    }
                    if (range != null) Marshal.ReleaseComObject(range);
                    break;
                }
                if (range != null) Marshal.ReleaseComObject(range);
            }
        }
        private void FindTable(Bill bill)
        {
            if (bill.BillNo.Length >= 9 && bill.BillNo.Length <= 10 && (bill.BillNo.Substring(0,3) == "38-" || bill.BillNo.Substring(0, 3) == "41-" || bill.BillNo.Substring(0, 2) == "MG"))
            {
                // this is a table
                m_item = bill;
            } else
            {
                // keep looking
                foreach(Bill componentBill in bill.ComponentBills)
                {
                    FindTable(componentBill);
                }
            }
        }
        private void FindWork(Bill bill)
        {
            // find work and or material
            foreach (Item componentItem in bill.ComponentItems)
            {
                string itemCode = componentItem.ItemCode;
                if (itemCode == "/LWKE1" || itemCode == "/LWKE2" || itemCode == "/LCNC1" || itemCode == "/LCNC2")
                {
                    // WEEKE labor
                    m_cnc = componentItem;
                }
                else if (itemCode == "/LBND2" || itemCode == "/LBND3")
                {
                    // Straight Edgebander labor
                }
                else if (itemCode == "/LPNL1" || itemCode == "/LPNL2")
                {
                    // Panel Saw labor
                }
                else if (itemCode == "/LCEB1" | itemCode == "/LCEB2")
                {
                    // Contour Edge Bander labor (vector)
                    m_vector = componentItem;
                }
                else if (itemCode == "/LATB1" || itemCode == "/LATB2" || itemCode == "/LATB3" )
                {
                    // Assembly labor
                    m_assm = componentItem;
                }
                else if (itemCode == "/LBOX1")
                {
                    // Box construction labor
                    m_box = componentItem;
                }
                else if (itemCode.Substring(0, 3) == "006")
                {
                    // Material
                    m_material = componentItem;
                }
                else if (itemCode.Substring(0, 2) == "87")
                {
                    // Edgeband
                }
            }
            // Go deeper into each component bill
            foreach (Bill componentBill in bill.ComponentBills)
            {
                FindWork(componentBill);
            }
            // check to see if the material hasn't been set
            if (bill.ComponentBills.Count == 0 && m_material == null)
            {
                // find an item code to use as the material
                foreach (Item componentItem in bill.ComponentItems)
                {
                    // if the item code starts with a number, use it
                    if (Char.IsNumber(componentItem.ItemCode[0]))
                    {
                        m_material = componentItem;
                    }
                }
            }
        }
        public void FindHardware(Bill bill)
        {
            foreach (Item componentItem in bill.ComponentItems)
            {
                string itemCode = componentItem.ItemCode;
                // Screw || screw || leg bracket || table stretcher
                if (itemCode == "80434" || itemCode == "80435" || itemCode == "220102" || itemCode == "220114" || itemCode == "220115")
                {
                    // Append to hardware
                    m_hardware += (m_hardware.Length > 0 ? ",   " : "") + "(" + componentItem.QuantityPerBill * m_quantity + ") " + itemCode;
                }
                else if (m_box == null && itemCode.Substring(0, 2) == "90")
                {
                    // Paid for box
                    m_boxItemCode = itemCode;
                }
            }
            // Go deeper into each component bill
            foreach (Bill componentBill in bill.ComponentBills)
            {
                FindHardware(componentBill);
            }
        }
        // Stores the printed traveler permanently in a .txt file
        public string Export()
        {
            string doc = "";
            doc += "{";
            doc += "\"ID\":" + '"' + m_ID.ToString("D6") + '"' + ",";
            doc += "\"copy\":" + '"' + m_copy.ToString() + '"' + ",";
            doc += "\"itemCode\":" + '"' + m_item.BillNo + '"' + ",";
            doc += "\"quantity\":" + '"' + m_quantity + '"' + ",";
            doc += "\"orders\":[";
            foreach (Order order in m_orders)
            {
                doc += order.Export();
                doc += ",";
            }
            doc += "]";
            doc += "},\n";
            return doc;
        }
        // Properties
        private Bill    m_item = null;
        private int     m_ID = 0;
        private bool     m_copy = false;
        private string  m_timeStamp = "";
        private bool    m_printed = false;
        private string  m_partNo = "";
        private string  m_shapeNo = "";
        private string  m_colorNo = "";
        private string  m_drawingNo = "";
        private int     m_quantity = 0;
        private string  m_blankNo = "";
        private int     m_partsPerBlank = 0;
        private int     m_blankQuantity = 0;
        private int     m_leftoverParts = 0;
        private string  m_color = "";
        private string  m_blankSize = "";
        // Labor
        private Item    m_cnc = null; // labor item
        private Item    m_vector = null; // labor item
        private Item    m_assm = null; // labor item
        private Item    m_box = null; // labor item
        // Material
        private Item    m_material = null; // board material
        private string  m_sheetSize = "";
        private string  m_eBand = ""; // E-Band color
        private string  m_hardware = "";
        private string  m_boxItemCode = "";
        private string  m_regPack = "N/A";
        private int     m_regPackQty = 0;
        private string  m_supPack = "N/A";
        private int     m_supPackQty = 0;
        private List<Order> m_orders = new List<Order>();

        internal Bill Item
        {
            get
            {
                return m_item;
            }

            set
            {
                m_item = value;
            }
        }

        internal List<Order> Orders
        {
            get
            {
                return m_orders;
            }

            set
            {
                m_orders = value;
            }
        }

        internal Item Cnc
        {
            get
            {
                return m_cnc;
            }

            set
            {
                m_cnc = value;
            }
        }

        internal Item Vector
        {
            get
            {
                return m_vector;
            }

            set
            {
                m_vector = value;
            }
        }

        internal Item Material
        {
            get
            {
                return m_material;
            }

            set
            {
                m_material = value;
            }
        }

        internal string EBand
        {
            get
            {
                return m_eBand;
            }

            set
            {
                m_eBand = value;
            }
        }

        public string Hardware
        {
            get
            {
                return m_hardware;
            }

            set
            {
                m_hardware = value;
            }
        }

        public string RegPack
        {
            get
            {
                return m_regPack;
            }

            set
            {
                m_regPack = value;
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

        public string Color
        {
            get
            {
                return m_color;
            }

            set
            {
                m_color = value;
            }
        }

        public string ColorNo
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

        public int Quantity
        {
            get
            {
                return m_quantity;
            }

            set
            {
                m_quantity = value;
            }
        }

        public bool Printed
        {
            get
            {
                return m_printed;
            }

            set
            {
                m_printed = value;
            }
        }

        internal Item Assm
        {
            get
            {
                return m_assm;
            }

            set
            {
                m_assm = value;
            }
        }

        internal Item Box
        {
            get
            {
                return m_box;
            }

            set
            {
                m_box = value;
            }
        }

        public string SupPack
        {
            get
            {
                return m_supPack;
            }

            set
            {
                m_supPack = value;
            }
        }

        public int RegPackQty
        {
            get
            {
                return m_regPackQty;
            }

            set
            {
                m_regPackQty = value;
            }
        }

        public int SupPackQty
        {
            get
            {
                return m_supPackQty;
            }

            set
            {
                m_supPackQty = value;
            }
        }

        public string BoxItemCode
        {
            get
            {
                return m_boxItemCode;
            }

            set
            {
                m_boxItemCode = value;
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

        public int ID
        {
            get
            {
                return m_ID;
            }

            set
            {
                m_ID = value;
            }
        }

        public bool Copy
        {
            get
            {
                return m_copy;
            }

            set
            {
                m_copy = value;
            }
        }

        public string TimeStamp
        {
            get
            {
                return m_timeStamp;
            }

            set
            {
                m_timeStamp = value;
            }
        }
    }
}
