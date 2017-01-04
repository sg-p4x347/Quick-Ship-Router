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
    class Traveler
    {
        //===========================
        // PUBLIC
        //===========================

        // Doesn't do anything
        public Traveler()
        {

        }
        // Gets the base properties and orders of the traveler from a json string
        public Traveler(string json)
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
                        case ' ':
                        case '\t':
                        case '\n':
                            continue;
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
            m_printed = true;
        }
        // Creates a traveler from a part number and quantity
        public Traveler(string partNo, int quantity)
        {
            // set META information
            m_partNo = partNo;
            m_quantity = quantity;
            m_colorNo = m_partNo.Substring(m_partNo.Length - 2);
            // open the currentID.txt file
            string exeDir = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            System.IO.StreamReader readID = new StreamReader(System.IO.Path.Combine(exeDir, "currentID.txt"));
            m_ID = Convert.ToInt32(readID.ReadLine());
            readID.Close();
            // increment the current ID
            File.WriteAllText(System.IO.Path.Combine(exeDir, "currentID.txt"), (m_ID + 1).ToString() + '\n');
        }
        // Creates a traveler from a part number and quantity, then loads the bill of materials
        public Traveler(string partNo, int quantity, OdbcConnection MAS)
        {
            // set META information
            m_partNo = partNo;
            m_quantity = quantity;
            m_colorNo = m_partNo.Substring(m_partNo.Length - 2);
            ImportPart(MAS);
            // open the currentID.txt file
            string exeDir = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            System.IO.StreamReader readID = new StreamReader(System.IO.Path.Combine(exeDir, "currentID.txt"));
            m_ID = Convert.ToInt32(readID.ReadLine());
            readID.Close();
            // increment the current ID
            File.WriteAllText(System.IO.Path.Combine(exeDir, "currentID.txt"), (m_ID + 1).ToString() + '\n');
        }
        public void ImportPart(OdbcConnection MAS)
        {
            if (m_partNo != "")
            {
                m_part = new Bill(m_partNo, m_quantity, MAS);
                m_drawingNo = m_part.DrawingNo;
            }
        }
        public void FindWorkMaterial(Bill bill)
        {
            // find work and or material
            foreach (Item componentItem in bill.ComponentItems)
            {
                string itemCode = componentItem.ItemCode;
                if (itemCode == "/LWKE1" || itemCode == "/LWKE2" || itemCode == "/LCNC1" || itemCode == "/LCNC2")
                {
                    // CNC labor
                    m_cnc = componentItem;
                }
                else if (itemCode == "/LBND2" || itemCode == "/LBND3")
                {
                    // Straight Edgebander labor
                    m_ebander = componentItem;
                }
                else if (itemCode == "/LPNL1" || itemCode == "/LPNL2")
                {
                    // Panel Saw labor
                    m_saw = componentItem;
                }
                else if (itemCode == "/LCEB1" | itemCode == "/LCEB2")
                {
                    // Contour Edge Bander labor (vector)
                    m_vector = componentItem;
                }
                else if (itemCode == "/LATB1" || itemCode == "/LATB2" || itemCode == "/LATB3")
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
                    m_eband = componentItem;
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
                FindWorkMaterial(componentBill);
            }
        }
        //check inventory to see how many actually need to be produced.
        public void CheckInventory(OdbcConnection MAS)
        {
            try
            {
                OdbcCommand command = MAS.CreateCommand();
                command.CommandText = "SELECT QuantityOnSalesOrder, QuantityOnHand FROM IM_ItemWarehouse WHERE ItemCode = '" + m_part.BillNo + "'";
                OdbcDataReader reader = command.ExecuteReader();
                if (reader.Read())
                {
                    int available = Convert.ToInt32(reader.GetValue(1)) - Convert.ToInt32(reader.GetValue(0));
                    if (available >= 0)
                    {
                        // No parts need to be produced
                        m_quantity = 0;
                    }
                    else
                    {
                        // adjust the quantity that needs to be produced
                        m_quantity = Math.Min(-available, m_quantity);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occured when accessing inventory: " + ex.Message);
            }
        }
        // returns a JSON formatted string containing traveler information
        public string Export()
        {
            string doc = "";
            doc += "{";
            doc += "\"ID\":" + '"' + m_ID.ToString("D6") + '"' + ",";
            doc += "\"itemCode\":" + '"' + m_part.BillNo + '"' + ",";
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
        
        //===========================
        // PRIVATE
        //===========================

        // Properties
        protected List<Order> m_orders { get; set; } = new List<Order>();
        protected Bill m_part { get; set; } = null;
        protected int m_ID { get; set; } = 0;
        protected string m_timeStamp { get; set; } = "";
        protected bool m_printed { get; set; } = false;
        protected string m_partNo { get; set; } = "";
        protected string m_drawingNo { get; set; } = "";
        protected int m_quantity { get; set; } = 0;
        protected string m_colorNo { get; set; } = "";
        protected string m_color { get; set; } = "";
        // Labor
        protected Item m_cnc { get; set; } = null; // labor item
        protected Item m_vector { get; set; } = null; // labor item
        protected Item m_ebander { get; set; } = null; // labor item
        protected Item m_saw { get; set; } = null; // labor item
        protected Item m_assm { get; set; } = null; // labor item
        protected Item m_box { get; set; } = null; // labor item
        // Material
        protected Item m_material { get; set; } = null; // board material
        protected Item m_eband { get; set; } = null; // edgebanding
        protected List<Item> m_components { get; set; } = new List<Item>(); // metal components
        // Box
        protected string m_boxItemCode { get; set; } = "";
        protected string m_regPack { get; set; } = "N/A";
        protected int m_regPackQty { get; set; } = 0;
        protected string m_supPack { get; set; } = "N/A";
        protected int m_supPackQty { get; set; } = 0;
        
    }
}
