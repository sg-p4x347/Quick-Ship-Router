using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Quick_Ship_Router;
namespace Traveler_Unraveler
{
    class Traveler
    {
        // Interface
        public Traveler(string project, Order order, Bill part)
        {
            m_project = project;
            string exeDir = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            System.IO.StreamReader readID = new System.IO.StreamReader(System.IO.Path.Combine(exeDir, "currentID.txt"));
            m_ID = Convert.ToInt32(readID.ReadLine());
            readID.Close();
            m_order = order;
            m_part = part;


            // search each sub-part for work and or material
            FindWork(m_part);
            // exclude the Saw operation if the bill number is duplicated at any other stage
            if (m_saw != null)
            {
                string sawNo = m_saw.BillNo;
                if ((m_weeke != null && m_weeke.BillNo == sawNo) ||
                    (m_edgebander != null && m_edgebander.BillNo == sawNo) ||
                    (m_miscLabor != null && m_miscLabor.BillNo == sawNo))
                {
                    m_saw = null;
                }
            }
            // exclude edgeband if there is no edgebander work
            if (m_edgebanding != null && m_edgebander == null)
            {
                m_edgebanding = null;
            }
        }
        // Getters
        public string GetProject()
        {
            return m_project;
        }
        public Bill GetPart()
        {
            return m_part;
        }
        public List<Bill> GetModels()
        {
            return m_models;
        }
        public DateTime GetDate()
        {
            return m_date;
        }
        public Bill GetEdgebander()
        {
            return m_edgebander;
        }
        public Item GetMaterial()
        {
            return m_material;
        }
        public Bill GetWeeke()
        {
            return m_weeke;
        }
        public Item GetEdgebanding()
        {
            return m_edgebanding;
        }
        // Setters
        public void AddModel(Bill model)
        {
            m_models.Add(model);
        }
        // Helper
        private void FindWork(Bill bill)
        {
            // find work and or material
            foreach (Item componentItem in bill.ComponentItems)
            {
                string itemCode = componentItem.ItemCode;
                if (itemCode == "/LWKE1" || itemCode == "/LCNC1" || itemCode == "/LCNC2")
                {
                    // WEEKE labor
                    m_weeke = new Bill(bill);
                    m_weeke.QuantityPerBill = componentItem.QuantityPerBill;
                    m_weeke.Unit = componentItem.Unit;
                }
                else if (itemCode == "/LBND2" || itemCode == "/LBND3")
                {
                    // Straight Edgebander labor
                    m_edgebander = new Bill(bill);
                    m_edgebander.QuantityPerBill = componentItem.QuantityPerBill;
                    m_edgebander.Unit = componentItem.Unit;
                }
                else if (itemCode == "/LPNL1" || itemCode == "/LPNL2")
                {
                    // Panel Saw labor
                    m_saw = new Bill(bill);
                    m_saw.QuantityPerBill = componentItem.QuantityPerBill;
                    m_saw.Unit = componentItem.Unit;
                }
                else if (itemCode == "/LCEB1")
                {
                    // Contour Edge Bander labor
                    m_edgebander = new Bill(bill);
                    m_contourEdgebander = true;
                }
                else if (itemCode == "/LLAB1")
                {
                    // Misc Labor
                    m_miscLabor = new Bill(bill);
                    m_miscLabor.QuantityPerBill = componentItem.QuantityPerBill;
                    m_miscLabor.Unit = componentItem.Unit;
                }
                else if (itemCode.Substring(0,3) == "006")
                {
                    // Material
                    m_material = componentItem;
                }
                else if (itemCode.Substring(0,2) == "87")
                {
                    // Edgeband
                    m_edgebanding = componentItem;
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
                    if (Char.IsNumber(componentItem.ItemCode[0])) {
                        m_material = componentItem;
                    }
                }
            }
        }
        public string Export()
        {
            string doc = "";
            doc += "{";
            doc += "\"ID\":" + '"' + m_ID.ToString("D6") + '"' + ",";
            doc += "\"itemCode\":" + '"' + m_part.BillNo + '"' + ",";
            doc += "\"quantity\":" + '"' + m_part.TotalQuantity + '"' + ",";
            doc += "\"type\":" + '"' + this.GetType().Name + '"' + ",";
            doc += "\"orders\":[";
                doc += m_order.Export();
                doc += ",";
            doc += "]";
            doc += "},\n";
            return doc;
        }
        // Properties
        private string m_project;
        private int m_ID;
        private Order m_order;
        private DateTime m_date = DateTime.Now;
        private List<Bill> m_models = new List<Bill>();
        private Bill m_part;
        private bool m_printed = false;

        private Item m_material = null;
        private Bill m_edgebander = null;
        private bool m_contourEdgebander = false;
        private Item m_edgebanding = null;
        private Bill m_weeke = null;
        private Bill m_saw = null;
        private Bill m_miscLabor = null;

        internal Bill Saw
        {
            get
            {
                return m_saw;
            }

            set
            {
                m_saw = value;
            }
        }

        internal Bill MiscLabor
        {
            get
            {
                return m_miscLabor;
            }

            set
            {
                m_miscLabor = value;
            }
        }

        public bool ContourEdgebander
        {
            get
            {
                return m_contourEdgebander;
            }

            set
            {
                m_contourEdgebander = value;
            }
        }

        public Order Order
        {
            get
            {
                return m_order;
            }

            set
            {
                m_order = value;
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
    }
}
