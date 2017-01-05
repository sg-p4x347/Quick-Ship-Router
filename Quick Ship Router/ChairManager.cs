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
        public ChairManager() { }
        public ChairManager(OdbcConnection mas, Excel.Worksheet travelerTemplate, ListView listview) {
            m_MAS = mas;
            m_travelerTemplate = travelerTemplate;
            m_chairListView = listview;
        }
        //=======================
        // Travelers
        //=======================
        public void CompileTravelers()
        {
            string exeDir = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            // clear any previous travelers
            m_travelers.Clear();
            // get the list of routers that have been printed
            List<Traveler> printedTravelers = new List<Traveler>();
            //==========================================
            // Remove any orders that have been printed
            //==========================================
            string line;
            System.IO.StreamReader file = new System.IO.StreamReader(System.IO.Path.Combine(exeDir, "printed.json"));
            while ((line = file.ReadLine()) != null && line != "")
            {
                Traveler printedTraveler = new Traveler(line);
                foreach (Order printedOrder in printedTraveler.Orders)
                {
                    foreach (Order order in m_orders)
                    {
                        if (order.SalesOrderNo == printedOrder.SalesOrderNo)
                        {
                            // throw this order out
                            m_orders.Remove(order);
                            break;
                        }
                    }
                }
            }
            file.Close();
            //==========================================
            // compile the routers
            //==========================================
            foreach (Order order in m_orders)
            {
                // Make a unique router for each order, while combining common parts from different models into single router
                bool foundBill = false;
                // search for existing traveler
                foreach (Chair traveler in m_travelers)
                {
                    if (traveler.Part.BillNo == order.ItemCode)
                    {
                        // update existing traveler
                        foundBill = true;
                        // add to the quantity of items
                        traveler.Quantity += order.QuantityOrdered;
                        // add to the order list
                        traveler.Orders.Add(order);
                    }
                }
                if (!foundBill)
                {
                    // create a new traveler from the new item
                    Chair newTraveler = new Chair(order.ItemCode, order.QuantityOrdered, MAS);
                    // add to the order list
                    newTraveler.Orders.Add(order);
                    // add the new router to the list
                    m_travelers.Add(newTraveler);
                }
            }
            ImportInformation();
            DisplayTravelers();
        }
        private void ImportInformation()
        {
            foreach (Chair traveler in m_travelers)
            {
                
                traveler.CheckInventory(MAS);
                // update and total the final parts
                traveler.Part.TotalQuantity = traveler.Quantity;
                traveler.FindComponents(traveler.Part);
            }
        }
        public void DisplayTravelers()
        {
            // display the results to the chairListView
            m_chairListView.Clear();
            // Set to details view.
            m_chairListView.View = View.Details;

            // production info
            m_chairListView.Columns.Add("Part No.", 100, HorizontalAlignment.Left);
            m_chairListView.Columns.Add("ID", 50, HorizontalAlignment.Left);
            m_chairListView.Columns.Add("Ordered", 75, HorizontalAlignment.Left);
            m_chairListView.Columns.Add("Need to Produce", 75, HorizontalAlignment.Left);
            // order info
            m_chairListView.Columns.Add("Order No.(s)", 200, HorizontalAlignment.Left);
            m_chairListView.Columns.Add("Customer(s)", 200, HorizontalAlignment.Left);
            m_chairListView.Columns.Add("Ship date(s)", 100, HorizontalAlignment.Left);
            

            foreach (Chair traveler in m_travelers)
            {
                string dateList = "";
                string customerList = "";
                string orderList = "";
                int totalOrdered = 0;
                int i = 0;
                foreach (Order order in traveler.Orders)
                {
                    totalOrdered += order.QuantityOrdered;
                    dateList += (i == 0 ? "" : ", ") + order.OrderDate.ToString("MM/dd/yyyy");
                    customerList += (i == 0 ? "" : ", ") + order.CustomerNo;
                    orderList += (i == 0 ? "" : ", ") + order.SalesOrderNo;
                    i++;
                }
                string[] row = {
                    traveler.Part.BillNo,
                    traveler.ID.ToString("D6"),
                    totalOrdered.ToString(),
                    traveler.Quantity.ToString(),
                    orderList,
                    customerList,
                    dateList
                };
                ListViewItem chairListViewItem = new ListViewItem(row);
                chairListViewItem.Checked = true;
                m_chairListView.Items.Add(chairListViewItem);
            }
        }
        //=======================
        // Printing
        //=======================
        public void PrintTravelers()
        {

        }
        //=======================
        // Properties
        //=======================
        private Excel.Worksheet m_travelerTemplate;
        private ListView m_chairListView = null;
        private List<Order> m_orders = new List<Order>();
        private List<Chair> m_travelers = new List<Chair>();
        private OdbcConnection m_MAS = new OdbcConnection();

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

        internal List<Chair> Travelers
        {
            get
            {
                return m_travelers;
            }
        }

        internal OdbcConnection MAS
        {
            get
            {
                return m_MAS;
            }

            set
            {
                m_MAS = value;
            }
        }
    }
}
