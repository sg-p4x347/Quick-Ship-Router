using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using Marshal = System.Runtime.InteropServices.Marshal;
using System.Drawing.Printing;

namespace Quick_Ship_Router
{
    class Summary
    {
        class BlankItem
        {
            public BlankItem(string sz, int qty)
            {
                size = sz;
                quantity = qty;
            }
            public string size;
            public int quantity;
        }
        struct SummaryItem
        {
            public SummaryItem(string ID, string itmCode, int qty, string itemDesc)
            {
                travelerID = ID;
                itemCode = itmCode;
                quantityOrdered = qty;
                itemDescription = itemDesc;
                orderNums = new List<string>();
                orderDates = new List<string>();
                customers = new List<string>();
            }
            public string StringifyOrders ()
            {
                string list = "";
                foreach (string num in orderNums)
                {
                    list += (list.Length > 0 ? ", " : "") + num;
                }
                return list;
            }
            public string StringifyDates()
            {
                string list = "";
                foreach (string date in orderDates)
                {
                    list += (list.Length > 0 ? ", " : "") + date;
                }
                return list;
            }
            public string StringifyCustomers()
            {
                string list = "";
                foreach (string customer in customers)
                {
                    list += (list.Length > 0 ? ", " : "") + customer;
                }
                return list;
            }
            public string travelerID;
            public string itemCode;
            public int quantityOrdered;
            public string itemDescription;
            public List<string> orderNums;
            public List<string> orderDates;
            public List<string> customers;
        }
        public Summary(List<Order> orders, List<Router> routers)
        {
            // date
            date = DateTime.Today.ToString("MM/dd/yyyy");
            foreach (Router router in routers)
            {
                totalParts += router.Quantity;
                totalTravelers++;
                // create the summary item
                SummaryItem item = new SummaryItem(router.ID.ToString("D6"), router.Item.BillNo, router.Quantity, router.Item.BillDesc);
                foreach (Order order in router.Orders)
                {
                    item.orderNums.Add(order.SalesOrderNo);
                    item.orderDates.Add(order.OrderDate.ToString("MM/dd/yyyy"));
                    item.customers.Add(order.CustomerNo);
                }
                items.Add(item);
                // tally blanks
                bool foundBlank = false;
                for (int i = 0; i < blanks.Count; i++)
                {
                    if (blanks[i].size == router.BlankNo || blanks[i].size == router.BlankSize)
                    {
                        blanks[i].quantity += router.BlankQuantity;
                        foundBlank = true;
                    }
                }
                if (!foundBlank)
                {
                    blanks.Add(new BlankItem(router.BlankNo != "" ? router.BlankNo : router.BlankSize, router.BlankQuantity));
                }
                // total work
                totalCNC += router.Cnc.QuantityPerBill;
                totalVector += router.Vector.QuantityPerBill;
                totalPack += router.Assm.QuantityPerBill;
            }

            //####################
            // SORT BY ORDER
            //####################
            //// each traveler
            //foreach (Router router in routers)
            //{
            //    // add the orders
            //    foreach(Order order in router.Orders)
            //    {
            //        orders.Remove(order);
            //        items.Add(new SummaryItem(order.SalesOrderNo, order.QuantityOrdered, order.ItemCode, router.Item.BillDesc, router.ID.ToString("D6"), order.OrderDate.ToString("MM/dd/yyyy"), order.CustomerNo));
            //    }
            //    // tally blanks
            //    bool foundBlank = false;
            //    for(int i = 0; i < blanks.Count; i++)
            //    {
            //        if (blanks[i].size == router.BlankNo || blanks[i].size == router.BlankSize)
            //        {
            //            blanks[i].quantity += router.BlankQuantity;
            //            foundBlank = true;
            //        }
            //    }
            //    if (!foundBlank)
            //    {
            //        blanks.Add(new BlankItem(router.BlankNo != "" ? router.BlankNo : router.BlankSize, router.BlankQuantity));
            //    }
            //}
            //// orders that do not have travelers
            //foreach (Order order in orders)
            //{
            //    items.Add(new SummaryItem(order.SalesOrderNo, order.QuantityOrdered, order.ItemCode, "", "", order.OrderDate.ToString("MM/dd/yyyy"), order.CustomerNo));
            //}
        }
        public void Print(Excel.Workbooks workbooks)
        {
            // Open the Summary template for printing
            string exeDir = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            var workbook = workbooks.Open(System.IO.Path.Combine(exeDir, "Summary Sheet.xlsx"),
               0, false, 5, "", "", false, 2, "",
               true, false, 0, true, false, false);
            var worksheets = workbook.Worksheets;
            var summarySheet = (Excel.Worksheet)worksheets.get_Item("Summary Template");
            // date
            Excel.Range title = summarySheet.get_Range("A1", "A1");
            title.Value2 = "Open Order Traveler Summary for " + date;
            Marshal.ReleaseComObject(title);
            // totals
            Excel.Range travelerTotal = summarySheet.get_Range("A2", "A2");
            travelerTotal.Value2 = totalTravelers;
            Marshal.ReleaseComObject(travelerTotal);
            Excel.Range partTotal = summarySheet.get_Range("C2", "C2");
            partTotal.Value2 = totalParts;
            Marshal.ReleaseComObject(partTotal);
            // print items
            int row = 4;
            foreach (SummaryItem item in items)
            {
                Excel.Range range = summarySheet.get_Range("A" + row, "G" + row);
                range.Item[1].Value2 = item.travelerID;
                range.Item[2].Value2 = item.itemCode;
                range.Item[3].Value2 = item.quantityOrdered;
                range.Item[4].Value2 = item.itemDescription;
                range.Item[5].Value2 = item.StringifyOrders();
                range.Item[6].Value2 = item.StringifyDates();
                range.Item[7].Value2 = item.StringifyCustomers();
                // clean up range
                Marshal.ReleaseComObject(range);
                // increment row
                row++;
            }
            
            // print blank summary
            var totals = (Excel.Worksheet)worksheets.get_Item("Totals");
            row = 4;
            foreach (BlankItem blank in blanks)
            {
                Excel.Range range = totals.get_Range("A" + row, "B" + row);
                range.Item[1].Value2 = blank.size;
                range.Item[2].Value2 = blank.quantity;
                // clean up range
                Marshal.ReleaseComObject(range);
                // increment row
                row++;
            }
            // work totals
            Excel.Range cnc = totals.get_Range("E4","E4");
            cnc.Value2 = totalCNC;
            Marshal.ReleaseComObject(cnc);
            Excel.Range vector = totals.get_Range("E5", "E5");
            vector.Value2 = totalVector;
            Marshal.ReleaseComObject(vector);
            Excel.Range pack = totals.get_Range("E6", "E6");
            pack.Value2 = totalPack;
            Marshal.ReleaseComObject(pack);
            try
            {
                //##### Print Summary Sheet #######
                summarySheet.PrintOut(
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                //###################################
                //##### Print Blank Total Sheet #######
                totals.PrintOut(
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                //###################################
            }
            catch (Exception ex)
            {
                MessageBox.Show("A problem occured when printing the summary sheet: " + ex.Message);
            }
            // clean up excel objects
            Marshal.ReleaseComObject(summarySheet);
            Marshal.ReleaseComObject(totals);
            Marshal.ReleaseComObject(worksheets);
        }
        // Properties
        private string date = "";
        private List<SummaryItem> items = new List<SummaryItem>();
        private int totalParts = 0;
        private int totalTravelers = 0;
        private List<BlankItem> blanks = new List<BlankItem>();
        private double totalCNC = 0;
        private double totalVector = 0;
        private double totalPack = 0;
    }
}
