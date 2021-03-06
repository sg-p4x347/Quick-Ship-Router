﻿using System;
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
using System.Net;
using System.Net.Http;

namespace Quick_Ship_Router
{
    class TableManager
    {
        public TableManager() { }
        public TableManager(OdbcConnection mas, Label infoLabel,ProgressBar progressBar, ListView listview, bool checkPrinted = true)
        {
            m_MAS = mas;
            m_infoLabel = infoLabel;
            m_progressBar = progressBar;
            m_tableListView = listview;
            m_checkPrinted = checkPrinted;
        }
        //=======================
        // Travelers
        //=======================
        public void CompileTravelers(BackgroundWorker backgroundWorker1,Mode mode,string specificID,string fromS, string toS)
        {
            int from = 0;
            int to = 0;
            Int32.TryParse(fromS, out from);
            Int32.TryParse(toS, out to);
            string exeDir = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            // clear any previous travelers
            m_travelers.Clear();
            // get the list of travelers that have been printed
            List<Traveler> printedTravelers = new List<Traveler>();
            //==========================================
            // Remove any orders that have been printed
            //==========================================
            if (mode == Mode.CreateEATS)
            {
                string json = File.ReadAllText(specificID);
                //Dictionary<string, string> EATSimported = new EATS.StringStream(json).ParseJSON();
                List<string> travelers = new EATS.StringStream(json).ParseJSONarray();
                //List<string> orders = new EATS.StringStream(EATSimported["orders"]).ParseJSONarray();
                //List<EATS.Order> EATSorders = new List<EATS.Order>();
                EATS.OrderManager orderManager = new EATS.OrderManager();
                orderManager.ImportOrders(MAS);

                foreach (string travelerJSON in travelers)
                {
                    Dictionary<string, string> obj = new EATS.StringStream(travelerJSON).ParseJSON();
                    if (obj["type"] == "Table")
                    {
                        if (obj["state"] == "PreProcess" && obj["station"] != "Start")
                        {
                            Table table = new Table(obj["itemCode"], Convert.ToInt32(obj["quantity"]));
                            table.ID = Convert.ToInt32(obj["ID"]);
                            foreach (string orderNo in new EATS.StringStream(obj["parentOrders"]).ParseJSONarray())
                            {
                                try
                                {
                                    EATS.Order eatsOrder = orderManager.FindOrder(orderNo);

                                    Order order = new Order();
                                    order.SalesOrderNo = orderNo;
                                    
                                    order.ItemCode = obj["itemCode"];
                                    order.QuantityOrdered = eatsOrder.Items.Find(i => i.ItemCode == order.ItemCode).QtyOrdered;
                                    order.CustomerNo = eatsOrder.CustomerNo;
                                    order.ShipDate = eatsOrder.ShipDate;
                                    order.ShipVia = eatsOrder.ShipVia;
                                    order.OrderDate = eatsOrder.OrderDate;
                                    table.Orders.Add(order);
                                } catch (Exception ex)
                                {

                                }
                            }
                            Travelers.Add(table);
                        }
                    }
                }
            } else
            if (m_checkPrinted || mode == Mode.CreatePrinted || mode == Mode.DeletePrinted) {
                string line;
                System.IO.StreamReader file = new System.IO.StreamReader(System.IO.Path.Combine(exeDir, "printed.json"));
                List<Traveler> travelersToNotDelete = new List<Traveler>();
                int deletedQty = 0;
                while ((line = file.ReadLine()) != null && line != "")
                {
                    Traveler printedTraveler = new Traveler(line);
                    switch (mode)
                    {
                        case Mode.CreatePrinted:
                            // just add this traveler to the finished list (if within the range)
                            if (IsTable(printedTraveler.PartNo)) {
                                Table table = new Table(line);
                                if (table.ID >= from && table.ID <= to)
                                {
                                    Travelers.Add(table);
                                    foreach (Order order in table.Orders)
                                    {
                                        Order loadedOrder = m_orders.Find(o => o.SalesOrderNo == order.SalesOrderNo);
                                        if (loadedOrder != null)
                                        {
                                            order.ShipDate = loadedOrder.ShipDate;
                                            order.ShipVia = loadedOrder.ShipVia;
                                            order.OrderDate = loadedOrder.OrderDate;
                                            order.ProductLine = loadedOrder.ProductLine;
                                            order.CustomerNo = loadedOrder.CustomerNo;
                                        }
                                    }
                                }
                            }
                            break;
                        case Mode.DeletePrinted:
                            if (!(printedTraveler.ID >= from && printedTraveler.ID <= to)) {
                                if (IsTable(printedTraveler.PartNo))
                                {
                                    travelersToNotDelete.Add(new Table(line));
                                } else if (IsChair(printedTraveler.PartNo))
                                {
                                    travelersToNotDelete.Add(new Chair(line));
                                }
                                deletedQty++;
                            }
                            break;
                        case Mode.CreateSpecific:
                            if (printedTraveler.ID.ToString("D6") == specificID && IsTable(printedTraveler.PartNo))
                            {
                                Travelers.Add(new Table(line));
                                break;
                            }
                            goto default;
                        case Mode.CreateEATS:
                            // remove already printed travelers
                            Travelers.RemoveAll(t => t.ID == printedTraveler.ID);
                            break;
                        default:
                            // check to see if these orders have been printed already
                            foreach (Order printedOrder in printedTraveler.Orders)
                            {
                                foreach (Order order in m_orders)
                                {
                                    if (order.SalesOrderNo == printedOrder.SalesOrderNo && order.ItemCode == printedOrder.ItemCode)
                                    {
                                        // throw this order out
                                        if (mode != Mode.CreateSpecific)
                                        {
                                            m_orders.Remove(order);
                                            break;
                                        }
                                    }
                                }
                            }
                            break;
                    }
                }
                file.Close();

                // delete the travelers
                if (mode == Mode.DeletePrinted) {
                    File.Delete(System.IO.Path.Combine(exeDir, "printed.json"));
                    System.IO.StreamWriter newFile = File.AppendText(System.IO.Path.Combine(exeDir, "printed.json"));
                    foreach (Traveler traveler in travelersToNotDelete)
                    {
                        newFile.Write(traveler.Export());
                    }
                    newFile.Close();
                    //backgroundWorker1.ReportProgress( "Deleted " + deletedQty + " table travelers";
                }
                
            }
            if (mode != Mode.CreatePrinted && mode != Mode.CreateEATS && mode != Mode.DeletePrinted)
            {
                //==========================================
                // compile the travelers
                //==========================================
                
                int index = 0;
                foreach (Order order in m_orders)
                {
                    backgroundWorker1.ReportProgress(Convert.ToInt32((Convert.ToDouble(index) / Convert.ToDouble(m_orders.Count)) * 100), "Compiling Tables...");
                    // Make a unique traveler for each order, while combining common parts from different models into single traveler
                    bool foundBill = false;
                    // search for existing traveler
                    foreach (Table traveler in m_travelers)
                    {
                        if (traveler.Part == null) traveler.ImportPart(MAS);
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
                        Table newTraveler = new Table(order.ItemCode, order.QuantityOrdered, MAS);
                        // add to the order list
                        newTraveler.Orders.Add(order);
                        // add the new traveler to the list
                        m_travelers.Add(newTraveler);
                    }
                    index++;
                }
            }
            ImportInformation(backgroundWorker1,mode);
        }
        private Traveler FindTraveler(string s)
        {
            string exeDir = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            string line;
            System.IO.StreamReader file = new System.IO.StreamReader(System.IO.Path.Combine(exeDir, "printed.json"));
            int travelerID = 0;
            try
            {
                if (s.Length < 7)
                {
                    travelerID = Convert.ToInt32(s);
                }
            }
            catch (Exception ex)
            {

            }
            while ((line = file.ReadLine()) != null && line != "")
            {
                Traveler printedTraveler = new Traveler(line);
                // check to see if the number matches a traveler ID
                if (travelerID == printedTraveler.ID)
                {
                    return printedTraveler;
                }
                // check to see if these orders have been printed already
                foreach (Order printedOrder in printedTraveler.Orders)
                {
                    if (printedOrder.SalesOrderNo == s)
                    {
                        return printedTraveler;
                    }
                }
            }
            return null;
        }
        private void ImportInformation(BackgroundWorker backgroundWorker1,Mode mode = Mode.CreateAll)
        { 
            int index = 0;
            foreach (Table traveler in m_travelers)
            {
                if (traveler.Part == null) traveler.ImportPart(MAS);
                backgroundWorker1.ReportProgress(Convert.ToInt32((Convert.ToDouble(index) / Convert.ToDouble(m_travelers.Count)) * 100), "Gathering Table Info...");
                if (mode != Mode.CreateEATS) traveler.CheckInventory(MAS);
                
                // Table specific
                GetColorInfo(traveler);
                GetTableInfo(traveler); // sets the number of blanks & finds leftover parts
                int leftoverParts = traveler.LeftoverParts;
                traveler.Quantity += leftoverParts;
                GetTableInfo(traveler); // recalculate for the new total
                traveler.LeftoverParts = leftoverParts;
                // update and total the final parts
                traveler.Part.TotalQuantity = traveler.Quantity;
                traveler.FindComponents(traveler.Part);

                index++;
            }
        }
        // get a reader friendly string for the color
        private void GetColorInfo(Table traveler)
        {
            // open the color ref csv file
            string exeDir = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            System.IO.StreamReader colorRef = new StreamReader(System.IO.Path.Combine(exeDir, "Color Reference.csv"));
            colorRef.ReadLine(); // read past the header
            string line = colorRef.ReadLine();
            while (line != "")
            {
                string[] row = line.Split(',');
                if (Convert.ToInt32(row[0]) == traveler.ColorNo)
                {
                    traveler.Color = row[1];
                    traveler.BlankColor = row[2];
                    break;
                }
                line = colorRef.ReadLine();
            }
            colorRef.Close();
        }
        // calculate how much of each box size
        private void GetTableInfo(Table traveler)
        {
            // open a MAS connection
            OdbcCommand command = MAS.CreateCommand();
            command.CommandText = "SELECT UDF_TABLE_BLANK_NAME, UDF_TABLE_BLANK_SIZE, UDF_TABLE_SHAPE FROM CI_item WHERE itemCode = '" + traveler.PartNo + "'";
            OdbcDataReader reader = command.ExecuteReader();
            // read info
            if (reader.Read())
            {
                if (!reader.IsDBNull(0)) traveler.BlankNo = reader.GetString(0);
                if (!reader.IsDBNull(1)) traveler.BlankSize = reader.GetString(1);
                //if (!reader.IsDBNull(2)) traveler.ShapeNo = reader.GetString(2);
                if (traveler.BlankNo == "") traveler.BlankNo = "Missing blank info";
            }
                
            // open the table ref csv file
            string exeDir = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            System.IO.StreamReader tableRef = new StreamReader(System.IO.Path.Combine(exeDir, "Table Reference.csv"));
            tableRef.ReadLine(); // read past the header
            string line = tableRef.ReadLine();
            while (!tableRef.EndOfStream)
            {
                if (line == null)
                {
                    line = tableRef.ReadLine();
                    continue;
                }
                string[] row = line.Split(',');
                if (row[0] == traveler.ShapeNo)
                {
                    //--------------------------------------------
                    // BLANK INFO
                    //--------------------------------------------

                    //traveler.BlankSize = row[2];
                    traveler.SheetSize = row[3];
                    // [column 4 contains # of blanks per sheet]
                    traveler.PartsPerBlank = row[5] != "" ? Convert.ToInt32(row[5]) : 0;

                    // Exception cases -!-!-!-!-!-!-!-!-!-!-!-!-!-!-!-!-!-!-!-!-!-!-!-!-!-!-!-!-!-!-!-!-!-!-!
                    List<int> exceptionColors = new List<int> { 60, 50, 49 };
                    if ((traveler.ShapeNo == "MG2247" || traveler.ShapeNo == "38-2247") && exceptionColors.IndexOf(traveler.ColorNo) != -1)
                    {
                        // Exceptions to the blank parent sheet (certain colors have grain that can't be used with the typical blank)
                        traveler.BlankSize = "925x1532";
                        traveler.BlankComment = "Use " + traveler.SheetSize + " sheet and align grain";
                        traveler.PartsPerBlank = 2;
                    }
                    //!-!-!-!-!-!-!-!-!-!-!-!-!-!-!-!-!-!-!-!-!-!-!-!-!-!-!-!-!-!-!-!-!-!-!-!-!-!-!-!-!-!-!-!

                    // check to see if there is a MAGR blank
                    //if (traveler.BlankColor == "MAGR" && row[6] != "")
                    //{
                    //    traveler.BlankNo = row[6];
                    //}
                    //// check to see if there is a CHOK blank
                    //else if (traveler.BlankColor == "CHOK" && row[7] != "")
                    //{
                    //    traveler.BlankNo = row[7];
                    //}
                    //// there are is no specific blank size in the kanban
                    //else
                    //{
                    //    traveler.BlankNo = "";
                    //}
                    // calculate production numbers
                    if (traveler.PartsPerBlank <= 0) traveler.PartsPerBlank = 1;
                    decimal tablesPerBlank = Convert.ToDecimal(traveler.PartsPerBlank);
                    traveler.BlankQuantity = Convert.ToInt32(Math.Ceiling(Convert.ToDecimal(traveler.Quantity) / tablesPerBlank));
                    int partsProduced = traveler.BlankQuantity * Convert.ToInt32(tablesPerBlank);
                    traveler.LeftoverParts = partsProduced - traveler.Quantity;
                    //--------------------------------------------
                    // PACK & BOX INFO
                    //--------------------------------------------
                    traveler.SupPack = row[8];
                    traveler.SupPackQty = 0;
                    traveler.RegPack = row[9];
                    traveler.RegPackQty = 0;
                    traveler.PalletQty = 0;

                    int qtyOnHand = 0;
                    int qtyOrdered = 0;
                    foreach (Order order in traveler.Orders)
                    {
                        qtyOnHand += order.QuantityOnHand;
                        qtyOrdered += order.QuantityOrdered;
                        // Get box information
                        if (order.ShipVia.ToUpper() == "PPADD" || order.ShipVia.ToUpper() == "UNSP")
                        {
                            // don't make boxes for items in inventory
                            traveler.RegPackQty += order.QuantityOrdered - order.QuantityOnHand;
                            // approximately 20 max tables per pallet
                            traveler.PalletQty += Convert.ToInt32(Math.Ceiling(Convert.ToDouble(order.QuantityOrdered) / 20));
                        }
                        else
                        {
                            // don't make boxes for items in inventory (mostly super packed)
                            traveler.SupPackQty += order.QuantityOrdered - order.QuantityOnHand;
                        }
                    }
                    // add boxes for extra items
                    if (traveler.SupPackQty >= traveler.RegPackQty)
                    {
                        traveler.SupPackQty += Math.Max(0,traveler.Quantity - (qtyOrdered - qtyOnHand));
                    } else
                    {
                        traveler.RegPackQty += Math.Max(0,traveler.Quantity - (qtyOrdered - qtyOnHand));
                    }

                    //--------------------------------------------
                    // PALLET
                    //--------------------------------------------
                    traveler.PalletSize = row[11];

                    break;
                }
                line = tableRef.ReadLine();
            }
            tableRef.Close();
        }
        private bool IsTable(string s)
        {
            return (s.Length == 9 && s.Substring(0, 2) == "MG") || (s.Length == 10 && (s.Substring(0, 3) == "38-" || s.Substring(0, 3) == "41-"));
        }
        private bool IsChair(string s)
        {
            if (s.Length == 14 && s.Substring(0, 2) == "38")
            {
                string[] parts = s.Split('-');
                return (parts[0].Length == 5 && parts[1].Length == 4 && parts[2].Length == 3);
            }
            else if (s.Length == 15 && s.Substring(0, 4) == "MG11")
            {
                string[] parts = s.Split('-');
                return (parts[0].Length == 6 && parts[1].Length == 4 && parts[2].Length == 3);
            }
            else
            {
                return false;
            }

        }
        public void DisplayTravelers()
        {
            // display the results to the chairListView
            m_tableListView.Clear();
            // Set to details view.
            m_tableListView.View = View.Details;

            // production info
            m_tableListView.Columns.Add("Part No.", 150, HorizontalAlignment.Left);
            m_tableListView.Columns.Add("ID", 100, HorizontalAlignment.Left);
            m_tableListView.Columns.Add("Ordered", 100, HorizontalAlignment.Left);
            m_tableListView.Columns.Add("Need to Produce", 100, HorizontalAlignment.Left);
            // order info
            m_tableListView.Columns.Add("Order No.(s)", 200, HorizontalAlignment.Left);
            m_tableListView.Columns.Add("Customer(s)", 200, HorizontalAlignment.Left);
            m_tableListView.Columns.Add("Ship date(s)", 200, HorizontalAlignment.Left);


            foreach (Table traveler in m_travelers)
            {
                string dateList = "";
                string customerList = "";
                string orderList = "";
                int totalOrdered = 0;
                int i = 0;
                foreach (Order order in traveler.Orders)
                {
                    totalOrdered += order.QuantityOrdered;
                    dateList += (i == 0 ? "" : ", ") + order.ShipDate.ToString("MM/dd/yyyy");
                    customerList += (i == 0 ? "" : ", ") + order.CustomerNo;
                    orderList += (i == 0 ? "" : ", ") + "(" + order.QuantityOrdered + ") " + order.SalesOrderNo + (order.Comment != "" ? " [" + order.Comment + "]" : "");
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
                ListViewItem tableListViewItem = new ListViewItem(row);
                tableListViewItem.Checked = true;
                m_tableListView.Items.Add(tableListViewItem);
            }
            m_tableListView.AutoResizeColumns(ColumnHeaderAutoResizeStyle.ColumnContent);
            m_tableListView.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize);
        }
        //=======================
        // Printing
        //=======================
        public void PrintTravelers(Excel.Sheets worksheets)
        {
            m_infoLabel.Text = "Printing Table Travelers...";
            string exeDir = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            // open printed log file
            System.IO.StreamWriter file = File.AppendText(System.IO.Path.Combine(exeDir, "printed.json"));

            // create the output workbook
            for (int itemIndex = 0; itemIndex < m_tableListView.Items.Count; itemIndex++)
            {
                if (m_tableListView.Items[itemIndex].Checked)
                {
                    Table traveler = m_travelers[itemIndex];

                    // copy the sheet
                    worksheets.get_Item("Table").Copy(Type.Missing, worksheets[worksheets.Count]);
                    Excel.Worksheet outputSheet = worksheets[worksheets.Count];

                    // Sales Orders
                    string customerList = "";
                    string orderList = "";
                    int i = 0;

                    int inventoryQty = 0;
                    foreach (Order order in traveler.Orders)
                    {
                        inventoryQty += order.QuantityOrdered;
                    }
                    inventoryQty -= traveler.Quantity;
                    foreach (Order order in traveler.Orders)
                    {
                        // Uncomment this code if orders that are covered by inventory are not desired
                        //if (order.QuantityOrdered > order.QuantityOnHand)
                        //{
                            customerList += (i == 0 ? "" : ", ") + order.CustomerNo;
                            orderList += (i == 0 ? "" : ", ") + "(" + order.QuantityOrdered + ") " + order.SalesOrderNo + (order.Comment != "" ? " [" + order.Comment + "]" : "");
                        //}
                        i++;
                    }
                    //#####################
                    // Production Traveler
                    //#####################
                    Excel.Range range;
                    int row = 1;
                    // Documentation
                    range = outputSheet.get_Range("A" + row, "A" + row);
                    range.Value2 = traveler.ID.ToString("D6") + (traveler.Printed ? " COPY" : "");
                    range.get_Characters(7, 15).Font.FontStyle = "bold";
                    range.get_Characters(7, 15).Font.Size = 20;
                    row++;
                    // Part
                    range = outputSheet.get_Range("B" + row, "C" + row);
                    range.Item[1] = traveler.Part.BillNo;
                    range.Item[2] = traveler.Quantity;
                    row++;
                    // Description
                    range = outputSheet.get_Range("B" + row, "B" + row);
                    range.Value2 = traveler.Part.BillDesc;
                    row++;
                    // Drawing
                    range = outputSheet.get_Range("B" + row, "B" + row);
                    range.Value2 = traveler.Part.DrawingNo;
                    row++;
                    // Sales Orders
                    range = outputSheet.get_Range("B" + row, "C" + row);
                    range.Item[1].Value2 = orderList;
                    row++;
                    // Customers
                    range = outputSheet.get_Range("B" + row, "C" + row);
                    range.Item[1].Value2 = customerList;
                    row++;
                    // Date printed
                    traveler.TimeStamp = DateTime.Now.ToString("MM/dd/yyyy   hh:mm tt");
                    range = outputSheet.get_Range("B" + row, "B" + row);
                    range.Value2 = traveler.TimeStamp;
                    row++;
                    // Blank
                    range = outputSheet.get_Range("B" + row, "C" + row);
                    string blankInfo = "";
                    if (traveler.BlankNo != "")
                    {
                        blankInfo += traveler.BlankNo;
                    } else
                    {
                        blankInfo += traveler.BlankColor;
                    }
                    blankInfo += "   (" + traveler.BlankSize + ") [" + traveler.SheetSize + "]";
                    blankInfo += " " + traveler.PartsPerBlank + " per blank";
                    if (traveler.BlankComment != "") blankInfo += " " + traveler.BlankComment;
                    range.Item[1].Value2 = blankInfo;
                    range.Item[2].Value2 = Math.Max(1,traveler.BlankQuantity);
                    row++;
                    // Leftover
                    range = outputSheet.get_Range("B" + row, "C" + row);
                    range.Item[2].Value2 = traveler.LeftoverParts;
                    row++;
                    // Parent material
                    range = outputSheet.get_Range("B" + row, "C" + row);
                    range.Item[1].Value2 = traveler.Material.ItemCode;
                    range.Item[2].Value2 = traveler.Material.QuantityPerBill + " " + traveler.Material.Unit;
                    row++;
                    // Color
                    range = outputSheet.get_Range("B" + row, "B" + row);
                    range.Value2 = traveler.Color;
                    row++;
                    // Heien/Weeke rate
                    range = outputSheet.get_Range("B" + row, "C" + row);
                    range.Item[1].Value2 = traveler.Cnc.QuantityPerBill + " " + traveler.Cnc.Unit;
                    range.Item[2].Value2 = traveler.Cnc.QuantityPerBill * traveler.Quantity + " " + traveler.Vector.Unit;
                    row++;
                    // Vector rate
                    range = outputSheet.get_Range("B" + row, "C" + row);
                    range.Item[1].Value2 = traveler.Vector.QuantityPerBill + " " + traveler.Vector.Unit;
                    range.Item[2].Value2 = traveler.Vector.QuantityPerBill * traveler.Quantity + " " + traveler.Vector.Unit;
                    row++;
                    // Pack rate
                    if (traveler.Assm != null)
                    {
                        range = outputSheet.get_Range("B" + row, "C" + row);
                        range.Item[1].Value2 = traveler.Assm.QuantityPerBill + " " + traveler.Assm.Unit;
                        range.Item[2].Value2 = traveler.Assm.QuantityPerBill * (traveler.RegPackQty + traveler.SupPackQty) + " " + traveler.Vector.Unit;
                    }
                    row++;
                    // Regular pack
                    range = outputSheet.get_Range("B" + row, "C" + row);
                    range.Item[1].Value2 = traveler.RegPack + (traveler.BoxItemCode != "" ? " Or box: " + traveler.BoxItemCode : "");
                    range.Item[2].Value2 = traveler.RegPackQty;
                    row++;
                    // Super pack
                    range = outputSheet.get_Range("B" + row, "C" + row);
                    range.Item[1].Value2 = traveler.SupPack;
                    range.Item[2].Value2 = traveler.SupPackQty;
                    row++;
                    // Hardware components
                    range = outputSheet.get_Range("B" + row, "B" + row);
                    string hardware = "";
                    foreach (Item component in traveler.Components)
                    {
                        hardware += (hardware.Length > 0 ? "," : "") + "(" + component.TotalQuantity + ") " + component.ItemCode;
                    }
                    range.Value2 = hardware;
                    row++;
                    // Pallet
                    range = outputSheet.get_Range("B" + row, "C" + row);
                    range.Item[1].Value2 = traveler.PalletSize;
                    range.Item[2].Value2 = traveler.PalletQty;
                    row++;
                    
                    //#####################
                    // Box Construction
                    //#####################

                    row = 21;
                    // Documentation
                    range = outputSheet.get_Range("A" + row, "A" + row);
                    range.Value2 = traveler.ID.ToString("D6") + (traveler.Printed ? " COPY" : "");
                    range.get_Characters(7, 15).Font.FontStyle = "bold";
                    range.get_Characters(7, 15).Font.Size = 20;
                    row++;
                    // Part
                    range = outputSheet.get_Range("B" + row, "C" + row);
                    range.Item[1] = traveler.Part.BillNo;
                    range.Item[2] = traveler.SupPackQty + traveler.RegPackQty;
                    row++;
                    // Description
                    range = outputSheet.get_Range("B" + row, "B" + row);
                    range.Value2 = traveler.Part.BillDesc;
                    row++;
                    // Regular pack
                    range = outputSheet.get_Range("B" + row, "C" + row);
                    range.Item[1].Value2 = traveler.RegPack + (traveler.BoxItemCode != "" ? " Or box: " + traveler.BoxItemCode : "");
                    range.Item[2].Value2 = traveler.RegPackQty;
                    row++;
                    // Super pack
                    range = outputSheet.get_Range("B" + row, "C" + row);
                    range.Item[1].Value2 = traveler.SupPack;
                    range.Item[2].Value2 = traveler.SupPackQty;
                    row++;
                    // Box rate
                    if (traveler.Box != null)
                    {
                        range = outputSheet.get_Range("B" + row, "C" + row);
                        range.Item[1].Value2 = traveler.Box.QuantityPerBill + " " + traveler.Box.Unit;
                        range.Item[2].Value2 = traveler.Box.QuantityPerBill * traveler.Quantity + " " + traveler.Vector.Unit;
                    }
                    row++;
                    // Sales Orders
                    range = outputSheet.get_Range("B" + row, "C" + row);
                    range.Item[1].Value2 = orderList;
                    row++;
                    // Customers
                    range = outputSheet.get_Range("B" + row, "C" + row);
                    range.Item[1].Value2 = customerList;
                    row++;
                    // Date printed
                    range = outputSheet.get_Range("B" + row, "B" + row);
                    range.Value2 = traveler.TimeStamp;
                    row++;
                    try
                    {
                        //##### Print the Cover sheet #######
                        outputSheet.PrintOut(
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        //###################################

                        // successfully printed, so we should log in the printed.json file
                        if (!traveler.Printed)
                        {
                            file.Write(traveler.Export());
                            file.Flush();
                        }
                        traveler.Printed = true;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("A problem occured when printing table travelers: " + ex.Message);
                    }
                }
            }
            file.Close();
            m_infoLabel.Text = "";
        }
        async public void PrintLabels()
        {
            foreach (Table table in m_travelers) {
                string result = "";
                using (var client = new WebClient())
                {
                    client.Credentials = new NetworkCredential("gage", "Stargatep4x347");
                    client.Headers[HttpRequestHeader.ContentType] = "application/json";
                    string json = "{\"ID\":\"" + table.ID + "\",";
                    json += "\"Desc1\":\"" + table.Part.BillDesc + "\",";
                    json += "\"Desc2\":\"" + table.Eband + "\",";
                    json += "\"Pack\":\"" + (table.SupPackQty > 0 ? "SP" : "RP") + "\",";
                    json += "\"Date\":\"" + DateTime.Today.ToString(@"yyyy\-MM\-dd") + "\"}";
                    result = client.UploadString(@"http://crashridge.net:8080", "POST", json);
                    //http://192.168.2.6:8080/printLabel
                }
            }
        }
        //=======================
        // Properties
        //=======================
        private ListView m_tableListView = null;
        private Label m_infoLabel = null;
        private ProgressBar m_progressBar = null;
        private bool m_checkPrinted = true;
        private List<Order> m_orders = new List<Order>();
        private List<Table> m_travelers = new List<Table>();
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

        internal List<Table> Travelers
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
