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
using System.Drawing.Printing;
using Spire.Pdf;
using Spire.Pdf.Graphics;
using Order = Quick_Ship_Router.Order;
using Bill = Quick_Ship_Router.Bill;
using Item = Quick_Ship_Router.Item;
using QSTraveler = Quick_Ship_Router.Traveler;

namespace Traveler_Unraveler
{
    class TravelerUnraveler
    {
        // Interface
        public TravelerUnraveler(OdbcConnection mas,string project, ListView listView)
        {
            MAS = mas;
            m_projectName = project;
            m_listView = listView;
        }
        // Properties
        private System.Data.Odbc.OdbcConnection MAS;  // the ODBC driver connection to the MAS database
        private string m_projectName;
        private List<Order> m_orders =  new List<Order>();
        private List<Bill> m_models = new List<Bill>();
        private List<Traveler> travelers = new List<Traveler>();
        private ListView m_listView = new ListView();
        internal List<Traveler> Travelers
        {
            get
            {
                return travelers;
            }

            set
            {
                travelers = value;
            }
        }

        // Create and print all Travelers
        public void DisplayTravelers()
        {
            // display the results to the chairListView
            m_listView.Clear();
            // Set to details view.
            m_listView.View = View.Details;

            // production info
            m_listView.Columns.Add("Part No.", 150, HorizontalAlignment.Left);
            m_listView.Columns.Add("ID", 100, HorizontalAlignment.Left);
            m_listView.Columns.Add("Ordered", 100, HorizontalAlignment.Left);
            m_listView.Columns.Add("Need to Produce", 100, HorizontalAlignment.Left);
            // order info
            m_listView.Columns.Add("Order No.(s)", 200, HorizontalAlignment.Left);
            m_listView.Columns.Add("Customer(s)", 200, HorizontalAlignment.Left);
            m_listView.Columns.Add("Ship date(s)", 200, HorizontalAlignment.Left);


            foreach (Traveler traveler in travelers)
            {
                string dateList = "";
                string customerList = "";
                string orderList = "";
                int totalOrdered = 0;
                int i = 0;
               
                    totalOrdered += traveler.Order.QuantityOrdered;
                    dateList += (i == 0 ? "" : ", ") + traveler.Order.ShipDate.ToString("MM/dd/yyyy");
                    customerList += (i == 0 ? "" : ", ") + traveler.Order.CustomerNo;
                    orderList += (i == 0 ? "" : ", ") + traveler.Order.SalesOrderNo;
                    i++;
                string[] row = {
                    traveler.GetPart().BillNo,
                    traveler.ID.ToString("D6"),
                    totalOrdered.ToString(),
                    traveler.GetPart().TotalQuantity.ToString(),
                    orderList,
                    customerList,
                    dateList
                };
                ListViewItem tableListViewItem = new ListViewItem(row);
                tableListViewItem.Checked = true;
                m_listView.Items.Add(tableListViewItem);
            }
            m_listView.AutoResizeColumns(ColumnHeaderAutoResizeStyle.ColumnContent);
            m_listView.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize);
        }
        public void Print() {
            string exeDir = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            // open printed log file
            System.IO.StreamWriter file = File.AppendText(System.IO.Path.Combine(exeDir, "printed.json"));
            // Open excel
            var excelApp = new Excel.Application();
            excelApp.DisplayAlerts = false;
            // Open the traveler template
            var workbooks = excelApp.Workbooks;
            //@"\\Mgfs01\share\common\Traveler Unraveler\Production traveler.xlsx"
            var workbook = workbooks.Open(System.IO.Path.Combine(exeDir, "Production traveler.xlsx"),
                0, false, 5, "", "", false, 2, "",
                true, false, 0, true, false, false);
            var worksheets = workbook.Worksheets;
            var templateSheet = (Excel.Worksheet)worksheets.get_Item("Sheet1");

            // create the output workbook
            int currentSheet = 2;
            int index = 0;
            foreach (Traveler traveler in travelers)
            {
                templateSheet.Copy(Type.Missing, workbook.Worksheets[currentSheet - 1]);

                Excel.Worksheet outputSheet = workbook.Worksheets[currentSheet];
                // Part
                Excel.Range excelCell = (Excel.Range)outputSheet.get_Range("A2", "A2");
                excelCell.Value2 = traveler.GetPart().BillNo + "        " + traveler.GetPart().BillDesc;
                // Group
                excelCell = (Excel.Range)outputSheet.get_Range("C1", "C1");
                excelCell.Value2 = traveler.GetProject() + "-" + traveler.ID.ToString("D6");
                // Date
                excelCell = (Excel.Range)outputSheet.get_Range("D2", "D2");
                excelCell.Value2 = traveler.GetDate().ToString("MM-dd-yyyy");
                // Project
                excelCell = (Excel.Range)outputSheet.get_Range("K11", "K11");
                excelCell.Value2 = traveler.Order.SalesOrderNo;

                // Models
                int row = 3;
                foreach (Bill model in traveler.GetModels())
                {
                    // Model number
                    excelCell = (Excel.Range)outputSheet.get_Range("I" + row.ToString(), "I" + row.ToString());
                    excelCell.Value2 = model.BillNo;
                    // Quantity
                    excelCell = (Excel.Range)outputSheet.get_Range("L" + row.ToString(), "L" + row.ToString());
                    excelCell.Value2 = model.QuantityPerBill;
                    // Group
                    excelCell = (Excel.Range)outputSheet.get_Range("M" + row.ToString(), "M" + row.ToString());
                    excelCell.Value2 = traveler.GetProject();
                    row++;
                }
                // Total quantity
                excelCell = (Excel.Range)outputSheet.get_Range("L8", "L8");
                excelCell.Value2 = traveler.GetPart().QuantityPerBill;

                // Saw
                excelCell = (Excel.Range)outputSheet.get_Range("B4", "B4");
                excelCell.Value2 = (traveler.Saw != null ? traveler.Saw.BillNo : "N/A");
                excelCell = (Excel.Range)outputSheet.get_Range("C4", "C4");
                excelCell.Value2 = (traveler.Saw != null ? traveler.Saw.QuantityPerBill.ToString() : "");
                excelCell = (Excel.Range)outputSheet.get_Range("E4", "E4");
                excelCell.Value2 = (traveler.Saw != null ? traveler.Saw.Unit : "");
                // Edgebander
                excelCell = (Excel.Range)outputSheet.get_Range("B5", "B5");
                excelCell.Value2 = (traveler.GetEdgebander() != null ? traveler.GetEdgebander().BillNo : "N/A");
                excelCell = (Excel.Range)outputSheet.get_Range("C5", "C5");
                excelCell.Value2 = (traveler.GetEdgebander() != null ? traveler.GetEdgebander().QuantityPerBill.ToString() : "");
                excelCell = (Excel.Range)outputSheet.get_Range("E5", "E5");
                excelCell.Value2 = (traveler.GetEdgebander() != null ? traveler.GetEdgebander().Unit : "");
                // Contour
                if (traveler.ContourEdgebander)
                {
                    excelCell = (Excel.Range)outputSheet.get_Range("A5", "A5");
                    excelCell.Value2 = "Contour EB";
                }
                // Weeke/CNC
                excelCell = (Excel.Range)outputSheet.get_Range("B6", "B6");
                excelCell.Value2 = (traveler.GetWeeke() != null ? traveler.GetWeeke().BillNo : "N/A");
                excelCell = (Excel.Range)outputSheet.get_Range("C6", "C6");
                excelCell.Value2 = (traveler.GetWeeke() != null ? traveler.GetWeeke().QuantityPerBill.ToString() : "");
                excelCell = (Excel.Range)outputSheet.get_Range("E6", "E6");
                excelCell.Value2 = (traveler.GetWeeke() != null ? traveler.GetWeeke().Unit : "");
                // Misc
                excelCell = (Excel.Range)outputSheet.get_Range("B7", "B7");
                excelCell.Value2 = (traveler.MiscLabor != null ? traveler.MiscLabor.BillNo : "N/A");
                excelCell = (Excel.Range)outputSheet.get_Range("C7", "C7");
                excelCell.Value2 = (traveler.MiscLabor != null ? traveler.MiscLabor.QuantityPerBill.ToString() : "");
                excelCell = (Excel.Range)outputSheet.get_Range("E7", "E7");
                excelCell.Value2 = (traveler.MiscLabor != null ? traveler.MiscLabor.Unit : "");
                // Edgebanding
                excelCell = (Excel.Range)outputSheet.get_Range("B10", "B10");
                excelCell.Value2 = (traveler.GetEdgebanding() != null ? traveler.GetEdgebanding().ItemCode + "       " + traveler.GetEdgebanding().ItemCodeDesc : "N/A");
                excelCell = (Excel.Range)outputSheet.get_Range("C10", "C10");
                excelCell.Value2 = (traveler.GetEdgebanding() != null ? traveler.GetEdgebanding().QuantityPerBill.ToString() : "");
                excelCell = (Excel.Range)outputSheet.get_Range("E10", "E10");
                excelCell.Value2 = (traveler.GetEdgebanding() != null ? traveler.GetEdgebanding().Unit : "");
                // Material used
                if (traveler.GetMaterial() != null)
                {
                    excelCell = (Excel.Range)outputSheet.get_Range("B11", "B11");
                    excelCell.Value2 = traveler.GetMaterial().ItemCode + "        " + traveler.GetMaterial().ItemCodeDesc;
                    excelCell = (Excel.Range)outputSheet.get_Range("C11", "C11");
                    excelCell.Value2 = traveler.GetMaterial().QuantityPerBill;
                    excelCell = (Excel.Range)outputSheet.get_Range("E11", "E11");
                    excelCell.Value2 = traveler.GetMaterial().Unit;
                }
                //##### Print the Cover sheet #######
                outputSheet.PrintOut(
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                //###################################

                //####### Print the drawing #########

                //get rid of any appendages on the part number
                string drawingNo = "";
                foreach (char ch in traveler.GetPart().BillNo)
                {
                    if (ch == '-')
                    {
                        break;
                    }
                    else
                    {
                        drawingNo += ch;
                    }
                }

                SendToPrinter(@"\\Mgfs01\share\common\Drawings\Marco\PDF\" + drawingNo + ".pdf");

                // successfully printed, so we should log in the printed.json file
                if (!traveler.Printed)
                {
                    file.Write(traveler.Export());
                    file.Flush();
                }
                traveler.Printed = true;
                //###################################

                // clean up resources
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelCell);
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(outputSheet);
                
                currentSheet++;
                index++;
            }
            file.Close();
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(templateSheet);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(worksheets);
            workbook.Close();
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(workbook);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(workbooks);
            excelApp.Quit();
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelApp);
        }
        private void SendToPrinter(string path)
        {
            PdfDocument doc = new PdfDocument();
            try
            {
                doc.LoadFromFile(path);
                PrintDialog dialogPrint = new PrintDialog();
                dialogPrint.AllowPrintToFile = true;
                dialogPrint.AllowSomePages = true;
                dialogPrint.PrinterSettings.MinimumPage = 1;
                dialogPrint.PrinterSettings.MaximumPage = doc.Pages.Count;
                dialogPrint.PrinterSettings.FromPage = 1;
                dialogPrint.PrinterSettings.ToPage = doc.Pages.Count;

                if (dialogPrint.ShowDialog() == DialogResult.OK)
                {

                    //Set the pagenumber which you choose as the start page to print
                    doc.PrintFromPage = dialogPrint.PrinterSettings.FromPage;
                    //Set the pagenumber which you choose as the final page to print
                    doc.PrintToPage = dialogPrint.PrinterSettings.ToPage;
                    //Set the name of the printer which is to print the PDF
                    doc.PrinterName = dialogPrint.PrinterSettings.PrinterName;
                    //set the page size to be output automatically from the size of the PDF file:
                    //doc.PageScaling = PdfPrintPageScaling.ActualSize;
                    PaperSize paper = new PaperSize("Ledger", 1100, 1700);
                    paper.RawKind = (int)PaperKind.Tabloid;
                    dialogPrint.PrinterSettings.DefaultPageSettings.PaperSize = paper;
                    dialogPrint.PrinterSettings.Collate = false;
                    //dialogPrint.PrinterSettings.PrinterName;
                    //dialogPrint.PrinterSettings.PrinterName = "MX-5001N";

                    //doc.PrintDocument.DefaultPageSettings.PaperSize = paper;
                    PrintDocument printDoc = doc.PrintDocument;

                    //printDoc.DefaultPageSettings.PaperSize = paper;
                    printDoc.PrinterSettings = dialogPrint.PrinterSettings;
                    dialogPrint.Document = printDoc;

                    printDoc.Print();
                }
                
            }
            catch (Exception ex)
            {

            }


        }
        public void CreateTravelers()
        {
            travelers.Clear();
            ImportFromMAS();
            string exeDir = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            System.IO.StreamReader file = new System.IO.StreamReader(System.IO.Path.Combine(exeDir, "printed.json"));
            string line;
            List<Order> printedOrders = new List<Order>();
            while ((line = file.ReadLine()) != null && line != "")
            {
                QSTraveler printedTraveler = new QSTraveler(line);
                foreach (Order printedOrder in printedTraveler.Orders)
                {
                    foreach (Order currOrder in m_orders)
                    {
                        if (currOrder.SalesOrderNo == printedOrder.SalesOrderNo)
                        {
                            printedOrders.Add(currOrder);
                        }
                    }
                }
            }
            // remove the printed orders from current orders
            foreach (Order printedOrder in printedOrders)
            {
                m_orders.Remove(printedOrder);
            }
            file.Close();
            foreach (Order order in m_orders) {
                Bill model = new Bill(order.ItemCode, order.QuantityOrdered,MAS);
                // Make a unique traveler for each part, while combining common parts from different models into single travelers
                //foreach (Bill billItem in model.ComponentBills)
                //{
                Bill billItem = model;
                    bool foundBill = false;
                    // search for existing traveler
                    //foreach (Traveler traveler in travelers)
                    //{
                    //    if (billItem.BillNo == traveler.GetPart().BillNo)
                    //    {
                    //        foundBill = true;
                    //    // quantity of items to be produced
                    //    double quantity = model.QuantityPerBill;// * billItem.QuantityPerBill;
                    //        // add the model that the item comes from
                    //        Bill newModel = new Bill(model);
                    //        newModel.QuantityPerBill = quantity;
                    //        newModel.Group = group.ToString();
                    //        traveler.AddModel(newModel);
                    //        // add to the current part sum
                    //        traveler.GetPart().QuantityPerBill = traveler.GetPart().QuantityPerBill + quantity;
                    //        // change the group to "Common" -- the batch group
                    //        traveler.SetGroup("Common");
                    //    }
                    //}
                    // create a new traveler for the part
                    if (!foundBill)
                    {
                    // quantity of items to be produced
                    double quantity = model.QuantityPerBill;// * billItem.QuantityPerBill;
                        // create a new traveler from the new item
                        Bill travelerPart = new Bill(billItem);
                        travelerPart.QuantityPerBill = quantity;

                        Traveler newTraveler = new Traveler(m_projectName, order, travelerPart);
                        // add the model that the item comes from
                        Bill newModel = new Bill(model);
                        newModel.QuantityPerBill = quantity;
                        newTraveler.AddModel(newModel);
                        // add the new traveler to the list
                        travelers.Add(newTraveler);
                    }
                }
        }
        private void TallyBills(Bill bill, List<Bill> bills)
        {
            //tally up component bills
            foreach (Bill billItem in bill.ComponentBills)
            {
                bool foundBill = false;
                foreach (Bill tallyBill in bills)
                {
                    if (billItem.BillNo == tallyBill.BillNo)
                    {
                        foundBill = true;
                        // add to the existing tally item
                        tallyBill.QuantityPerBill = tallyBill.QuantityPerBill + bill.QuantityPerBill * billItem.QuantityPerBill;
                    }
                }
                // add a new item to the tally list
                if (!foundBill)
                {
                    bills.Add(new Bill(billItem.BillNo, bill.QuantityPerBill * billItem.QuantityPerBill, MAS));
                }
            }
        }
        private void TallyItems(Bill bill, List<Item> items)
        {
            // tally up component items
            foreach (Item billItem in bill.ComponentItems)
            {
                bool foundItem = false;
                foreach (Item tallyItem in items)
                {
                    if (billItem.ItemCode == tallyItem.ItemCode)
                    {
                        foundItem = true;
                        // add to the existing tally item
                        tallyItem.QuantityPerBill = tallyItem.QuantityPerBill + bill.QuantityPerBill * billItem.QuantityPerBill;
                    }
                }
                // add a new item to the tally list
                if (!foundItem)
                {
                    items.Add(new Item(billItem.ItemCode, bill.QuantityPerBill * billItem.QuantityPerBill, MAS));
                }
            }
            // tally up each bill
            foreach (Bill billComponent in bill.ComponentBills)
            {
                TallyItems(billComponent, items);
            }
        }
        // Updates travelers when project name is changed
        // Adds an item when clicked
        public void AddOrder(Order order)
        {
            m_orders.Add(order);
        }
        // Import from mas
        private void ImportFromMAS()
        {
            foreach (Bill bill in m_models)
            {
                bill.Import(MAS);
            }
        }
    }
}
