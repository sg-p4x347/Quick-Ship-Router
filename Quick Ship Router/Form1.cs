// Program: Used to create traveling sheets of information that guide the production process
// Developer: Gage Coates
// Date started: 12/13/16

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
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            excelApp = new Excel.Application();
            excelApp.DisplayAlerts = false;
            // Open the traveler template
            workbooks = excelApp.Workbooks;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            PopulateCustomers();
            PopulateProductLines();
            // connect to MAS
            if (ConnectToData())
            {
                login.Enabled = false;
            }
        }
        // Interface
        ~Form1()
        {
            // close the MAS connection on exit
            MAS.Close();
            // close excel

            workbooks.Close();
            if (workbooks != null) Marshal.FinalReleaseComObject(workbooks);
            excelApp.Quit();
            if (excelApp != null) Marshal.FinalReleaseComObject(excelApp);
        }
        // Properties
        private Excel.Application excelApp;
        private Excel.Workbooks workbooks;
        private OdbcConnection MAS = new OdbcConnection();
        private List<Order> orders = new List<Order>();
        private List<Router> routers = new List<Router>();
        // Opens a connection to the MAS database
        private bool ConnectToData()
        {
            loadingLabel.Text = "Logging in...";
            MAS = new OdbcConnection();
            // initialize the MAS connection
            MAS.ConnectionString = "DSN=SOTAMAS90;Company=MGI;";
            try
            {
                MAS.Open();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to log in :(");
                return false;
            }
            loadingLabel.Text = "";
            return true;
        }
        // Fill the customer list with customers
        private void PopulateCustomers()
        {
            List<string> customers = new List<string>
            {
                "ABARGAS","ACEEDUC","ADP    ","AEROBLD","AFC    ","AGILE  ","AJAXSCH","ALABOUT","ALCON  ","ALLGLAS","ALTRA  ","ALUMBAU","AMAZOND","AMERICA","AMTAB  ","ANDERSN","ANIMAL ","AQADVOH","AQRSERV","AQUABUY","AQUADEP","AQUAIMP","AQUALND","AQUARAD","AQUARIA","AQUARIU","AQUATIC","ARVEST ","ATDAMER","ATDCAPL","BARNES ","BASSETT","BEAL   ","BECKER ","BEENEES","BEIMDIK","BIMART ","BLAZARC","BLUEFIN","BLUETHB","BMCOFF ","BOTSON ","BRALEY ","BRANCO ","BRNSNGR","C&HDIST","CAROLIN","CASCO  ","CENTRAL","CENTRPT","CFM    ","CHLDCAR","CHUCK  ","CLAIMS ","CLARK  ","CLAWPAW","COFACE ","COFCROC","COLEMAN","COMMERC","COMPACK","CONTFRN","COSTCO ","COSTCOW","COWPUBS","COYOTE ","CRAFTSH","CREATMK","CREWSMA","CROSBY ","CROSS  ","CROWDER","D&WSALE","DAISYBB","DALEY  ","DALMIDW","DAVCO  ","DAVENPO","DAVIDS ","DAVIDSB","DAVIDSO","DAVIS  ","DETWILE","DFWAQUA","DIABLO ","DONLOLL","DUBOSE ","DWYER  ","EASTSID","EBAY   ","EDUCDEP","ELEGAB ","ERICKSO","EVERFUR","EXFACTO","EXOTIC ","FACTSEL","FCOONEY","FELBAPT","FIECANA","FIESTAD","FINTAST","FISHPLA","FISHTAN","FISHWIS","FOREMAN","FOSTER ","FOURSOP","FRANKS ","FRYE   ","FURNNET","GASKET ","GENERAL","GODSCRE","GODSRES","GREATOU","GRIERIN","GRILLST","GUNLOCK","H&H    ","HARRIS ","HDHAPPY","HEMBREE","HEMBREM","HEMCO  ","HERTZFN","HOMEDEP","HOMETOW","HONCOMP","HOOVER ","HUNTE  ","HUNTER ","IFD    ","IFURN  ","INDOFF ","INTEGFN","IPAEDUC","IQSI   ","JACK'S ","JAMESCH","JAMESSH","JARDEN ","JJAMESO","JMJWORK","JSATRAD","K&S    ","KAMWOOD","KAY12  ","KENDAL ","KLOG   ","KSCONDS","KURTZBR","LAKESQ ","LANDMAN","LARSON ","LATTAS ","LAVACA ","LAZBOY ","LEGGETT","LIBERTY","LOISSCH","LONESTR","LOVELAN","LOWES  ","LOZIER ","MARTIN ","MARTLUT","MATEL  ","MAVIATN","MCALIST","MCCAULE","MCCOOL ","MEIJER ","MENARDS","MIDSTAT","MILLERS","MILLERZ","MILLS' ","MISC   ","MISCELL","MODOR  ","MOP    ","MORETHN","MOSER  ","MSSC   ","NATBOND","NATWIDE","NBFLA  ","NEEL   ","NEOSHOB","NEOSHOD","NEOSHR5","NETSHOP","NEWDISP","NEXTGEN","NOAHS  ","NOBIS  ","NOLANS ","OAKRIDG","OF003  ","OF008  ","OF011  ","OF012  ","OF013  ","OF014  ","OF023  ","OF031  ","OF032  ","OF034  ","OF035  ","OF046  ","OF059  ","OF065  ","OF067  ","OF070  ","OF071  ","OF083  ","OF088  ","OF090  ","OF091  ","OF093  ","OF094  ","OF099  ","OF109  ","OF111  ","OF112  ","OF113  ","OF114  ","OF120  ","OF123  ","OF124  ","OF141  ","OF150  ","OF153  ","OF156  ","OF158  ","OF167  ","OF169  ","OF172  ","OF174  ","OF180  ","OF181  ","OF182  ","OF184  ","OF185  ","OF190  ","OF197  ","OF205  ","OF208  ","OF209  ","OF211  ","OF226  ","OF232  ","OF233  ","OF237  ","OF243  ","OF250  ","OF252  ","OF255  ","OF256  ","OF268  ","OF270  ","OF272  ","OF273  ","OF279  ","OF283  ","OF284  ","OF289  ","OF291  ","OF293  ","OF295  ","OF298  ","OF299  ","OF314  ","OF317  ","OF321  ","OF323  ","OF326  ","OF327  ","OF329  ","OF330  ","OF335  ","OF336  ","OF338  ","OF341  ","OF348  ","OF352  ","OF353  ","OF354  ","OF360  ","OF362  ","OF368  ","OF371  ","OF372  ","OF374  ","OF377  ","OF378  ","OF384  ","OF385  ","OF386  ","OF387  ","OF391  ","OF392  ","OF393  ","OF394  ","OF395  ","OF396  ","OF397  ","OF398  ","OF399  ","OF401  ","OF402  ","OF403  ","OF404  ","OF405  ","OF406  ","OF407  ","OF408  ","OF409  ","OF410  ","OF411  ","OF413  ","OF422  ","OFCCON ","OFFDEP ","OFFDEPF","OFFDEPV","OFFDPBS","OFFMAX ","OFFSOUR","OFFSTAR","OFGPART","OFGWARR","OLDTOWN","OLPI   ","ONEWAYF","OSBORNE","OSULLIV","OVATION","OVERSTO","OZRKPLS","PALLETS","PARTS  ","PAYROLL","PEOPLES","PETCOCO","PETS PA","PETSMAR","PETSPLS","PETSUPE","PETZONE","PLAYTIM","PREWETT","QL     ","R&R MAC","R.G. AP","RACCHRC","REDNECK","REYNOLD","RJRAY  ","RMIND  ","ROEBLNG","ROGARDS","ROGERS ","ROSS   ","RTC    ","SAGECC ","SAMPLES","SAMS   ","SARTINF","SCHAEFE","SCHENKE","SCHLAID","SCHLBOX","SCHLPRD","SCHLSIN","SCLHSPR","SENECR7","SHERWIN","SHICK  ","SHORE  ","SIBLEY ","SMARKET","SPORTS ","SSIFURN","SSWORLD","STANDBY","STAOFMO","STAPLBI","STAPLES","STATELI","STEELMN","STRFRKD","STRONG ","SUNBEAM","TALBOT ","TANKSTO","TARPLEY","TEACHED","TEACHLG","TEENCHA","TEST   ","THOMPSN","THORCO ","TIENLE ","TOEWS  ","TOWNSEN","TPCINC ","TRAVIS ","TROPICL","TURNING","TWIN   ","TXSHCHS","UNBEAT ","UNITY  ","UPDATPT","UPS    ","USTOYCO","VASTMKT","VEASECM","VTINDUS","WALMART","WALMCOM","WAREHSE","WAYFAIR","WBMASON","WES MAT","WESTPOR","WILDSAL","WILPPET","WOODSPE","WORLDS ","WORLDWI","WORTHCN","WORTHDR","WOZEN  ","WYLIE  ","YELLOW ","ZERO   "
            };
            foreach (string customerNo in customers)
            {
                customerList.Items.Add(customerNo, (customerNo == "AMAZOND" || customerNo == "WAYFAIR"));
            }
        }
        private void PopulateProductLines()
        {
            productLineList.Items.Add("EDAT", true);
            productLineList.Items.Add("EDXD", true);
            productLineList.Items.Add("EDXT", true);
            productLineList.Items.Add("EDCT", true);
            productLineList.Items.Add("ECUS", false);
            productLineList.Items.Add("EDBS", false);
            productLineList.Items.Add("EDCM", false);
            productLineList.Items.Add("EDDK", false);
            productLineList.Items.Add("EDMC", false);
            productLineList.Items.Add("EDSE", false);
            productLineList.Items.Add("EDSS", false);
            productLineList.Items.Add("EDST", false);
            productLineList.Items.Add("EDXC", false);
            productLineList.Items.Add("EDXS", false);
            productLineList.Items.Add("EDXW", false);
            productLineList.Items.Add("EMAS", false);
        }
        private bool CreateTravelers(string specificOrderNo, bool invertCustomers)
        {
            // Import new orders
            if (ImportOrders(specificOrderNo,invertCustomers))
            {
                //orders = FilterByProduct(orders);
                // Create routers
                if (CreateRouters(false,true))
                {
                    loadingLabel.Text = "Travelers";
                    return true;
                }
            }
            return false;
        }
       
        private bool IsTable(string s)
        {
            return (s.Length == 9 && s.Substring(0, 2) == "MG") || (s.Length == 10 && (s.Substring(0, 3) == "38-" || s.Substring(0, 3) == "41-"));
        }
        private List<Order> FilterByProduct(List<Order> orders)
        {
            loadingLabel.Text = "Filtering...";
            string productLines = "";
            for (int i = 0; i < productLineList.Items.Count; i++)
            {
                if (productLineList.GetItemCheckState(i) == CheckState.Checked)
                {
                    productLines += (productLines.Length > 0 ? "," : "") + "'" + productLineList.GetItemText(productLineList.Items[i]) + "'";
                }
            }
            List<Order> filtered = new List<Order>();
            foreach (Order order in orders) {
                try
                {
                    OdbcCommand command = MAS.CreateCommand();
                    command.CommandText = "SELECT ProductLine FROM CI_item WHERE ItemCode = '" + order.ItemCode + "' AND ProductLine IN (" + productLines + ")";
                    OdbcDataReader reader = command.ExecuteReader(); 
                    // read info
                    if (reader.Read())
                    {
                        filtered.Add(order);
                    }
                } catch (Exception ex)
                {
                    MessageBox.Show("problem when filtering by product line: " + ex.Message);
                }
            }
            return filtered;
        }
        private bool CreateRouters(bool printed,bool checkPrinted)
        {
            // Open excel
            loadingLabel.Text = "Generating Travelers...";
            //@"\\Mgfs01\share\common\Traveler Unraveler\Production traveler.xlsx"
            string exeDir = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            var workbook = workbooks.Open(System.IO.Path.Combine(exeDir, "Kanban Blank Color Cross Reference.xlsx"),
                0, false, 5, "", "", false, 2, "",
                true, false, 0, true, false, false);
            //var workbook = workbooks.Open(@"\\Mgfs01\share\common\Quick Ship Traveler\Kanban Blank Color Cross Reference.xlsx",
            //    0, false, 5, "", "", false, 2, "",
            //    true, false, 0, true, false, false);
            var worksheets = workbook.Worksheets;
            var crossRef = (Excel.Worksheet)worksheets.get_Item("Blank Cross Reference");
            var colorRef = (Excel.Worksheet)worksheets.get_Item("Color Families");
            var blankRef = (Excel.Worksheet)worksheets.get_Item("Blank Parent");
            var boxRef = (Excel.Worksheet)worksheets.get_Item("Box Size");
            

            // get the list of routers that have been printed
            List<Router> printedRouters = new List<Router>();
            if (checkPrinted || printed)
            {
                string line;
                System.IO.StreamReader file = new System.IO.StreamReader(System.IO.Path.Combine(exeDir, "printed.json"));
                loadingLabel.Text = "Checking printed travelers...";
                while ((line = file.ReadLine()) != null && line != "")
                {
                    printedRouters.Add(new Router(line));
                }
                loadingLabel.Text = "";
                file.Close();
            }
            routers.Clear();
            if (printed)
            {
                foreach (Router router in printedRouters)
                {
                    router.GatherInfo(MAS, crossRef, colorRef, boxRef);
                }
                routers = printedRouters;
            }
            else
            {
                loadingLabel.Text = "Compiling travelers...";
                // compile the routers
                foreach (Order order in orders)
                {
                    if (checkPrinted)
                    {
                        // do not include this order if it has been printed
                        bool foundMatch = false;
                        foreach (Router printedRouter in printedRouters)
                        {
                            if (printedRouter.Orders.FindIndex(j => j.SalesOrderNo == order.SalesOrderNo) >= 0)
                            {
                                foundMatch = true;
                                break;
                            }
                        }
                        if (foundMatch) continue;
                    }
                    // Make a unique router for each order, while combining common parts from different models into single router
                    bool foundBill = false;
                    // search for existing traveler
                    if (combineOrders.Checked)
                    {
                        foreach (Router router in routers)
                        {
                            if (router.Item.BillNo == order.ItemCode)
                            {
                                foundBill = true;
                                router.Printed = printed;
                                // add to the quantity of items
                                router.Quantity += order.QuantityOrdered;
                                // add to the order list
                                router.Orders.Add(order);
                            }
                        }
                    }
                    // create a new traveler for the part
                    if (!foundBill)
                    {
                        // create a new traveler from the new item
                        Router newRouter = new Router(order.ItemCode, order.QuantityOrdered, order.ShipVia, MAS, crossRef, colorRef, boxRef);
                        newRouter.Printed = printed;
                        // add to the order list
                        newRouter.Orders.Add(order);
                        // add the new router to the list
                        routers.Add(newRouter);
                    }
                }
            }
            FinalizeRouters(crossRef, colorRef, blankRef, boxRef);
            // Clean up excel
            if (crossRef != null) Marshal.ReleaseComObject(crossRef);
            if (colorRef != null) Marshal.ReleaseComObject(colorRef);
            if (blankRef != null) Marshal.ReleaseComObject(blankRef);
            if (boxRef != null) Marshal.ReleaseComObject(boxRef);
            if (worksheets != null) Marshal.ReleaseComObject(worksheets);
            workbook.Close(false);
            if (workbook != null) Marshal.ReleaseComObject(workbook);

            DisplayRouters();
            return true;
        }
        private void FinalizeRouters(Excel.Worksheet crossRef, Excel.Worksheet colorRef, Excel.Worksheet blankRef, Excel.Worksheet boxRef)
        {
            // Loop through the routers to total up the final quantities of materials needed
            List<Router> remove = new List<Router>();
            foreach (Router router in routers)
            {
                //---------------------------------------------------------------
                // check inventory to see how many actually need to be produced.
                //---------------------------------------------------------------
                try
                {
                    OdbcCommand command = MAS.CreateCommand();
                    command.CommandText = "SELECT QuantityOnSalesOrder, QuantityOnHand FROM IM_ItemWarehouse WHERE ItemCode = '" + router.Item.BillNo + "'";
                    OdbcDataReader reader = command.ExecuteReader();
                    if (reader.Read())
                    {
                        int available = Convert.ToInt32(reader.GetValue(1)) - Convert.ToInt32(reader.GetValue(0));
                        if (available >= 0)
                        {
                            // remove this router, there are parts already in inventory
                            router.Quantity = 0;
                            remove.Add(router);
                        }
                        else
                        {
                            // adjust the quantity that need to be produced
                            router.Quantity = Math.Min(-available, router.Quantity);
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("An error occured when accessing inventory: " + ex.Message);
                }
                // calculate total hardware
                router.FindHardware(router.Item);
                // calculate total material sqft
                router.Material.QuantityPerBill *= router.Quantity;
                //---------------------------------------------------------------
                // calculate how much of each box size
                //---------------------------------------------------------------
                for (int row = 2; row < 78; row++)
                {

                    var range = crossRef.get_Range("B" + row.ToString(), "F" + row.ToString());
                    // find the correct model number in the spreadsheet
                    if (range.Item[1].Value2 == router.ShapeNo)
                    {
                        foreach (Order order in router.Orders)
                        {
                            // Get box information
                            if (order.ShipVia != "" && (order.ShipVia.ToUpper().IndexOf("FEDEX") != -1 || order.ShipVia.ToUpper().IndexOf("UPS") != -1))
                            {
                                var boxRange = boxRef.get_Range("C" + (row + 1), "H" + (row + 1)); // Super Pack
                                router.SupPack = (boxRange.Item[1].Value2 != null ? boxRange.Item[5].Value2 + " ( " + boxRange.Item[1].Value2 + " x " + boxRange.Item[2].Value2 + " x " + boxRange.Item[3].Value2 + " )" + (boxRange.Item[4].Value2 != null ? boxRange.Item[4].Value2 + " pads" : "") : "Missing information") + (boxRange.Item[6].Value2 != null ? " " + boxRange.Item[6].Value2 : "");
                                router.SupPackQty += order.QuantityOrdered;
                                if (boxRange != null) Marshal.ReleaseComObject(boxRange);
                            }
                            else
                            {
                                var boxRange = boxRef.get_Range("I" + (row + 1), "N" + (row + 1)); // Regular Pack
                                router.RegPack = (boxRange.Item[1].Value2 != null ? boxRange.Item[5].Value2 + " ( " + boxRange.Item[1].Value2 + " x " + boxRange.Item[2].Value2 + " x " + boxRange.Item[3].Value2 + " )" : "Missing information") + (boxRange.Item[6].Value2 != null ? " " + boxRange.Item[6].Value2 : "");
                                router.RegPackQty += order.QuantityOrdered;
                                if (boxRange != null) Marshal.ReleaseComObject(boxRange);
                            }
                        }
                    }
                    if (range != null) Marshal.ReleaseComObject(range);
                }
                //--------------------------------------
                // Calculate how many will be left over + Blank Size
                //--------------------------------------
                for (int row = 2; row < 78; row++)
                {
                    var blankRange = blankRef.get_Range("A" + row.ToString(), "H" + row.ToString());
                    // find the correct model number in the spreadsheet
                    if (blankRange.Item[1].Value2 == router.ShapeNo)
                    {
                        // set the blank size
                        List<string> exceptionColors = new List<string> { "60", "50", "49" };
                        if ((router.ShapeNo == "MG2247" || router.ShapeNo == "38-2247") && exceptionColors.IndexOf(router.ColorNo) != -1)
                        {
                            // Exceptions to the blank parent sheet (certain colors have grain that can't be used with the typical blank)
                            router.BlankSize = "(920x1532)";
                            router.PartsPerBlank = 1;
                        }
                        else
                        {
                            // All normal
                            if (Convert.ToInt32(blankRange.Item[7].Value2) > 0)
                            {
                                router.BlankSize = "(" + blankRange.Item[8].Value2 + ")";
                                router.PartsPerBlank = Convert.ToInt32(blankRange.Item[7].Value2);
                            } else
                            {
                                if (blankRange.Item[5].Value2 != "-99999")
                                {
                                    router.BlankSize = "(" + blankRange.Item[5].Value2 + ") ~sheet";
                                } else
                                {
                                    router.BlankSize = "No Blank";
                                }
                            }
                            
                        }
                        // calculate production numbers
                        if (router.PartsPerBlank < 0) router.PartsPerBlank = 0;
                        decimal tablesPerBlank = Convert.ToDecimal(blankRange.Item[7].Value2);
                        if (tablesPerBlank <= 0) tablesPerBlank = 1;
                        router.BlankQuantity = Convert.ToInt32(Math.Ceiling(Convert.ToDecimal(router.Quantity) / tablesPerBlank));
                        int partsProduced = router.BlankQuantity * Convert.ToInt32(tablesPerBlank);
                        router.LeftoverParts = partsProduced - router.Quantity;
                        break;
                    }
                    if (blankRange != null) Marshal.ReleaseComObject(blankRange);
                }
                // subtract the inventory parts from the box quantity
                // router.RegPackQty = Math.Max(0, router.RegPackQty - ((router.RegPackQty + router.SupPackQty) - router.Quantity));
            }
            // remove the removed routers
            foreach (Router router in remove)
            {
                // check to see if there are super-packed items, if so, don't remove this traveler
                if (router.SupPackQty == 0)
                {
                    routers.Remove(router);
                }
            }
        }
        private void DisplayRouters()
        {
            loadingLabel.Text = "Travelers";
            // display the results to the tableListView
            tableListView.Clear();
            // Set to details view.
            tableListView.View = View.Details;

            // production info
            tableListView.Columns.Add("Part No.", 100, HorizontalAlignment.Left);
            tableListView.Columns.Add("ID", 50, HorizontalAlignment.Left);
            tableListView.Columns.Add("Printed", 50, HorizontalAlignment.Left);
            tableListView.Columns.Add("Ordered", 75, HorizontalAlignment.Left);
            tableListView.Columns.Add("Need to Produce", 75, HorizontalAlignment.Left);
            tableListView.Columns.Add("Blanks", 75, HorizontalAlignment.Left);
            tableListView.Columns.Add("Leftover", 75, HorizontalAlignment.Left);
            // order info
            tableListView.Columns.Add("Ship date(s)", 100, HorizontalAlignment.Left);
            tableListView.Columns.Add("Customer(s)", 200, HorizontalAlignment.Left);
            tableListView.Columns.Add("Order No.(s)", 200, HorizontalAlignment.Left);
            // Pack info
            tableListView.Columns.Add("Reg pack", 100, HorizontalAlignment.Left);
            tableListView.Columns.Add("Sup pack", 100, HorizontalAlignment.Left);
            // Traveler info
            tableListView.Columns.Add("Drawing No.", 100, HorizontalAlignment.Left);
            tableListView.Columns.Add("Blank No.", 100, HorizontalAlignment.Left);
            tableListView.Columns.Add("Blank Size", 100, HorizontalAlignment.Left);
            tableListView.Columns.Add("Heian/Weeke Labor", 100, HorizontalAlignment.Left);
            tableListView.Columns.Add("Vector Labor", 100, HorizontalAlignment.Left);
            tableListView.Columns.Add("Color", 200, HorizontalAlignment.Left);
            tableListView.Columns.Add("Hardware", 200, HorizontalAlignment.Left);
            
            
            foreach (Router router in routers)
            {
                string dateList = "";
                string customerList = "";
                string orderList = "";
                int totalOrdered = 0;
                int i = 0;
                foreach (Order order in router.Orders)
                {
                    totalOrdered += order.QuantityOrdered;
                    dateList += (i == 0 ? "" : ", ") + order.OrderDate.ToString("MM/dd/yyyy");
                    customerList += (i == 0 ? "" : ", ") + order.CustomerNo;
                    orderList += (i == 0 ? "" : ", ") + order.SalesOrderNo;
                    i++;
                }
                string[] row = {
                    router.Item.BillNo,router.ID.ToString("D6"),
                    (router.Printed ? "Yes" : "No"),
                    totalOrdered.ToString(),
                    router.Quantity.ToString(),
                    router.BlankQuantity.ToString(),
                    router.LeftoverParts.ToString(),
                    dateList,
                    customerList,
                    orderList,
                    router.RegPackQty.ToString(),
                    router.SupPackQty.ToString(),
                    router.Item.DrawingNo,
                    router.BlankNo,
                    router.BlankSize,
                    router.Cnc.QuantityPerBill.ToString() + " " + router.Cnc.Unit,
                    router.Vector.QuantityPerBill.ToString() + " " + router.Vector.Unit,
                    router.Color,router.Hardware,
                };
                ListViewItem tableListViewItem = new ListViewItem(row);
                tableListViewItem.Checked = true;
                tableListView.Items.Add(tableListViewItem);
            }
        }
        //======================
        // Open Order
        //======================
        private bool ImportOrders(string specificOrderNo, bool invertCustomers)
        {
            List<Order> tableOrders = new List<Order>();
            List<Order> chairOrders = new List<Order>();
            loadingLabel.Text = "Importing Orders...";
            // only leave the previous orders if we are adding one by one
            if (specificOrderNo == "")
            {
                orders.Clear();
            }
            string today = DateTime.Today.ToString(@"yyyy\-MM\-dd");
            // OrderDate >= {d '" +  todayString + "'}
            string customerNames = "";
            for (int i = 0; i < customerList.Items.Count; i++)
            {
                if (customerList.GetItemCheckState(i) == CheckState.Checked)
                {
                    customerNames += (customerNames.Length > 0 ? "," : "") + "'" + customerList.GetItemText(customerList.Items[i]) + "'";
                }
            }
            // get informatino from header
            OdbcCommand command = MAS.CreateCommand();
            command.CommandText = "SELECT SalesOrderNo, ShipExpireDate, CustomerNo, ShipVia FROM SO_SalesOrderHeader WHERE " + (specificOrderNo != "" ? "SalesOrderNo = '" + specificOrderNo + "'" : "CustomerNo " + (invertCustomers ? "NOT" : "") + " IN (" + customerNames + ")" + (showToday.Checked ? "AND OrderDate >= {d '" + today + "'}" : ""));
            OdbcDataReader reader = command.ExecuteReader();
            // read info
            while (reader.Read())
            {
                // get information from detail
                OdbcCommand detailCommand = MAS.CreateCommand();
                detailCommand.CommandText = "SELECT ItemCode, QuantityOrdered, UnitOfMeasure FROM SO_SalesOrderDetail WHERE SalesOrderNo = '" + reader.GetString(0) + "'";
                OdbcDataReader detailReader = detailCommand.ExecuteReader();
                // Read each line of the Sales Order, looking for the base Table items, ignoring kits
                while (detailReader.Read())
                {
                    // Import bill & quantity 
                    string billCode = detailReader.GetString(0);
                    if (!detailReader.IsDBNull(2) && detailReader.GetString(2) != "KIT" && IsTable(billCode))
                    {
                        Order order = new Order();
                        // scrap this order if anything is missing
                        if (reader.IsDBNull(0)) continue;
                        order.SalesOrderNo = reader.GetString(0);
                        if (reader.IsDBNull(1)) continue;
                        order.OrderDate = reader.GetDate(1);
                        if (reader.IsDBNull(2)) continue;
                        order.CustomerNo = reader.GetString(2);
                        if (reader.IsDBNull(3)) continue;
                        order.ShipVia = reader.GetString(3);
                        // this is a table
                        order.ItemCode = billCode;
                        order.QuantityOrdered = Convert.ToInt32(detailReader.GetValue(1));
                        orders.Add(order);
                        continue;
                    }
                }
            }
            loadingLabel.Text = "";
            return true;
        }
        //======================
        // Chairs
        //======================
        private ChairManager chairManager = new ChairManager();
        //======================
        // Events
        //======================
        private void btnPrint_Click(object sender, EventArgs e)
        {
            loadingLabel.Text = "Printing Travelers...";
            string exeDir = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            //var workbook = workbooks.Open(System.IO.Path.Combine(exeDir, "traveler.xlsx"),
            //    0, false, 5, "", "", false, 2, "",
            //    true, false, 0, true, false, false);
            var workbook = workbooks.Open(System.IO.Path.Combine(exeDir, "traveler.xlsx"),
                0, false, 5, "", "", false, 2, "",
                true, false, 0, true, false, false);
            var worksheets = workbook.Worksheets;
            var templateSheet = (Excel.Worksheet)worksheets.get_Item("Table");
            // open printed log file
            System.IO.StreamWriter file = File.AppendText(System.IO.Path.Combine(exeDir, "printed.json"));
            
            
            // create the output workbook
            int currentSheet = 2;
            for (int itemIndex = 0; itemIndex < tableListView.Items.Count; itemIndex++)
            {
                if (tableListView.Items[itemIndex].Checked)
                {
                    Router router = routers[itemIndex];
                    templateSheet.Copy(Type.Missing, workbook.Worksheets[currentSheet - 1]);

                    Excel.Worksheet outputSheet = workbook.Worksheets[currentSheet];

                    // Sales Orders
                    string customerList = "";
                    string orderList = "";
                    int i = 0;
                    foreach (Order order in router.Orders)
                    {
                        customerList += (i == 0 ? "" : ", ") + order.CustomerNo;
                        orderList += (i == 0 ? "" : ", ") + "(" + order.QuantityOrdered + ") " + order.SalesOrderNo;
                        i++;
                    }
                    //#####################
                    // Production Traveler
                    //#####################
                    Excel.Range range;
                    int row = 1;
                    // Documentation
                    range = outputSheet.get_Range("A" + row, "A" + row);
                    range.Value2 = router.ID.ToString("D6") + (router.Copy ? " COPY" : "");
                    range.get_Characters(7, 15).Font.FontStyle = "bold";
                    range.get_Characters(7, 15).Font.Size = 20;
                    row++;
                    // Part
                    range = outputSheet.get_Range("B" + row, "C" + row);
                    range.Item[1] = router.Item.BillNo;
                    range.Item[2] = router.Quantity;
                    row++;
                    // Description
                    range = outputSheet.get_Range("B" + row, "B" + row);
                    range.Value2 = router.Item.BillDesc;
                    row++;
                    // Drawing
                    range = outputSheet.get_Range("B" + row, "B" + row);
                    range.Value2 = router.Item.DrawingNo;
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
                    router.TimeStamp = DateTime.Now.ToString("MM/dd/yyyy   hh:mm tt");
                    range = outputSheet.get_Range("B" + row, "B" + row);
                    range.Value2 = router.TimeStamp;
                    row++;
                    // Blank
                    range = outputSheet.get_Range("B" + row, "C" + row);
                    range.Item[1].Value2 = router.BlankNo + "   " + router.BlankSize + " (" + router.PartsPerBlank + " per blank)";
                    range.Item[2].Value2 = router.BlankQuantity;
                    row++;
                    // Leftover
                    range = outputSheet.get_Range("B" + row, "C" + row);
                    range.Item[2].Value2 = router.LeftoverParts;
                    row++;
                    // Parent material
                    range = outputSheet.get_Range("B" + row, "C" + row);
                    range.Item[1].Value2 = router.Material.ItemCode;
                    range.Item[2].Value2 = router.Material.QuantityPerBill + " " + router.Material.Unit;
                    row++;
                    // Color
                    range = outputSheet.get_Range("B" + row, "B" + row);
                    range.Value2 = router.Color;
                    row++;
                    // Heien/Weeke rate
                    range = outputSheet.get_Range("B" + row, "C" + row);
                    range.Item[1].Value2 = router.Cnc.QuantityPerBill + " " + router.Cnc.Unit;
                    range.Item[2].Value2 = router.Cnc.QuantityPerBill * router.Quantity + " " + router.Vector.Unit;
                    row++;
                    // Vector rate
                    range = outputSheet.get_Range("B" + row, "C" + row);
                    range.Item[1].Value2 = router.Vector.QuantityPerBill + " " + router.Vector.Unit;
                    range.Item[2].Value2 = router.Vector.QuantityPerBill * router.Quantity + " " + router.Vector.Unit;
                    row++;
                    // Pack rate
                    if (router.Assm != null)
                    {
                        range = outputSheet.get_Range("B" + row, "C" + row);
                        range.Item[1].Value2 = router.Assm.QuantityPerBill + " " + router.Assm.Unit;
                        range.Item[2].Value2 = router.Assm.QuantityPerBill * router.Quantity + " " + router.Vector.Unit;
                    }
                    row++;
                    // Regular pack
                    range = outputSheet.get_Range("B" + row, "C" + row);
                    range.Item[1].Value2 = (router.BoxItemCode == "" ? router.RegPack : "Use box: " + router.BoxItemCode);
                    range.Item[2].Value2 = router.RegPackQty;
                    row++;
                    // Super pack
                    range = outputSheet.get_Range("B" + row, "C" + row);
                    range.Item[1].Value2 = router.SupPack;
                    range.Item[2].Value2 = router.SupPackQty;
                    row++;
                    // Hardware
                    range = outputSheet.get_Range("B" + row, "B" + row);
                    range.Value2 = router.Hardware;
                    row++;

                    //#####################
                    // Box Construction
                    //#####################

                    row ++;
                    // Documentation
                    range = outputSheet.get_Range("A" + row, "A" + row);
                    range.Value2 = router.ID.ToString("D6") + (router.Copy ? " COPY" : "");
                    range.get_Characters(7, 15).Font.FontStyle = "bold";
                    range.get_Characters(7, 15).Font.Size = 20;
                    row++;
                    // Part
                    range = outputSheet.get_Range("B" + row, "C" + row);
                    range.Item[1] = router.Item.BillNo;
                    range.Item[2] = router.Quantity;
                    row++;
                    // Description
                    range = outputSheet.get_Range("B" + row, "B" + row);
                    range.Value2 = router.Item.BillDesc;
                    row++;
                    // Regular pack
                    range = outputSheet.get_Range("B" + row, "C" + row);
                    range.Item[1].Value2 = (router.BoxItemCode == "" ? router.RegPack : "Use box: " + router.BoxItemCode);
                    range.Item[2].Value2 = router.RegPackQty;
                    row++;
                    // Super pack
                    range = outputSheet.get_Range("B" + row, "C" + row);
                    range.Item[1].Value2 = router.SupPack;
                    range.Item[2].Value2 = router.SupPackQty;
                    row++;
                    // Box rate
                    if (router.Box != null)
                    {
                        range = outputSheet.get_Range("B" + row, "C" + row);
                        range.Item[1].Value2 = router.Box.QuantityPerBill + " " + router.Box.Unit;
                        range.Item[2].Value2 = router.Box.QuantityPerBill * router.Quantity + " " + router.Vector.Unit;
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
                    range.Value2 = router.TimeStamp;
                    row++;
                    try
                    {
                        // log that this these orders have been printed

                        //foreach (Order order in router.Orders)
                        //{
                        //    file.WriteLine(order.SalesOrderNo);
                        //    file.Flush();
                        //}


                        //##### Print the Cover sheet #######
                        outputSheet.PrintOut(
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        //###################################

                        // successfully printed, so we should log in the printed.json file
                        if (!router.Copy)
                        {
                            file.Write(router.Export());
                            file.Flush();
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("A problem occured when printing: " + ex.Message);
                    }
                }
            }
            file.Close();
            loadingLabel.Text = "Travelers";
        }
        private void btnCreateTravelers_Click(object sender, EventArgs e)
        {
            tableListView.Clear();
            CreateTravelers("",false);
        }

        private void login_Click(object sender, EventArgs e)
        {
            // connect to MAS
            if (ConnectToData())
            {
                login.Enabled = false;
            }
        }

        private void btnPrintSummary_Click(object sender, EventArgs e)
        {
            Summary summary = new Summary(orders, routers);
            summary.Print(workbooks);
        }

        private void btnCreateSpecificOrder_Click(object sender, EventArgs e)
        {
            // Import new orders
            if (ImportOrders(specificOrder.Text, false))
            {
                //orders = FilterByProduct(orders);
                // Create routers
                if (CreateRouters(false, false))
                {
                    loadingLabel.Text = "Travelers";
                }
            }
        }

        private void btnInvertCustomers_Click(object sender, EventArgs e)
        {
            tableListView.Clear();
            CreateTravelers("", true);
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            tableListView.Clear();
            orders.Clear();
        }
        // Create only previously printed travelers
        private void btnCreatedPrinted_Click(object sender, EventArgs e)
        {
            CreateRouters(true,false);
        }
    }
}
