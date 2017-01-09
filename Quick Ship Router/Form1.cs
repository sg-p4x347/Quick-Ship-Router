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
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            infoLabel.Text = "";
            excelApp = new Excel.Application();
            excelApp.DisplayAlerts = false;
            // Open the traveler template
            workbooks = excelApp.Workbooks;

            PopulateCustomers();
            // connect to MAS
                
            InitializeManagers();
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
        private string sortInfo = "";
        private bool invertCustomers = false;
        private Summary summary = null;
        // Opens a connection to the MAS database
        private void ConnectToData()
        {
            infoLabel.Text = "Logging in...";
            MAS = new OdbcConnection();
            // initialize the MAS connection
            MAS.ConnectionString = "DSN=SOTAMAS90;Company=MGI;";
            //MAS.ConnectionString = "DSN=SOTAMAS90;Company=MGI;UID=GKC;PWD=sgp4x347;";
            try
            {
                MAS.Open();
                login.Enabled = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to log in :(");
            }
            infoLabel.Text = "";
        }
        private void InitializeManagers()
        {
            ConnectToData();

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

            tableManager = new TableManager(MAS, infoLabel, progressBar, tableListView,crossRef,boxRef,blankRef,colorRef);
            chairManager = new ChairManager(MAS, infoLabel, progressBar, chairListView);
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
        private bool CreateTravelers(string specificOrderNo, bool invertCustomers)
        {
            // Import new orders
            if (ImportOrders(specificOrderNo,invertCustomers))
            {
                // Create travelers
                tableManager.CompileTravelers(backgroundWorker1);
                chairManager.CompileTravelers(backgroundWorker1);
                return true;
            }
            return false;
        }
       
        
        //======================
        // Open Order
        //======================
        private bool ImportOrders(string specificOrderNo, bool invertCustomers)
        {
            //infoLabel.Text = "Importing Orders...";

            string today = DateTime.Today.ToString(@"yyyy\-MM\-dd");
            // OrderDate >= {d '" +  todayString + "'}
            string customerNames = "";
            string displayNames = "";
            for (int i = 0; i < customerList.Items.Count; i++)
            {
                if (customerList.GetItemCheckState(i) == CheckState.Checked)
                {
                    customerNames += (customerNames.Length > 0 ? "," : "") + "'" + customerList.GetItemText(customerList.Items[i]) + "'";
                    displayNames += (displayNames.Length > 0 ? ", " : "") + customerList.GetItemText(customerList.Items[i]);
                }
            }
            
            sortInfo = invertCustomers ? "" : "(" + displayNames + ")";
            // get informatino from header
            OdbcCommand command = MAS.CreateCommand();
            command.CommandText = "SELECT SalesOrderNo, ShipExpireDate, CustomerNo, ShipVia FROM SO_SalesOrderHeader WHERE CustomerNo " + (invertCustomers ? "NOT" : "") + " IN (" + customerNames + ")" + (showToday.Checked ? "AND OrderDate >= {d '" + today + "'}" : "");
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
                    if (!detailReader.IsDBNull(2) && detailReader.GetString(2) != "KIT")
                    {
                        if (IsTable(billCode))
                        {
                            // this is a table
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

                            order.ItemCode = billCode;
                            order.QuantityOrdered = Convert.ToInt32(detailReader.GetValue(1));
                            tableManager.Orders.Add(order);
                            continue;
                        } else if (IsChair(billCode))
                        {
                            // this is a table
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

                            order.ItemCode = billCode;
                            order.QuantityOrdered = Convert.ToInt32(detailReader.GetValue(1));
                            chairManager.Orders.Add(order);
                            continue;
                        }

                    }
                }
                detailReader.Close();
            }
            reader.Close();
            //infoLabel.Text = "";
            return true;
        }
        private bool IsTable(string s)
        {
            return (s.Length == 9 && s.Substring(0, 2) == "MG") || (s.Length == 10 && (s.Substring(0, 3) == "38-" || s.Substring(0, 3) == "41-"));
        }
        private bool IsChair(string s)
        {
            if (s.Substring(0,2) == "38")
            {
                string[] parts = s.Split('-');
                return (parts[0].Length == 5 && parts[1].Length == 4 && parts[2].Length == 3);
            } else
            {
                return false;
            }
            
        }
        //======================
        // Tables
        //======================
        private TableManager tableManager = new TableManager();
        //======================
        // Chairs
        //======================
        private ChairManager chairManager = new ChairManager();
        //======================
        // Events
        //======================
        // Print
        private void btnPrint_Click(object sender, EventArgs e)
        {
            infoLabel.Text = "Printing Travelers...";
            string exeDir = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            //var workbook = workbooks.Open(System.IO.Path.Combine(exeDir, "traveler.xlsx"),
            //    0, false, 5, "", "", false, 2, "",
            //    true, false, 0, true, false, false);
            var workbook = workbooks.Open(System.IO.Path.Combine(exeDir, "traveler.xlsx"),
                0, false, 5, "", "", false, 2, "",
                true, false, 0, true, false, false);
            var worksheets = workbook.Worksheets;

            tableManager.PrintTravelers(worksheets);
            chairManager.PrintTravelers(worksheets);
            
            
            //// create the output workbook
            //int currentSheet = 2;
            //for (int itemIndex = 0; itemIndex < tableListView.Items.Count; itemIndex++)
            //{
            //    if (tableListView.Items[itemIndex].Checked)
            //    {
            //        Router router = routers[itemIndex];
            //        templateSheet.Copy(Type.Missing, workbook.Worksheets[currentSheet - 1]);

            //        Excel.Worksheet outputSheet = workbook.Worksheets[currentSheet];

            //        // Sales Orders
            //        string customerList = "";
            //        string orderList = "";
            //        int i = 0;
            //        foreach (Order order in router.Orders)
            //        {
            //            customerList += (i == 0 ? "" : ", ") + order.CustomerNo;
            //            orderList += (i == 0 ? "" : ", ") + "(" + order.QuantityOrdered + ") " + order.SalesOrderNo;
            //            i++;
            //        }
            //        //#####################
            //        // Production Traveler
            //        //#####################
            //        Excel.Range range;
            //        int row = 1;
            //        // Documentation
            //        range = outputSheet.get_Range("A" + row, "A" + row);
            //        range.Value2 = router.ID.ToString("D6") + (router.Copy ? " COPY" : "");
            //        range.get_Characters(7, 15).Font.FontStyle = "bold";
            //        range.get_Characters(7, 15).Font.Size = 20;
            //        row++;
            //        // Part
            //        range = outputSheet.get_Range("B" + row, "C" + row);
            //        range.Item[1] = router.Item.BillNo;
            //        range.Item[2] = router.Quantity;
            //        row++;
            //        // Description
            //        range = outputSheet.get_Range("B" + row, "B" + row);
            //        range.Value2 = router.Item.BillDesc;
            //        row++;
            //        // Drawing
            //        range = outputSheet.get_Range("B" + row, "B" + row);
            //        range.Value2 = router.Item.DrawingNo;
            //        row++;
            //        // Sales Orders
            //        range = outputSheet.get_Range("B" + row, "C" + row);
            //        range.Item[1].Value2 = orderList;
            //        row++;
            //        // Customers
            //        range = outputSheet.get_Range("B" + row, "C" + row);
            //        range.Item[1].Value2 = customerList;
            //        row++;
            //        // Date printed
            //        router.TimeStamp = DateTime.Now.ToString("MM/dd/yyyy   hh:mm tt");
            //        range = outputSheet.get_Range("B" + row, "B" + row);
            //        range.Value2 = router.TimeStamp;
            //        row++;
            //        // Blank
            //        range = outputSheet.get_Range("B" + row, "C" + row);
            //        range.Item[1].Value2 = router.BlankNo + "   " + router.BlankSize + " (" + router.PartsPerBlank + " per blank)";
            //        range.Item[2].Value2 = router.BlankQuantity;
            //        row++;
            //        // Leftover
            //        range = outputSheet.get_Range("B" + row, "C" + row);
            //        range.Item[2].Value2 = router.LeftoverParts;
            //        row++;
            //        // Parent material
            //        range = outputSheet.get_Range("B" + row, "C" + row);
            //        range.Item[1].Value2 = router.Material.ItemCode;
            //        range.Item[2].Value2 = router.Material.QuantityPerBill + " " + router.Material.Unit;
            //        row++;
            //        // Color
            //        range = outputSheet.get_Range("B" + row, "B" + row);
            //        range.Value2 = router.Color;
            //        row++;
            //        // Heien/Weeke rate
            //        range = outputSheet.get_Range("B" + row, "C" + row);
            //        range.Item[1].Value2 = router.Cnc.QuantityPerBill + " " + router.Cnc.Unit;
            //        range.Item[2].Value2 = router.Cnc.QuantityPerBill * router.Quantity + " " + router.Vector.Unit;
            //        row++;
            //        // Vector rate
            //        range = outputSheet.get_Range("B" + row, "C" + row);
            //        range.Item[1].Value2 = router.Vector.QuantityPerBill + " " + router.Vector.Unit;
            //        range.Item[2].Value2 = router.Vector.QuantityPerBill * router.Quantity + " " + router.Vector.Unit;
            //        row++;
            //        // Pack rate
            //        if (router.Assm != null)
            //        {
            //            range = outputSheet.get_Range("B" + row, "C" + row);
            //            range.Item[1].Value2 = router.Assm.QuantityPerBill + " " + router.Assm.Unit;
            //            range.Item[2].Value2 = router.Assm.QuantityPerBill * router.Quantity + " " + router.Vector.Unit;
            //        }
            //        row++;
            //        // Regular pack
            //        range = outputSheet.get_Range("B" + row, "C" + row);
            //        range.Item[1].Value2 = (router.BoxItemCode == "" ? router.RegPack : "Use box: " + router.BoxItemCode);
            //        range.Item[2].Value2 = router.RegPackQty;
            //        row++;
            //        // Super pack
            //        range = outputSheet.get_Range("B" + row, "C" + row);
            //        range.Item[1].Value2 = router.SupPack;
            //        range.Item[2].Value2 = router.SupPackQty;
            //        row++;
            //        // Hardware
            //        range = outputSheet.get_Range("B" + row, "B" + row);
            //        range.Value2 = router.Hardware;
            //        row++;

            //        //#####################
            //        // Box Construction
            //        //#####################

            //        row ++;
            //        // Documentation
            //        range = outputSheet.get_Range("A" + row, "A" + row);
            //        range.Value2 = router.ID.ToString("D6") + (router.Copy ? " COPY" : "");
            //        range.get_Characters(7, 15).Font.FontStyle = "bold";
            //        range.get_Characters(7, 15).Font.Size = 20;
            //        row++;
            //        // Part
            //        range = outputSheet.get_Range("B" + row, "C" + row);
            //        range.Item[1] = router.Item.BillNo;
            //        range.Item[2] = router.Quantity;
            //        row++;
            //        // Description
            //        range = outputSheet.get_Range("B" + row, "B" + row);
            //        range.Value2 = router.Item.BillDesc;
            //        row++;
            //        // Regular pack
            //        range = outputSheet.get_Range("B" + row, "C" + row);
            //        range.Item[1].Value2 = (router.BoxItemCode == "" ? router.RegPack : "Use box: " + router.BoxItemCode);
            //        range.Item[2].Value2 = router.RegPackQty;
            //        row++;
            //        // Super pack
            //        range = outputSheet.get_Range("B" + row, "C" + row);
            //        range.Item[1].Value2 = router.SupPack;
            //        range.Item[2].Value2 = router.SupPackQty;
            //        row++;
            //        // Box rate
            //        if (router.Box != null)
            //        {
            //            range = outputSheet.get_Range("B" + row, "C" + row);
            //            range.Item[1].Value2 = router.Box.QuantityPerBill + " " + router.Box.Unit;
            //            range.Item[2].Value2 = router.Box.QuantityPerBill * router.Quantity + " " + router.Vector.Unit;
            //        }
            //        row++;
            //        // Sales Orders
            //        range = outputSheet.get_Range("B" + row, "C" + row);
            //        range.Item[1].Value2 = orderList;
            //        row++;
            //        // Customers
            //        range = outputSheet.get_Range("B" + row, "C" + row);
            //        range.Item[1].Value2 = customerList;
            //        row++;
            //        // Date printed
            //        range = outputSheet.get_Range("B" + row, "B" + row);
            //        range.Value2 = router.TimeStamp;
            //        row++;
            //        try
            //        {
            //            // log that this these orders have been printed

            //            //foreach (Order order in router.Orders)
            //            //{
            //            //    file.WriteLine(order.SalesOrderNo);
            //            //    file.Flush();
            //            //}


            //            //##### Print the Cover sheet #######
            //            outputSheet.PrintOut(
            //                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            //                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            //            //###################################

            //            // successfully printed, so we should log in the printed.json file
            //            if (!router.Copy)
            //            {
            //                file.Write(router.Export());
            //                file.Flush();
            //            }
            //        }
            //        catch (Exception ex)
            //        {
            //            MessageBox.Show("A problem occured when printing: " + ex.Message);
            //        }
            //    }
            //}
            //file.Close();
            //infoLabel.Text = "Travelers";
        }
        // Print summary
        private void btnPrintSummary_Click(object sender, EventArgs e)
        {
            if (summary == null)
            {
                summary = new Summary(tableManager.Travelers, chairManager.Travelers, sortInfo);
                summary.Print(workbooks);
            } else
            {
                summary.Print(workbooks);
            }
            
        }
        // Create Travelers (from selected customers)
        private void btnCreateTravelers_Click(object sender, EventArgs e)
        {
            summary = null; // a new summary will need to be created
            invertCustomers = false;
            tableListView.Clear();
            backgroundWorker1.RunWorkerAsync();
        }
        // Create Travelers(from unselected customers)
        private void btnInvertCustomers_Click(object sender, EventArgs e)
        {
            summary = null; // a new summary will need to be created
            invertCustomers = true;
            tableListView.Clear();
            backgroundWorker1.RunWorkerAsync();
        }
        // Login to MAS
        private void login_Click(object sender, EventArgs e)
        {
            // connect to MAS and initialize managers
            InitializeManagers();
        }
        
        // Add specific order
        private void btnCreateSpecificOrder_Click(object sender, EventArgs e)
        {
            // Import new orders
            //if (ImportOrders(specificOrder.Text, false))
            //{
            //    //orders = FilterByProduct(orders);
            //    // Create routers
            //    if (CreateRouters(false, false))
            //    {
            //        infoLabel.Text = "Travelers";
            //    }
            //}
        }
        // Clear Travelers
        private void btnClear_Click(object sender, EventArgs e)
        {
            tableListView.Clear();
            tableManager.Orders.Clear();
            tableManager.Travelers.Clear();
            chairListView.Clear();
            chairManager.Orders.Clear();
            chairManager.Travelers.Clear();
        }
        // Create only previously printed travelers
        private void btnCreatedPrinted_Click(object sender, EventArgs e)
        {
            //CreateRouters(true,false);
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
                CreateTravelers("", invertCustomers);
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar.Visible = true;
            infoLabel.Text = e.UserState.ToString();
            progressBar.Value = e.ProgressPercentage;
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            tableManager.DisplayTravelers();
            chairManager.DisplayTravelers();
            infoLabel.Text = "Complete";
            progressBar.Visible = false;
        }
    }
}
