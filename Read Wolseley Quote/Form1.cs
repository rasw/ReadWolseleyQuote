using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Windows.Forms;
using RASW;
using System.Data.SqlClient;
using System.Text;

namespace Read_Wolseley_Quote
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        // string connectionString = @"Data Source=(localdb)\ProjectsV13;Initial Catalog=PricingTool;Integrated Security=True;ConnectTimeout=30;Encrypt=False;TrustServerCertificate=True;ApplicationIntent=ReadWrite;MultiSubnetFailover=False";
        // string connectionString = @"Data Source=(localdb)\ProjectsV13;Initial Catalog=PricingTool;Integrated Security=True;Encrypt=False;TrustServerCertificate=True;ApplicationIntent=ReadWrite;MultiSubnetFailover=False";

        string connectionString = "Server=tcp:akvasql1.database.windows.net,1433;Initial Catalog = PricingToolDev1; Persist Security Info=False;User ID=rwilson; Password=pCj7uu573;MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout = 30;";   // Azure SQL

        AppSettings appSet = new AppSettings(Path.Combine(Directory.GetCurrentDirectory(), "ReadQuotesShowHide.xml"));
        string outputLine = "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -";
        Timer trigger = new Timer();

        public string ConnectionString { get => connectionString; set => connectionString = value; }

        private void Form1_Activated(object sender, EventArgs e)
        {
            //ReadQuoteSpreadsheet();   // seems to fire lots of times
            btnClose.Enabled = true;

            //AppSettings appSet = new AppSettings(Path.Combine(Directory.GetCurrentDirectory(), "ReadQuotesShowHide.xml"));
            if(!Convert.ToBoolean(appSet.getValue("Wolseley")))
                Dispose();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            trigger.Interval = 500;
            trigger.Tick += new EventHandler(TriggerTick);
            trigger.Start();
        }

        void TriggerTick(object sender, EventArgs e)
        {
            trigger.Stop();
            ReadQuoteSpreadsheet();
        }

        public void ReadQuoteSpreadsheet()
        {
            string[] sheetNames = { "Polarcirkel 400 FR", "Polarcirkel 450 FR", "Polarcirkel 500 FR" };
            // string[] sheetNames = { "Polarcirkel 400 FR" , "Polarcirkel 450 FR" }; //, "Polarcirkel 500 FR" };
           
            //string[] sheetNames = {"Polarcirkel 400 FR"};
            //string[] sheetNames = { "Polarcirkel 400 FR", "Polarcirkel 450 FR"};

            string filename = Path.Combine(Directory.GetCurrentDirectory() + @"\Data", "170131 Wolseley Price Quote template.xlsx");

            lstOutput.Items.Add("");
            lstOutput.Items.Add("Reading spreadsheet tabs...");

            foreach (var sheetName in sheetNames)
            {    
                lstOutput.Items.Add("Reading tab " + sheetName);        // output progress to GUI
                DataTable dt = new DataTable();

                using (OleDbConnection conn = new OleDbConnection())
                {
                    conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filename + ";" + "Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1;MAXSCANROWS=0'";
                    using (OleDbCommand comm = new OleDbCommand())
                    {
                        comm.CommandText = "Select * from [" + sheetName + "$]";
                        //comm.CommandText = "Select F1, F3, F4 from [" + sheetName + "$]";
                        // comm.CommandText = "Select C1, C2, C3, C4, C5,C6,C7,C8,C9,C10 from [" + sheetName + "$]";
                        comm.Connection = conn;

                        using (OleDbDataAdapter da = new OleDbDataAdapter())
                        {
                            da.SelectCommand = comm;
                            da.Fill(dt);
                        }
                    }
                    conn.Close();
                }

               dataGridView1.DataSource = dt.DefaultView;
                           
                List<string> sheetData  = new  List<string>();

                lstOutput.Items.Add("Processing tab data... " + sheetName);        // output progress to GUI

                using (StreamWriter errors = new StreamWriter(Path.Combine(Directory.GetCurrentDirectory() + @"\Data", "WolseleyErrors.txt"), false))
                {
                    try
                    {
                        sheetData.Add(dt.Rows[0][1].ToString());                                                                                                        // Title 
                        lstOutput.Items.Add("Found Title - " + dt.Rows[0][1].ToString());
                        sheetData.Add(dt.Rows[1][1].ToString() + "|" + dt.Rows[1][2].ToString() + "|" + dt.Rows[1][3].ToString());                                      // Pen Quantity   
                        //lstOutput.Items.Add("Processing -> " + dt.Rows[1][1].ToString()); // This is for debug only
                        sheetData.Add(dt.Rows[2][1].ToString() + "|" + dt.Rows[2][2].ToString() + "|" + dt.Rows[2][3].ToString());                                      // Pipe Spec
                        //lstOutput.Items.Add("Processing -> " + dt.Rows[2][1].ToString()); // This is for debug only
                        sheetData.Add(dt.Rows[3][1].ToString() + "|" + dt.Rows[3][2].ToString() + "|" + dt.Rows[3][3].ToString() + "|" + dt.Rows[3][4].ToString());     // Circumference
                        //lstOutput.Items.Add("Processing -> " + dt.Rows[30][1].ToString()); // This is for debug only
                        sheetData.Add(dt.Rows[4][1].ToString() + "|" + dt.Rows[4][2].ToString() + "|" + dt.Rows[4][3].ToString() + "|" + dt.Rows[4][4].ToString());     // SinkerTube
                        //lstOutput.Items.Add("Processing -> " + dt.Rows[4][1].ToString()); // This is for debug only
                    }
                    catch (Exception ex)
                    {
                        errors.WriteLine("Header section error - " + ex.Message);
                    }

                    try
                    {
                        sheetData.Add(dt.Rows[6][0].ToString() + "|" + dt.Rows[6][1].ToString() + "|" + dt.Rows[6][2].ToString() + "|" + dt.Rows[6][3].ToString() + "|" + dt.Rows[6][4].ToString() + "|" + dt.Rows[6][5].ToString() + "|" + dt.Rows[6][6].ToString() + "|" + dt.Rows[6][7].ToString() + "|" + dt.Rows[6][8].ToString() + "|" + dt.Rows[6][9].ToString());   // Column Titles

                        lstOutput.Items.Add("Processing Header Section");
                        //lstOutput.Items.Add("Processing -> " + dt.Rows[6][0].ToString());     // These are for debug only
                        //lstOutput.Items.Add("Processing -> " + dt.Rows[6][1].ToString());
                        //lstOutput.Items.Add("Processing -> " + dt.Rows[6][2].ToString());
                        //lstOutput.Items.Add("Processing -> " + dt.Rows[6][3].ToString());
                        //lstOutput.Items.Add("Processing -> " + dt.Rows[6][4].ToString());
                        //lstOutput.Items.Add("Processing -> " + dt.Rows[6][5].ToString());
                        //lstOutput.Items.Add("Processing -> " + dt.Rows[6][6].ToString());
                        //lstOutput.Items.Add("Processing -> " + dt.Rows[6][7].ToString());
                        //lstOutput.Items.Add("Processing -> " + dt.Rows[6][8].ToString());
                        //lstOutput.Items.Add("Processing -> " + dt.Rows[6][9].ToString());
                    }
                    catch (Exception ex)
                    {
                        errors.WriteLine("Column titles section error - " + ex.Message);
                    }

                    try
                    {
                        sheetData.Add("#" + dt.Rows[8][1].ToString());    // Equipment Title (# = title)
                        lstOutput.Items.Add("Processing Section -> " + dt.Rows[8][1].ToString());

                        sheetData.Add(dt.Rows[9][0].ToString() + "|" + dt.Rows[9][1].ToString() + "|" + dt.Rows[9][2].ToString() + "|" + dt.Rows[9][3].ToString() + "|" + dt.Rows[9][4].ToString() + "|" + dt.Rows[9][5].ToString() + "|" + dt.Rows[9][6].ToString() + "|" + dt.Rows[9][7].ToString() + "|" + dt.Rows[9][8].ToString() + "|" + dt.Rows[9][9].ToString());                   // Floating Structures 1

                        sheetData.Add(dt.Rows[10][0].ToString() + "|" + dt.Rows[10][1].ToString() + "|" + dt.Rows[10][2].ToString() + "|" + dt.Rows[10][3].ToString() + "|" + dt.Rows[10][4].ToString() + "|" + dt.Rows[10][5].ToString() + "|" + dt.Rows[10][6].ToString() + "|" + dt.Rows[10][7].ToString() + "|" + dt.Rows[10][8].ToString() + "|" + dt.Rows[10][9].ToString());   // Floating Structures 2
                    }
                    catch (Exception ex)
                    {
                        errors.WriteLine("Floating Structures section error - " + ex.Message);
                    }

                    try
                    {
                        sheetData.Add("#" + dt.Rows[11][1].ToString());    // Equipment Title  (# = title)

                        sheetData.Add(dt.Rows[12][0].ToString() + "|" + dt.Rows[12][1].ToString() + "|" + dt.Rows[12][2].ToString() + "|" + dt.Rows[12][3].ToString() + "|" + dt.Rows[12][4].ToString() + "|" + dt.Rows[12][5].ToString() + "|" + dt.Rows[12][6].ToString() + "|" + dt.Rows[12][7].ToString() + "|" + dt.Rows[12][8].ToString() + "|" + dt.Rows[12][9].ToString());   // Sinker tube

                    }
                    catch (Exception ex)
                    {
                        errors.WriteLine("Sinkertube section error - " + ex.Message);
                    }

                    try
                    {
                        sheetData.Add("#" + dt.Rows[13][1].ToString());    // Equipment Title  (# = title) 

                        sheetData.Add(dt.Rows[14][0].ToString() + "|" + dt.Rows[14][1].ToString() + "|" + dt.Rows[14][2].ToString() + "|" + dt.Rows[14][3].ToString() + "|" + dt.Rows[14][4].ToString() + "|" + dt.Rows[14][5].ToString() + "|" + dt.Rows[14][6].ToString() + "|" + dt.Rows[14][7].ToString() + "|" + dt.Rows[14][8].ToString() + "|" + dt.Rows[14][9].ToString());   // Walkways

                    }
                    catch (Exception ex)
                    {
                        errors.WriteLine("Walkways section error - " + ex.Message);
                    }

                    StreamWriter file = new StreamWriter(Path.Combine(Directory.GetCurrentDirectory() + @"\Data", sheetName + ".ecf"), false);   // overwrite file each time
                    foreach (var line in sheetData)
                    {
                        file.WriteLine(line);
                    }

                    file.Close();
                    lstOutput.Items.Add("End of tab processing....");
                    lstOutput.Items.Add(outputLine);
                }

                try
                {
                    string path = Path.Combine(Directory.GetCurrentDirectory() + @"\Data", "WolseleyErrors.txt");

                    if (new FileInfo(path).Length == 0)
                    {
                        File.Delete(path);   // if file size is equal to zero then no errors have been reported so delete the file
                        lstOutput.Items.Add("Import processing complete...");        // output progress to GUI
                    }
                    else
                    {
                        lstOutput.Items.Clear();
                        lstOutput.BackColor = System.Drawing.Color.Crimson;
                        lstOutput.ForeColor = System.Drawing.Color.White;
                        using (StreamReader reader = new StreamReader(path))
                        {
                            string line = null;
                            while ((line = reader.ReadLine()) != null)
                            {
                                lstOutput.Items.Add(line); // Add to list.
                            }
                        }
                    }
                }catch{}

               string id = ProcessBaseDataToDB(sheetData);
               ProcessFloatingStructuresDataToDB(sheetData, id);
               ProcessSinkerTubeDataToDB(sheetData, id);
              // ProcessWalkwaysDataToDB(sheetData, id);
            }
        }

        private string ProcessBaseDataToDB(List<string> listData)
        {
            /*  >> listData Contains The full contents
                [0] Polarcirkel 400 cage
                [1] Pen Quantity||1
                [2] Pipe spec|SDR|17
                [3] Circumference||90|m
                [4] Sinkertube||1|pc

                [5] Item no/Part No|Equipment|Unit|Quantity|Price per M|discount  %|Price per ton|Del Y/N|12|13.5
                [6] #Floating Structures
                [7]     |Inner floating pipe SDR17|m|186||0|0||15.5|13.7777777777778
                [8]     |Handrail pipe 140mm SD11|m|90||0|0||7.5|
                [9] #Sinkertube
                [10]     |Pipe 250 SDR11|m|96||0|0||8|7.11111111111111
                [11] #Decking pipe
                [12]     |Pipe 50x 100m coils SDR11|pc|200||0|0|||
             */

            string wolseleyID = String.Empty;
            object newID = 0;

            try
            {
               wolseleyID = GetScalarData("Select SupplierID FROM dbo.Supplier WHERE SupplierName = 'Wolseley'");     // Get the id from the DB - get the wolseley supplierID
            }
            catch (Exception ex)
            {
                Logger.WriteToLog(ex.Message);
            }

            try
            {
                List<string> sqlScriptLines = new List<string>();       // holds the sql script to update and insert into database
                StringBuilder scriptLine = new StringBuilder();         // builder for the query
               
                string baseSQL = "SET DATEFORMAT dmy; INSERT INTO CageQuotes ([QuoteDateTime], [QuoteSupplierID], [QuoteReference], [QuoteCageType], [QuoteQuantity], [QuotePipeSpec], [QuoteCircumference], [QuoteCircumferenceUnits], [QuoteSinkertube], [QuoteSinkertubeUnits]";
                string midSQL = ") VALUES (";
                string endSQL = "); SELECT SCOPE_IDENTITY();";
             
                scriptLine.Append(baseSQL);
                scriptLine.Append(midSQL);
                scriptLine.Append("'" + DateTime.Now + "',");
                scriptLine.Append(wolseleyID + ",'','"); 
                scriptLine.Append(listData[0] + "',");   // QuoteCageType

                string[] a = listData[1].Split('|');
                scriptLine.Append("'" + a[2] + "',");    // QuoteQuantity

                string[] b = listData[2].Split('|');
                scriptLine.Append(b[2] + ",");           // QuotePipeSpec
                                
                string[] c = listData[3].Split('|');
                scriptLine.Append(c[2] + ",");           // QuoteCircumference
                scriptLine.Append("'" + c[3] + "',");    // QuoteCircumferenceUnits

                string[] d = listData[4].Split('|');
                scriptLine.Append(d[2] + ",");           // QuoteSinkertube
                scriptLine.Append("'" + d[3] + "'");     // QuoteSinkertubeUnits
                scriptLine.Append(endSQL);
                
                try
                {
                    using (SqlConnection cnn = new SqlConnection(ConnectionString))
                    {
                        cnn.Open();
                        using (SqlCommand cmd = new SqlCommand(scriptLine.ToString(), cnn))
                        {
                             newID = cmd.ExecuteScalar();
                            //cmd.ExecuteNonQuery();
                            cmd.Dispose();
                            cnn.Close();
                        }
                    }
                }
                catch (Exception ex)
                {
                    lstOutput.Items.Add("");
                    lstOutput.Items.Add("ERROR: " + ex.Message);
                    lstOutput.Items.Add("");
                    return "";
                }
            }
            catch (Exception ex)
            {
                lstOutput.Items.Add("");
                lstOutput.Items.Add("ERROR: " + ex.Message);
                lstOutput.Items.Add("");
                return "";
            }

            return newID.ToString();
        }

        private void ProcessFloatingStructuresDataToDB(List<string> listData, string QuoteID)
        {
            /*  >> listData Contains The full contents
                [0] Polarcirkel 400 cage
                [1] Pen Quantity||1
                [2] Pipe spec|SDR|17
                [3] Circumference||90|m
                [4] Sinkertube||1|pc
                
            [5] Item no/Part No|Equipment|Unit|Quantity|Price per M|discount  %|Price per ton|Del Y/N|12|13.5
                [6] #Floating Structures
                [7] |Inner floating pipe SDR17|m|186||0|0||15.5|13.7777777777778
                [8] |Handrail pipe 140mm SD11|m|90||0|0||7.5|
                
            [9] #Sinkertube
                [10] |Pipe 250 SDR11|m|96||0|0||8|7.11111111111111
                
            [11] #Decking pipe
                [12] |Pipe 50x 100m coils SDR11|pc|200||0|0|||
             */

            // 1.   Get quote ID
            // 2.   Get the data rows
            // 3.   Add them to DB

            int line = 7;  // 7;

            try
            {
                List<string> sqlScriptLines = new List<string>();       // holds the sql script to update and insert into database
                StringBuilder scriptLine = new StringBuilder();         // builder for the query

                string baseSQL = "INSERT INTO [dbo].[FloatingStructures] ([QuoteID],[FloatingItemPartNo],[FloatingEquipment],[FloatingUnit],[FloatingQuantity],[FloatingPricePerM],[FloatingDiscountPercentage],[FloatingPrice PerTon],[FloatingDelivery],[FloatingCost12],[FloatingCost135]";
                string midSQL = ") VALUES (";
                string endSQL = ")";

                string sqs = "'";

                while (listData[line - 1].ToLower() != "#sinkertube")       // Process ** Floating Structures **
                { 
                    scriptLine.Append(baseSQL);
                    scriptLine.Append(midSQL);
                    scriptLine.Append("'" + QuoteID + "',");                // 'QuoteID',

                    string[] s = listData[line].Split('|');                 // Length values split

                    if(s.Length == 1)
                    {
                        break;
                    }

                    scriptLine.Append(sqs + s[0] + "',");                   // FloatingItemPartNo
                    scriptLine.Append(sqs + s[1] + "',");                   // FloatingEquipment
                    scriptLine.Append(sqs + s[2] + "',");                   // FloatingUnit
                    scriptLine.Append(sqs + s[3] + "',");                   // FloatingQuantity

                    scriptLine.Append(sqs + s[4] + "',");                   // FloatingPricePerM

                    if (s[5].Length > 0)                                    // FloatingDiscount %
                    {
                        if (s[5].EndsWith("%"))
                        {
                            s[5] = s[5].Substring(0, s[5].Length - 1);
                        }
                    }

                    scriptLine.Append(sqs + s[5] + "',");                   // FloatingDiscount %

                    if (s[6].StartsWith("£"))                               // FloatingPrice PerTon  (Remove any £ signs)
                    {
                        s[6] = s[6].Substring(1, s[6].Length-1);
                        scriptLine.Append(sqs + s[6] + "',");
                    }
                    else
                    {
                        scriptLine.Append(sqs + s[6] + "',");               
                    }
                    
                    scriptLine.Append(sqs + s[7] + "',");                   // FloatingDelivery 
                    scriptLine.Append(sqs + s[8] + "',");                   // FloatingCost12
                    scriptLine.Append(sqs + s[9] + "'");                    // FloatingCost135
                    line++;

                    scriptLine.Append(endSQL);
                    
                    try
                    {
                        using (SqlConnection cnn = new SqlConnection(ConnectionString))
                        {
                            cnn.Open();
                            using (SqlCommand cmd = new SqlCommand(scriptLine.ToString(), cnn))
                            {
                                cmd.ExecuteNonQuery();
                                cmd.Dispose();
                                cnn.Close();
                                scriptLine.Clear();
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        scriptLine.Clear();
                        lstOutput.Items.Add("");
                        lstOutput.Items.Add("Database ERROR: " + ex.Message);
                        lstOutput.Items.Add("");
                    }
                }
            }
            catch (Exception ex)
            {
                lstOutput.Items.Add("");
                lstOutput.Items.Add("ERROR: " + ex.Message);
                lstOutput.Items.Add("");
            }
        }


        // process sinkertubes
        private void ProcessSinkerTubeDataToDB(List<string> listData, string QuoteID)
        {
            /*  >> listData Contains The full contents
                [0] Polarcirkel 400 cage
                [1] Pen Quantity||1
                [2] Pipe spec|SDR|17
                [3] Circumference||90|m
                [4] Sinkertube||1|pc
                
            [5] Item no/Part No|Equipment|Unit|Quantity|Price per M|discount  %|Price per ton|Del Y/N|12|13.5
                [6] #Floating Structures
                [7] |Inner floating pipe SDR17|m|186||0|0||15.5|13.7777777777778
                [8] |Handrail pipe 140mm SD11|m|90||0|0||7.5|
                
            [9] #Sinkertube
                [10] |Pipe 250 SDR11|m|96||0|0||8|7.11111111111111
                
            [11] #Decking pipe
                [12] |Pipe 50x 100m coils SDR11|pc|200||0|0|||
             */

            // 1.   Get quote ID
            // 2.   Get the data rows
            // 3.   Add them to DB

            int line = 0;

            // get the start line for sinkertube data
            for (int i = 0; i < listData.Count; i++)
            {
                if(listData[i].Substring(0,5).ToLower() == "#sink")  // #SinkerTube
                {
                    line = i;
                }
            }

            line++;     // point to next line

            try
            {
                List<string> sqlScriptLines = new List<string>();       // holds the sql script to update and insert into database
                StringBuilder scriptLine = new StringBuilder();         // builder for the query

                string baseSQL = "INSERT INTO [dbo].[Sinkertube] ([QuoteID],[SinkertubeItemPartNo],[SinkertubeEquipment],[SinkertubeUnit],[SinkertubeQuantity],[SinkertubePricePerM],[SinkertubeDiscountPercentage],[SinkertubePrice PerTon],[SinkertubeDelivery],[SinkertubeCost12],[SinkertubeCost135] ";
                string midSQL = ") VALUES (";
                string endSQL = ")";

                string sqs = "'";

                while (listData[line].Substring(0, 5).ToLower() != "#walk")       // Process ** SinkerTubes
                {
                    scriptLine.Append(baseSQL);
                    scriptLine.Append(midSQL);
                    scriptLine.Append("'" + QuoteID + "',");                // 'QuoteID',

                    string[] s = listData[line].Split('|');                 // Length values split

                    if (s.Length == 1)
                    {
                        break;
                    }

                    scriptLine.Append(sqs + s[0] + "',");                   // SinkertubeItemPartNo
                    scriptLine.Append(sqs + s[1] + "',");                   // SinkertubeEquipment
                    scriptLine.Append(sqs + s[2] + "',");                   // SinkertubeUnit
                    scriptLine.Append(sqs + s[3] + "',");                   // SinkertubeQuantity

                    scriptLine.Append(sqs + s[4] + "',");                   // SinkertubePricePerM

                    if (s[5].Length > 0)                                    // SinkerDiscount %
                    {
                        if (s[5].EndsWith("%"))
                        {
                            s[5] = s[5].Substring(0, s[5].Length - 1);
                        }
                    }

                    scriptLine.Append(sqs + s[5] + "',");                   // SinkerDiscount %

                    if (s[6].StartsWith("£"))                               // SinkertubePrice PerTon  (Remove any £ signs)
                    {
                        s[6] = s[6].Substring(1, s[6].Length - 1);
                        scriptLine.Append(sqs + s[6] + "',");
                    }
                    else
                    {
                        scriptLine.Append(sqs + s[6] + "',");
                    }

                    scriptLine.Append(sqs + s[7] + "',");                   // SinkertubeDelivery 
                    scriptLine.Append(sqs + s[8] + "',");                   // SinkertubeCost12
                    scriptLine.Append(sqs + s[9] + "'");                    // SinkertubeCost135
                    line++;

                    scriptLine.Append(endSQL);

                    try
                    {
                        using (SqlConnection cnn = new SqlConnection(ConnectionString))
                        {
                            cnn.Open();
                            using (SqlCommand cmd = new SqlCommand(scriptLine.ToString(), cnn))
                            {
                                cmd.ExecuteNonQuery();
                                cmd.Dispose();
                                cnn.Close();
                                scriptLine.Clear();
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        scriptLine.Clear();
                        lstOutput.Items.Add("");
                        lstOutput.Items.Add("Database ERROR: " + ex.Message);
                        lstOutput.Items.Add("");
                    }
                }
            }
            catch (Exception ex)
            {
                lstOutput.Items.Add("");
                lstOutput.Items.Add("ERROR: " + ex.Message);
                lstOutput.Items.Add("");
            }
        }

       
        private string GetScalarData(string query)  // retrieve a single column result
        {
            using (SqlConnection cnn = new SqlConnection(ConnectionString))
            {
                using (SqlCommand cmd = new SqlCommand())
                {
                    object returnValue;
                    cmd.CommandText = query;
                    cmd.CommandType = CommandType.Text;
                    cmd.Connection = cnn;
                    cnn.Open();
                    returnValue = cmd.ExecuteScalar();
                    cnn.Close();
                    return returnValue.ToString();
                }
            }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void chkShowGUI_CheckedChanged(object sender, EventArgs e)
        {
            //AppSettings appSet = new AppSettings(Path.Combine(Directory.GetCurrentDirectory(), "ReadQuotesShowHide.xml"));
            if(chkShowGUI.Checked)
                appSet.setValue("Wolseley", true.ToString());
            else
                appSet.setValue("Wolseley", false.ToString());
        }

       
        private void btnSave_Click(object sender, EventArgs e)
        {
         
        }

        //private void GetData()
        //{
        //    //string queryString = "SELECT SupplierName, SupplierAddress1, SupplierPostCode FROM dbo.Supplier;";

        //    //using (SqlConnection connection = new SqlConnection(connectionString))
        //    //{
        //    //    SqlCommand command = new SqlCommand(queryString, connection);

        //    //    connection.Open();

        //    //    using (SqlDataReader reader = command.ExecuteReader())
        //    //    {
        //    //        try
        //    //        {
        //    //            while (reader.Read())
        //    //            {
        //    //                listBox1.Items.Add(reader[0] + " - " + reader[1] + " - " + reader[2]);
        //    //            }
        //    //        }
        //    //        finally
        //    //        {
        //    //            // Always call Close when done reading.
        //    //            reader.Close();
        //    //        }
        //    //    }
        //    //}
        //}

        private void btnClose_Click(object sender, EventArgs e)
        {
            Dispose();
        }
    }
}
