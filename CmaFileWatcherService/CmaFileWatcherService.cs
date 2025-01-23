using Dapper;
using IniParser;
using IniParser.Model;
using Microsoft.Data.SqlClient;
using Syncfusion.XlsIO;
using System;
using System.IO;
using System.ServiceProcess;


namespace CmaFileWatcherService
{
    // cd \source\repos\CmaFileWatcherSvc\CmaFileWatcherService\bin\Debug
    // sc delete CmaFileWatcherService
    // sc create CmaFileWatcherService binPath= "D:\source\repos\CmaFileWatcherSvc\CmaFileWatcherService\bin\Debug\CmaFileWatcherService.exe"

    public partial class CmaFileWatcherService : ServiceBase
    {
        private FileSystemWatcher _fileWatcher;
        private string _folderPath;

        public CmaFileWatcherService()
        {
            //InitializeComponent();
            // _fileWatcher = new FileSystemWatcher();

            InitializeComponent();
            LoadConfiguration();
            InitializeWatcher();
        }
        private void LoadConfiguration()
        {
            var parser = new FileIniDataParser();
            IniData data = parser.ReadFile("config.ini");
            _folderPath = data["Settings"]["WatchFolder"];
        }

        private void InitializeWatcher()
        {
            _fileWatcher = new FileSystemWatcher
            {
                Path = _folderPath,
                Filter = "*.xlsx",
                NotifyFilter = NotifyFilters.FileName | NotifyFilters.LastWrite
            };

            _fileWatcher.Created += OnNewExcelFile;
        }


        protected override void OnStart(string[] args)
        {
            // _fileWatcher.Path = @"C:\CMAInbound"; // Folder to watch
            //  _fileWatcher.Filter = "*.xlsx"; // Watch for Excel files
            // _fileWatcher.NotifyFilter = NotifyFilters.FileName | NotifyFilters.LastWrite;
            //  _fileWatcher.Created += OnNewExcelFile;
            _fileWatcher.EnableRaisingEvents = true;

            WriteLog("Service started and folder watching enabled.");
        }

        protected override void OnStop()
        {
            _fileWatcher.EnableRaisingEvents = false;
            WriteLog("Service stopped.");
        }

        private void OnNewExcelFile(object sender, FileSystemEventArgs e)
        {
            try
            {
                // Delay to ensure the file is fully copied before processing
                System.Threading.Thread.Sleep(5000);
                WriteLog($"New file detected: {e.FullPath}");
                ProcessExcelFile(e.FullPath);
            }
            catch (Exception ex)
            {
                WriteLog($"Error processing file {e.FullPath}: {ex.Message}");
            }
        }

        private void ProcessExcelFile(string filePath)
        {
            // Read Excel data using Syncfusion
            using (var inputStream = new FileStream(filePath, FileMode.Open))
            {
                using (ExcelEngine excelEngine = new ExcelEngine())
                {
                    IApplication application = excelEngine.Excel;
                    IWorkbook workbook = application.Workbooks.Open(inputStream);
                    IWorksheet worksheet = workbook.Worksheets["CMA Template"];

                    // Find the row that contains "Corporate Customer Number"
                    // Use that to determine what row  
                    int rowNumber = 0;
                    for (int i = 1; i <= worksheet.Rows.Length; i++)
                    {
                        if (worksheet.Range["A" + i].DisplayText == "Corporate Customer Name")
                        {
                            rowNumber = i;
                            break;
                        }
                    }

                    string promoTermsText = "";
                    string promoFreightTermsText = "";
                    string promoFreightMinimumsText = "";
                    string PcfTypeText = "";
                    string promoFreightTermsOtherAmtText = "";

                    // Extract data (example: Customer Name from cell A1)
                    //string customerNumber = worksheet.Range["B7"].DisplayText;
                    string customerNumber = worksheet.Range["B" + (rowNumber + 2)].DisplayText;
                    string corpNumber = worksheet.Range["B" + (rowNumber + 1)].DisplayText;
                    string customerName = worksheet.Range["B" + rowNumber].DisplayText;
                    customerNumber = !string.IsNullOrEmpty(corpNumber) ? corpNumber : customerNumber;
                    int corpFlag = !string.IsNullOrEmpty(corpNumber) ? 1 : 0;
                    string buyingGroup = worksheet.Range["K" + (rowNumber + 2)].DisplayText;
                    string SubmittedBy = worksheet.Range["M" + (rowNumber + 2)].Text;
                    DateTime startDate = DateTime.Now;
                    DateTime endDate = DateTime.Now;
                    if (rowNumber != 1)  // this must be a new CMA format
                    {
                        promoTermsText = worksheet.Range["F2"].DisplayText;
                        promoFreightTermsText = worksheet.Range["C2"].DisplayText;
                        promoFreightMinimumsText = worksheet.Range["D2"].DisplayText;
                        PcfTypeText = worksheet.Range["B2"].DisplayText;
                        promoFreightTermsOtherAmtText = worksheet.Range["D3"].DisplayText;
                    }
                    else  // original format did not have these fields
                    {
                        promoTermsText = "";
                        promoFreightTermsText = "";
                        promoFreightMinimumsText = "";
                        PcfTypeText = "";
                        promoFreightTermsOtherAmtText = "";
                    }





                    string site = "BAT";

                    string cmaFilename = Path.GetFileName(filePath);
                    string status = "N"; // New


                    object cellValue = worksheet.Range["P" + (rowNumber + 0)].Value;
                    if (cellValue != null && DateTime.TryParse(cellValue.ToString(), out DateTime startDateTime))
                    {
                        // Successfully parsed the date

                        WriteLog($"Startdate successfully set to: {startDateTime.ToShortDateString()}");
                        startDate = startDateTime;

                    }
                    else
                    {
                        // Handle cases where the cell does not contain a valid date
                        WriteLog("Unable to convert Start date");
                    }

                    object cellValue2 = worksheet.Range["P" + (rowNumber + 2)].Value;
                    if (cellValue != null && DateTime.TryParse(cellValue2.ToString(), out DateTime endDateTime))
                    {
                        // Successfully parsed the date

                        WriteLog($"Enddate successfully set to: {endDateTime.ToShortDateString()}");
                        endDate = endDateTime;
                    }
                    else
                    {
                        // Handle cases where the cell does not contain a valid date
                        WriteLog("Unable to convert End date");
                    }



                    // Insert into SQL using Dapper
                    using (SqlConnection connection = new SqlConnection("Data Source=ciisql10;Database=BAT_App;User Id=sa;Password='*id10t*';TrustServerCertificate=True;"))
                    {

                        // Query to get the current maximum CMA_Sequence for this customer
                        int nextSequence = connection.QueryFirstOrDefault<int>(
                            "SELECT ISNULL(MAX(CMA_Sequence), 0) + 1 FROM Chap_CmaItems WHERE cust_num = @CustNum",
                            new { CustNum = customerNumber }
                        );


                        int currentRow = rowNumber + 5; // Start from 5 rows below the "Corporate Customer Name" row

                        while (!string.IsNullOrEmpty(worksheet.Range["B" + currentRow].Text))
                        {
                            // Extract Detail Information
                            string item = worksheet.Range["B" + currentRow].Text;            // Item from column B
                            string description = worksheet.Range["A" + currentRow].DisplayText;     // Description from column A
                            //string sellPriceText = worksheet.Range["D" + currentRow].Text;   // SellPrice from column D
                            var proposedPriceValue = worksheet.Range[$"D{currentRow}"].Value;

                            // Parse SellPrice
                            decimal sellPrice = 0;
                            if (proposedPriceValue != null && decimal.TryParse(proposedPriceValue.ToString(), out sellPrice))
                            {
                                //WriteLog($"Price is: {proposedPriceValue} to {sellPrice}");
                            }
                            else
                            {
                                sellPrice = 0;
                                //throw new Exception($"Unable to parse SellPrice on row {currentRow}: '{proposedPriceValue}'");
                            }

                            // Perform the combined insert
                            string combinedQuery = @"
            INSERT INTO Chap_CmaItems 
            (Cust_name, Cust_num, CMA_Sequence, BuyingGroup, StartDate, EndDate, SubmittedBy, Site, Corp_flag, CmaFilename, Status, Item, Description, SellPrice, PromoTermsText, PromoFreightTermsText, PromoFreightMinimumText, PcfTypeText, PromoFreightMinimumsOtherText)
            VALUES 
            (@Cust_name, @Cust_num, @CMA_Sequence, @BuyingGroup, @StartDate, @EndDate, @SubmittedBy, @Site, @Corp_flag, @CmaFilename, @Status, @Item, @Description, @SellPrice, @PromoTermsText, @PromoFreightTermsText, @PromoFreightMinimumsText, @PcfTypeText, @PromoFreightMinimumsText)";

                            connection.Execute(combinedQuery, new
                            {
                                Cust_name = customerName,
                                Cust_num = customerNumber,
                                CMA_Sequence = nextSequence,
                                BuyingGroup = buyingGroup, // Assuming you have this value pre-defined
                                StartDate = startDate,
                                EndDate = endDate.Date,
                                SubmittedBy = SubmittedBy,
                                Site = site,
                                Corp_flag = corpFlag,
                                CmaFilename = cmaFilename,
                                Status = status,
                                Item = item,
                                Description = description,
                                SellPrice = sellPrice,
                                PromoTermsText = promoTermsText,
                                PromoFreightTermsText = promoFreightTermsText,
                                PromoFreightMinimumsText = promoFreightMinimumsText,
                                PromoFreightMinimumsOtherText = promoFreightTermsOtherAmtText,
                                PcfTypeText = PcfTypeText
                            });

                            // Move to the next row
                            currentRow++;
                        }








                    }
                }
            }

            WriteLog($"File processed successfully: {filePath}");
        }

        private void WriteLog(string message)
        {
            string logPath = @"C:\CMAInbound\CmaFileWatcherService.txt";
            using (StreamWriter writer = new StreamWriter(logPath, true))
            {
                writer.WriteLine($"{DateTime.Now}: {message}");
            }
        }
    }
}
