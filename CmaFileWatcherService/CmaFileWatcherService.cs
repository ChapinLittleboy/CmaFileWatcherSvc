using Dapper;
using IniParser;
using IniParser.Model;
using Microsoft.Data.SqlClient;
using Syncfusion.XlsIO;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.ServiceProcess;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace CmaFileWatcherService
{
    /*
    cd \source\repos\CmaFileWatcherSvc\CmaFileWatcherService\bin\Debug
    sc stop CmaFileWatcherService
    sc delete CmaFileWatcherService
    sc create CmaFileWatcherService binPath= "D:\source\repos\CmaFileWatcherSvc\CmaFileWatcherService\bin\Debug\CmaFileWatcherService.exe" obj= "NT AUTHORITY\NetworkService" type= own start= auto
    sc start CmaFileWatcherService

    sc query CmaFileWatcherService
    */

    public partial class CmaFileWatcherService : ServiceBase
    {
        private FileSystemWatcher _fileWatcher;
        private string _folderPath;
        private string CustNum; // used for renaming
        private string CorpNum; // used for renaming
        private string archiveName;
        private string _pcfDatabase;
        private string baseFilename;
        private string SenderEmail;
        private string GenNotes;
        private int RowWithPcfTypeHeading = 1;
        private bool _debugMode;

        private ConcurrentQueue<string> _fileQueue;
        private SemaphoreSlim _semaphore;
        private bool _isProcessing;

        public CmaFileWatcherService()
        {
            InitializeComponent();
            LoadConfiguration();
            InitializeWatcher();
            _fileQueue = new ConcurrentQueue<string>();
            _semaphore = new SemaphoreSlim(1, 1);
            _isProcessing = false;
            _debugMode = Environment.GetEnvironmentVariable("DebugCMAWatcher")?.ToLower() == "true";
        }

        private void LoadConfiguration()
        {
            var parser = new FileIniDataParser();
            IniData data = parser.ReadFile(@"\\ciiedi01\SendDocs\CMAInbound\config.ini");
            _folderPath = @"\\ciiedi01\SendDocs\CMAInbound";
            _pcfDatabase = data["Settings"]["PcfDatabase"];
            WriteLog($"pcfdatabase = {_pcfDatabase}");
            WriteDebugLog($"pcfdatabase = {_pcfDatabase}");
            WriteDebugLog("LoadConfiguration completed.");
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
            WriteDebugLog("InitializeWatcher completed.");
        }

        protected override void OnStart(string[] args)
        {
            _fileWatcher.EnableRaisingEvents = true;
            WriteLog("Service started and folder watching enabled.");
            WriteDebugLog("OnStart completed.");
        }

        protected override void OnStop()
        {
            _fileWatcher.EnableRaisingEvents = false;
            WriteLog("Service stopped.");
            WriteDebugLog("OnStop completed.");
        }

        private void OnNewExcelFile(object sender, FileSystemEventArgs e)
        {
            Task.Run(async () =>
            {
                try
                {
                    bool fileAccessible = await WaitForFileAccess(e.FullPath);
                    if (fileAccessible)
                    {
                        _fileQueue.Enqueue(e.FullPath);
                        WriteLog($"New file detected and enqueued: {e.FullPath}");
                        WriteDebugLog($"OnNewExcelFile: File enqueued at {e.FullPath}");

                        if (!_isProcessing)
                        {
                            _isProcessing = true;
                            ProcessQueue();
                        }
                    }
                    else
                    {
                        WriteLog($"File not accessible after multiple attempts: {e.FullPath}");
                        WriteDebugLog($"OnNewExcelFile: File not accessible after multiple attempts: {e.FullPath}");
                    }
                }
                catch (Exception ex)
                {
                    WriteLog($"Unhandled exception in OnNewExcelFile: {ex.Message}");
                    WriteDebugLog($"Unhandled exception in OnNewExcelFile: {ex.Message}");
                }
            });
        }

        private async Task<bool> WaitForFileAccess(string filePath, int maxRetries = 5, int delayMilliseconds = 1000)
        {
            for (int i = 0; i < maxRetries; i++)
            {
                try
                {
                    using (FileStream stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.None))
                    {
                        if (stream.Length > 0)
                        {
                            return true;
                        }
                    }
                }
                catch (IOException)
                {
                    await Task.Delay(delayMilliseconds);
                }
            }
            return false;
        }

        private async void ProcessQueue()
        {
            while (_fileQueue.TryDequeue(out string filePath))
            {
                await _semaphore.WaitAsync();
                try
                {
                    WriteLog($"Processing file: {filePath}");
                    WriteDebugLog($"ProcessQueue: Processing file at {filePath}");
                    await ProcessExcelFile(filePath);
                }
                catch (Exception ex)
                {
                    WriteLog($"Error processing file {filePath}: {ex.Message}");
                    WriteDebugLog($"ProcessQueue: Error processing file {filePath}: {ex.Message}");
                }
                finally
                {
                    _semaphore.Release();
                }
            }
            _isProcessing = false;
        }






        private void OnNewExcelFilexx(object sender, FileSystemEventArgs e)
        {
            try
            {
                System.Threading.Thread.Sleep(5000);
                WriteLog($"New file detected: {e.FullPath}");
                WriteDebugLog($"OnNewExcelFile: File detected at {e.FullPath}");
                ProcessExcelFile(e.FullPath);
            }
            catch (Exception ex)
            {
                WriteLog($"Error processing file {e.FullPath}: {ex.Message}");
                WriteDebugLog($"OnNewExcelFile: Error processing file {e.FullPath}: {ex.Message}");
            }
        }

        private async Task ProcessExcelFile(string filePath)
        {
            bool cmaIsValid = false;
            bool isReplacementCMA = false;
            string existingPCF = string.Empty;
            baseFilename = Path.GetFileName(GetBaseFilename(filePath));
            SenderEmail = GetSenderEmail(filePath);
            WriteLog(baseFilename);
            WriteDebugLog($"ProcessExcelFile: baseFilename = {baseFilename}, SenderEmail = {SenderEmail}");
            try
            {
                using (var inputStream = new FileStream(filePath, FileMode.Open))
                {
                    using (ExcelEngine excelEngine = new ExcelEngine())
                    {
                        IApplication application = excelEngine.Excel;
                        IWorkbook workbook = application.Workbooks.Open(inputStream);
                        IWorksheet worksheet = workbook.Worksheets.FirstOrDefault(ws =>
                                                   ws.Name.Equals("CMA Template", StringComparison.OrdinalIgnoreCase))
                                               ?? workbook.Worksheets.FirstOrDefault(ws =>
                                                   ws.Name.Equals("CMA", StringComparison.OrdinalIgnoreCase))
                                               ?? workbook.Worksheets.FirstOrDefault();

                        if (worksheet == null)
                        {
                            WriteLog("No worksheets available in the workbook.");
                            WriteDebugLog("ProcessExcelFile: No worksheets available in the workbook.");
                            return;
                        }

                        WriteLog($"Worksheet selected: {worksheet.Name}");
                        WriteDebugLog($"ProcessExcelFile: Worksheet selected: {worksheet.Name}");

                        existingPCF = worksheet.Range["A2"].DisplayText;
                        isReplacementCMA = int.TryParse(existingPCF, out int existingPCFint);
                        WriteDebugLog(
                            $"ProcessExcelFile: existingPCF = {existingPCF}, isReplacementCMA = {isReplacementCMA}");

                        int rowNumber = 1;
                        while (true)
                        {
                            var cellValueP = worksheet.Range["P" + rowNumber].DisplayText;
                            if (DateTime.TryParse(cellValueP, out DateTime _))
                            {
                                break;
                            }

                            rowNumber++;
                            if (rowNumber > 10)
                            {
                                break;
                            }
                        }

                        WriteLog($"Firstrow set to: {rowNumber}");
                        WriteDebugLog($"ProcessExcelFile: Firstrow set to: {rowNumber}");

                        RowWithPcfTypeHeading = 1;

                        while (true)
                        {
                            var teststring = worksheet.Range["B" + RowWithPcfTypeHeading].DisplayText;
                            if (string.Equals(teststring, "PCF Type", StringComparison.OrdinalIgnoreCase))
                            {
                                break;
                            }

                            RowWithPcfTypeHeading++;
                            if (RowWithPcfTypeHeading > 10)
                            {
                                break;
                            }
                        }

                        WriteDebugLog($"ProcessExcelFile: RowWithPcfTypeHeading set to: {RowWithPcfTypeHeading}");

                        string promoTermsText = worksheet.Range["F" + (rowNumber != 1 ? 5 : 2)]?.DisplayText ?? "";
                        WriteDebugLog(
                            $"PromoTermsText extracted from Cell F{(rowNumber != 1 ? 5 : 2)}: {promoTermsText}");

                        string promoFreightTermsText =
                            worksheet.Range["C" + (rowNumber != 1 ? 5 : 2)]?.DisplayText ?? "";
                        WriteDebugLog(
                            $"PromoFreightTermsText extracted from Cell C{(rowNumber != 1 ? 5 : 2)}: {promoFreightTermsText}");

                        string promoFreightMinimumsText =
                            worksheet.Range["D" + (rowNumber != 1 ? 5 : 2)]?.DisplayText ?? "";
                        WriteDebugLog(
                            $"promoFreightMinimumsText extracted from Cell D{(rowNumber != 1 ? 5 : 2)}: {promoFreightMinimumsText}");

                        string PcfTypeText = worksheet.Range["B" + (rowNumber != 1 ? 5 : 2)]?.DisplayText ?? "";
                        WriteDebugLog($"PcfTypeText extracted from Cell B{(rowNumber != 1 ? 5 : 2)}: {PcfTypeText}");

                        string promoFreightTermsOtherAmtText =
                            worksheet.Range["D" + (rowNumber != 1 ? 6 : 3)]?.DisplayText ?? "";
                        WriteDebugLog(
                            $"promoFreightTermsOtherAmtText extracted from Cell D{(rowNumber != 1 ? 5 : 2)}: {promoFreightTermsOtherAmtText}");

                        string corpNumber = string.Empty;
                        string customerNumber = string.Empty;
                        string customerName = string.Empty;



                        if (worksheet.Range["A10"].DisplayText == "Sales Manager")
                        {
                            WriteDebugLog($"A10 has --{worksheet.Range["D10"].DisplayText}--");
                            customerNumber = worksheet.Range["B" + (rowNumber + 2 - 1)].DisplayText;
                            WriteDebugLog($"customerNumber extracted from Cell B{rowNumber + 2 - 1}: {customerNumber}");
                            corpNumber = "";
                            WriteDebugLog($"corpNumber and CorpNum Not used in this CMA");
                            customerName = worksheet.Range["B8"].DisplayText;
                            WriteDebugLog($"customerName extracted from Cell B{rowNumber}: {customerName}");
                        }
                        else
                        {
                            WriteDebugLog($"A10x has --{worksheet.Range["D10"].DisplayText}--");
                            customerNumber = worksheet.Range["B" + (rowNumber + 2)].DisplayText;
                            WriteDebugLog($"customerNumber extracted from Cell B{rowNumber + 2}: {customerNumber}");
                            corpNumber = worksheet.Range["B" + (rowNumber + 1)].DisplayText;
                            WriteDebugLog($"corpNumber and CorpNum extracted from Cell B{rowNumber + 1}: {corpNumber}");
                            customerName = worksheet.Range["B" + rowNumber].DisplayText;
                            WriteDebugLog($"customerName extracted from Cell B{rowNumber}: {customerName}");
                        }

                        CustNum = customerNumber;
                        WriteDebugLog($"CustNum set to {CustNum}");



                        CorpNum = corpNumber;


                        customerNumber = !string.IsNullOrEmpty(customerNumber) ? customerNumber : corpNumber;
                        WriteDebugLog($"customerNumber [cust vs corp] now set to : {customerNumber}");

                        int corpFlag = !string.IsNullOrEmpty(corpNumber) ? 1 : 0;
                        WriteDebugLog($"corpFlag set to : {corpFlag}");

                        string buyingGroup = worksheet.Range["K" + (rowNumber + 2)].DisplayText;
                        WriteDebugLog($"buyingGroup extracted from Cell K{rowNumber + 2}: {buyingGroup}");
                        if (string.Equals(buyingGroup, "#N/A", StringComparison.OrdinalIgnoreCase))
                        {
                            buyingGroup = "";
                            WriteDebugLog("Buying Group is #N/A, setting it to empty.");

                        }

                        string SubmittedBy = worksheet.Range["M" + (rowNumber + 2)].DisplayText;
                        WriteDebugLog($"SubmittedBy extracted from Cell M{rowNumber + 2}: {SubmittedBy}");

                        DateTime startDate = DateTime.Now;
                        DateTime endDate = DateTime.Now;

                        WriteDebugLog(
                            $"ProcessExcelFile: customerNumber = {customerNumber}, corpNumber = {corpNumber}, customerName = {customerName}");

                        if (rowNumber != 1)
                        {
                            if (worksheet.Range["A1"].DisplayText == "General Notes")
                            {
                                GenNotes = worksheet.Range["B1"].DisplayText;
                            }
                        }

                        string site = "BAT";
                        string cmaFilename = Path.GetFileName(filePath);
                        string dateTimeStamp = DateTime.Now.ToString("yyyyMMddHHmmss");
                        string prefix = string.IsNullOrEmpty(corpNumber) ? customerNumber : corpNumber;
                        archiveName = $"{prefix}_{dateTimeStamp}_{baseFilename}";
                        string status = "N";

                        object cellValue = worksheet.Range["P" + rowNumber].Value;
                        if (cellValue != null && DateTime.TryParse(cellValue.ToString(), out DateTime startDateTime))
                        {
                            startDate = startDateTime;
                        }
                        else
                        {
                            WriteLog($"Unable to convert Start date {cellValue}");
                        }

                        object cellValue2 = worksheet.Range["P" + (rowNumber + 2)].Value;
                        if (cellValue2 != null && DateTime.TryParse(cellValue2.ToString(), out DateTime endDateTime))
                        {
                            endDate = endDateTime;
                        }
                        else
                        {
                            WriteLog($"Unable to convert End date {cellValue2}");
                        }

                        WriteDebugLog($"ProcessExcelFile: startDate = {startDate}, endDate = {endDate}");

                        using (SqlConnection connection = new SqlConnection(
                                   "Data Source=ciisql10;Database=BAT_App;User Id=sa;Password='*id10t*';TrustServerCertificate=True;"))
                        {
                            int nextSequence = connection.QueryFirstOrDefault<int>(
                                "SELECT ISNULL(MAX(CMA_Sequence), 0) + 1 FROM Chap_CmaItems WHERE cust_num = @CustNum",
                                new { CustNum = customerNumber }
                            );

                            int currentRow = rowNumber + 5;
                            WriteDebugLog($"Looking for first item in B{currentRow}");
                            while (!string.IsNullOrEmpty(worksheet.Range["B" + currentRow].DisplayText))
                            {
                                string item = worksheet.Range["B" + currentRow].DisplayText;
                                string description = worksheet.Range["A" + currentRow].DisplayText;
                                var proposedPriceValue = worksheet.Range[$"D{currentRow}"].Value;
                                decimal sellPrice = 0;
                                if (proposedPriceValue != null &&
                                    decimal.TryParse(proposedPriceValue.ToString(), out sellPrice))
                                {
                                    // Successfully parsed the price
                                }
                                else
                                {
                                    sellPrice = 0;
                                }

                                int? replacesValue = (existingPCFint > 0) ? existingPCFint : (int?)null;
                                string combinedQuery = @"
                                INSERT INTO Chap_CmaItems 
                                (Cust_name, Cust_num, CMA_Sequence, BuyingGroup, StartDate, EndDate, SubmittedBy, Site, 
                                Corp_flag, CmaFilename, Status, Item, Description, SellPrice, PromoTermsText, PromoFreightTermsText, 
                                PromoFreightMinimumsText, PcfTypeText, PromoFreightMinimumsOtherText, SenderEmail, ReplacesPCF, GenNotes)
                                VALUES 
                                (@Cust_name, @Cust_num, @CMA_Sequence, @BuyingGroup, @StartDate, @EndDate, @SubmittedBy, @Site, @Corp_flag, @CmaFilename, @Status, @Item,
                                @Description, @SellPrice, @PromoTermsText, @PromoFreightTermsText, @PromoFreightMinimumsText, @PcfTypeText, @PromoFreightMinimumsOtherText, @SenderEmail,
                                @ReplacesPCF, @GenNotes)";
                                connection.Execute(combinedQuery, new
                                {
                                    Cust_name = customerName,
                                    Cust_num = customerNumber,
                                    CMA_Sequence = nextSequence,
                                    BuyingGroup = buyingGroup,
                                    StartDate = startDate,
                                    EndDate = endDate.Date,
                                    SubmittedBy = SubmittedBy,
                                    Site = site,
                                    Corp_flag = corpFlag,
                                    CmaFilename = archiveName,
                                    Status = status,
                                    Item = item,
                                    Description = description,
                                    SellPrice = sellPrice,
                                    PromoTermsText = promoTermsText,
                                    PromoFreightTermsText = promoFreightTermsText,
                                    PromoFreightMinimumsText = promoFreightMinimumsText,
                                    PromoFreightMinimumsOtherText = promoFreightTermsOtherAmtText,
                                    PcfTypeText = PcfTypeText,
                                    SenderEmail = SenderEmail,
                                    ReplacesPCF = replacesValue,
                                    GenNotes = GenNotes
                                });
                                WriteDebugLog($"cmaitems add Query: {combinedQuery}");
                                currentRow++;
                            }

                            var validator = new CMAValidator(connection, _pcfDatabase, archiveName);
                            var validationErrors = validator.ValidateCMARecords();
                            WriteLog($"Number of validation errors: {validationErrors.Count}");
                            WriteDebugLog($"ProcessExcelFile: Number of validation errors: {validationErrors.Count}");

                            string numRecsQuery = @"
                            SELECT COUNT(*) FROM Chap_CmaItems WHERE CmaFilename = @CmaFilename AND PCFNumber IS NULL";
                            int numRecords =
                                connection.QueryFirstOrDefault<int>(numRecsQuery, new { CmaFilename = archiveName });

                            WriteLog($"Number of records in Chap_CmaItems: {numRecords}");
                            WriteDebugLog($"ProcessExcelFile: Number of records in Chap_CmaItems: {numRecords}");


                            if (validationErrors.Count > 0)
                            {
                                WriteLog("Validation failed:");
                                foreach (var error in validationErrors)
                                {
                                    WriteLog(error);
                                }

                                WriteLog($"Failure {SenderEmail}, {baseFilename}, {validationErrors.Count}");
                                SendValidationFailureEmail(SenderEmail, baseFilename, validationErrors);

                                //instead of deleting, change status to FAILURE
                                string deletesql = @"
                                Update Chap_CmaItems
                                SET Status = 'F'
                                WHERE CmaFilename = @CmaFilename
                                AND PcfNumber IS NULL
                                AND Status = 'N'";
                                connection.Execute(deletesql, new { CmaFilename = archiveName });
                            }
                            else
                            {
                                WriteLog("Validation passed");
                                // Introduce a 10-second delay
                                await Task.Delay(10000);

                                if (existingPCFint > 0)
                                {
                                    var parametersR = new { OriginalPcfNum = existingPCFint };
                                    connection.Execute("sp_ArchiveReplacedPCF", parametersR,
                                        commandType: System.Data.CommandType.StoredProcedure);

                                    var parametersN = new { CmaName = archiveName, PCFNum = existingPCFint };
                                    connection.Execute("CreatePcfFromChapCmaItems_sp", parametersN,
                                        commandType: System.Data.CommandType.StoredProcedure);
                                }
                                else
                                {
                                    var parameters = new { CmaName = archiveName, PCFNum = 0 };
                                    connection.Execute("CreatePcfFromChapCmaItems_sp", parameters,
                                        commandType: System.Data.CommandType.StoredProcedure);
                                }
                            }

                            var parameters2 = new { CmaName = archiveName };
                            var pcfNumber = connection.QuerySingleOrDefault<string>(
                                "SELECT TOP 1 PCFNumber FROM Chap_CmaItems WHERE cmaFileName = @CmaName AND PCFNumber IS NOT NULL",
                                parameters2);

                            if (!string.IsNullOrEmpty(pcfNumber))
                            {
                                WriteLog($"PCF {pcfNumber} created for CMA {cmaFilename} {archiveName}");
                                SendSuccessfulPCFCreationEmail(SenderEmail, baseFilename, pcfNumber);
                                worksheet.Range["A2"].Text = pcfNumber;

                                string archiveFolderPath = Path.Combine(_folderPath, "Processed");
                                if (!Directory.Exists(archiveFolderPath))
                                {
                                    Directory.CreateDirectory(archiveFolderPath);
                                }

                                string newFilePath = Path.Combine(archiveFolderPath, archiveName);
                                using (FileStream outputStream = new FileStream(newFilePath, FileMode.Create,
                                           FileAccess.Write, FileShare.None))
                                {
                                    workbook.SaveAs(outputStream);
                                }
                            }
                            else
                            {
                                WriteLog($"PCF not created for CMA {cmaFilename} {archiveName}");
                                string archiveFolderPath = Path.Combine(_folderPath, "Rejected");
                                if (!Directory.Exists(archiveFolderPath))
                                {
                                    Directory.CreateDirectory(archiveFolderPath);
                                }

                                string newFilePath = Path.Combine(archiveFolderPath, archiveName);
                                using (FileStream outputStream = new FileStream(newFilePath, FileMode.Create,
                                           FileAccess.Write, FileShare.None))
                                {
                                    workbook.SaveAs(outputStream);
                                }
                            }


                        }

                        WriteLog($"File processed successfully: {filePath}");
                        WriteDebugLog($"ProcessExcelFile: File processed successfully: {filePath}");
                        // Ensure the workbook and inputStream are properly disposed of
                        workbook.Close();
                        excelEngine.Dispose();
                        inputStream.Close();
                        inputStream.Dispose();

                        string processedFolder = Path.Combine(_folderPath, "Processed");
                        string rejectedFolder = Path.Combine(_folderPath, "Rejected");
                        string processedFilePath = Path.Combine(processedFolder, archiveName);
                        string rejectedFilePath = Path.Combine(rejectedFolder, archiveName);

                        try
                        {


                            if (File.Exists(processedFilePath) || File.Exists(rejectedFilePath))
                            {
                                if (File.Exists(filePath))
                                {
                                    File.Delete(filePath);
                                    Console.WriteLine($"Original file deleted: {filePath}");
                                }
                                else
                                {
                                    Console.WriteLine("Original file does not exist, skipping deletion.");
                                }
                            }
                            else
                            {
                                Console.WriteLine(
                                    "File was not found in Processed or Rejected folders. Skipping deletion.");
                            }
                        }
                        catch (Exception ex)
                        {
                            WriteLog($"Error deleting original file: {ex.Message}");
                            WriteDebugLog($"ProcessExcelFile: Error deleting original file: {ex.Message}");
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                WriteLog($"Unhandled exception in ProcessExcelFile: {ex}");
                WriteDebugLog($"Unhandled exception in ProcessExcelFile: {ex}");
                throw; // Re-throw the exception to ensure it is logged in the calling method
            }
        }

        private void WriteLog(string message)
        {
            string logPath = @"\\ciiedi01\SendDocs\CMAInbound\CmaFileWatcherService.log";
            using (StreamWriter writer = new StreamWriter(logPath, true))
            {
                writer.WriteLine($"{DateTime.Now}: {message}");
            }
        }

        private void WriteDebugLog(string message)
        {
            if (_debugMode)
            {
                string debugLogPath = @"\\ciiedi01\SendDocs\CMAInbound\debug.log";
                using (StreamWriter writer = new StreamWriter(debugLogPath, true))
                {
                    writer.WriteLine($"{DateTime.Now}: {message}");
                }
            }
        }

        // Other methods (SendValidationFailureEmail, SendSuccessfulPCFCreationEmail, etc.) remain unchanged

        public void SendValidationFailureEmail(string senderEmail, string baseFileName, List<string> validationErrors)
        {
            // if (validationErrors == null || validationErrors.Count == 0)
            //     return; // No errors, no need to send an email

            // Construct email message
            string emailSubject = $"CMA Validation Failed: {baseFileName}";
            string emailBody = $@"
Dear {senderEmail},

Your CMA submission '{baseFileName}' failed validation due to the following issues:

{string.Join("\n\n", validationErrors)}

Please review the errors, make the necessary corrections, and resubmit the file.

If you have any questions, please contact the SalesOps team.

Best regards,  
Sales Operations Team";


            using (SqlConnection _dbConnection =
                   new SqlConnection(
                       "Data Source=ciisql10;Database=BAT_App;User Id=sa;Password='*id10t*';TrustServerCertificate=True;"))
            {

                // Send email using SQL Server Database Mail (dbMail)

                try
                {
                    _dbConnection.Execute(@"
        EXEC msdb.dbo.sp_send_dbmail
            @profile_name = 'SalesOps',
            @recipients = @RecipientEmail,
            @subject = @EmailSubject,
            @body = @EmailBody",
                        new
                        {
                            RecipientEmail = senderEmail,
                            EmailSubject = emailSubject,
                            EmailBody = emailBody
                        });

                    WriteLog($"Email successfully sent to {senderEmail} for {baseFileName}.");
                }
                catch (Exception ex)
                {
                    WriteLog($"Error sending email: {ex.Message}");
                }
            }
        }

        public void SendValidationFailureEmailHtml(string senderEmail, string baseFileName, List<string> validationErrors)
        {
            if (validationErrors == null || validationErrors.Count == 0)
                return; // No errors, no need to send an email

            // Construct email message with line breaks
            string emailSubject = $"CMA Validation Failed: {baseFileName}";

            string emailBody = $@"
Dear {senderEmail},<br><br>

Your CMA submission '<b>{baseFileName}</b>' failed validation due to the following issues:<br><br>

{string.Join("<br>", validationErrors)}<br><br>

Please review the errors, make the necessary corrections, and resubmit the file.<br><br>

If you have any questions, please contact the SalesOps team.<br><br>

Best regards,<br>
Sales Operations Team";


            using (SqlConnection _dbConnection =
                   new SqlConnection(
                       "Data Source=ciisql10;Database=BAT_App;User Id=sa;Password='*id10t*';TrustServerCertificate=True;"))
            {
                // Send email using SQL Server Database Mail (dbMail)
                _dbConnection.Execute(@"
        EXEC msdb.dbo.sp_send_dbmail
            @profile_name = 'SalesOps',
            @recipients = @RecipientEmail,
            @subject = @EmailSubject,
            @body = @EmailBody",
                    new
                    {
                        RecipientEmail = senderEmail,
                        EmailSubject = emailSubject,
                        EmailBody = emailBody
                    });
            }
        }








        public void SendSuccessfulPCFCreationEmail(string senderEmail, string baseFileName, string pcfNumber)
        {
            // if (validationErrors == null || validationErrors.Count == 0)
            //     return; // No errors, no need to send an email

            // Construct email message
            string emailSubject = $"CMA {baseFileName}    PCF {pcfNumber} created";
            string emailBody = $@"
Dear {senderEmail},

Your CMA submission with file '{baseFileName}' was successfully processed and a PCF was created. You can view the PCF on the www.ChapinPcfManager.com website for review.



If you have any questions, please contact the SalesOps team.

Best regards,  
Sales Operations Team";


            using (SqlConnection _dbConnection =
                   new SqlConnection(
                       "Data Source=ciisql10;Database=BAT_App;User Id=sa;Password='*id10t*';TrustServerCertificate=True;"))
            {

                // Send email using SQL Server Database Mail (dbMail)

                try
                {
                    _dbConnection.Execute(@"
        EXEC msdb.dbo.sp_send_dbmail
            @profile_name = 'SalesOps',
            @recipients = @RecipientEmail,
            @subject = @EmailSubject,
            @body = @EmailBody",
                        new
                        {
                            RecipientEmail = senderEmail,
                            EmailSubject = emailSubject,
                            EmailBody = emailBody
                        });

                    WriteLog($"Email successfully sent to {senderEmail} for {baseFileName}.");
                }
                catch (Exception ex)
                {
                    WriteLog($"Error sending email: {ex.Message}");
                }
            }
        }
        static string GetSenderEmail(string filename)
        {
            // Use regex to extract the part after "_sentby_" and before the extension
            //var match = Regex.Match(filename, @"_sentby_([\w\.-]+@[a-zA-Z0-9\.-]+\.[a-zA-Z]{2,})");
            //var match = Regex.Match(filename, @"_sentby_([\w\.-]+@[a-zA-Z0-9\.-]+\.[a-zA-Z]{2,})\b");
            var match = Regex.Match(filename, @"_sentby_([\w\.-]+@chapinmfg\.com)", RegexOptions.IgnoreCase);

            return match.Success ? match.Groups[1].Value : "";
        }
        static string GetBaseFilename(string filename)
        {
            // Find position of "_sentby_"
            int index = filename.IndexOf("_sentby_");
            if (index != -1)
            {
                return filename.Substring(0, index) + Path.GetExtension(filename);
            }
            return filename; // Return original if "_sentby_" is not found
        }

    }
}
