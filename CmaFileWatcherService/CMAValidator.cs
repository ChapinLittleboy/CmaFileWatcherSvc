namespace CmaFileWatcherService
{
    using Dapper;
    using System;
    using System.Collections.Generic;
    using System.Data;
    using System.Linq;

    public class CMAValidator
    {
        private readonly IDbConnection _dbConnection;
        private readonly string _pcfDatabase;
        private readonly string _linkedServer = "CiiSQL01";
        private readonly string _cmaFileName;

        public CMAValidator(IDbConnection dbConnection, string pcfDatabase, string cmaFileName)
        {
            _dbConnection = dbConnection;
            _pcfDatabase = pcfDatabase;
            _cmaFileName = cmaFileName;
        }

        public List<string> ValidateCMARecords()
        {
            var validationErrors = new List<string>();

            // Step 1: Get newly added records
            var cmaRecords = _dbConnection.Query<dynamic>(@"
           SELECT * FROM Chap_CmaItems WHERE CmaFilename = @CmaFilename",
                new { CmaFilename = _cmaFileName }).ToList();

            if (!cmaRecords.Any())
                return validationErrors; // No records found, nothing to validate


            bool isFirstRecord = true; // Flag to identify the first record in the set

            foreach (var record in cmaRecords)
            {
                string pcfTypeText = record.PcfTypeText ?? "";
                string promoTermsText = record.PromoTermsText ?? "";
                string promoFreightTermsText = record.PromoFreightTermsText ?? "";
                string promoFreightMinimumsText = record.PromoFreightMinimumsText ?? "";
                string promoFreightMinimumsOtherText = record.PromoFreightMinimumsOtherText ?? "";

                DateTime? startDate = record.StartDate;
                DateTime? endDate = record.EndDate;



                // **Run Steps 2 & 3 only for the first record**
                if (isFirstRecord)
                {

                    // Step 2: Check PcfTypeText rules
                    if (string.IsNullOrWhiteSpace(pcfTypeText))
                    {
                        validationErrors.Add($"Pcf Type is required");
                    }

                    if (pcfTypeText.StartsWith("PD") || pcfTypeText.StartsWith("PW"))
                    {
                        if (string.IsNullOrWhiteSpace(promoTermsText))
                            validationErrors.Add(
                                $"Promo Termsis required when Pcf Type is 'PD' or 'PW'.");
                        if (string.IsNullOrWhiteSpace(promoFreightTermsText))
                            validationErrors.Add(
                                $"Promo Freight Terms is required when  Pcf Type is 'PD' or 'PW'.");
                        if (string.IsNullOrWhiteSpace(promoFreightMinimumsText))
                            validationErrors.Add(
                                $"Promo Freight Minimums is required when Pcf Type is 'PD' or 'PW'.");
                        //if (string.IsNullOrWhiteSpace(promoFreightMinimumsOtherText))
                        //   validationErrors.Add(
                        //       $"CMA ID {_cmaFileName}: PromoFreightMinimumsOtherText is required when Pcf Type is 'PD' or 'PW'.");
                    }

                    // Step 3: Validate Dates
                    if (!startDate.HasValue)
                        validationErrors.Add($"StartDate is required.");
                    if (!endDate.HasValue)
                        validationErrors.Add($"EndDate is required.");
                    if (startDate.HasValue && endDate.HasValue && startDate > endDate)
                        validationErrors.Add($"StartDate cannot be after EndDate.");
                }


                // Step 4: Check ProgControl conflicts
                var progControlConflicts = _dbConnection.Query<int>($@"
                SELECT COUNT(*) FROM [{_linkedServer}].[{_pcfDatabase}].dbo.ProgControl pc
                JOIN [{_linkedServer}].[{_pcfDatabase}].dbo.pcitems pi ON CAST(pc.PcfNum AS NVARCHAR) = pi.PcfNumber
                WHERE pc.Custnum = TRIM(@Cust_num) AND pc.ProgSDate = @StartDate
                AND pi.ItemNum = @Item", new { record.Cust_num, StartDate = record.StartDate, record.Item }).FirstOrDefault();

                if (progControlConflicts > 0)
                {
                    validationErrors.Add($"Date Conflict: Existing PCF record for Customer {record.Cust_num} and Item {record.Item} has same start date. You must use a different date.");
                }

                // Step 5: Check Items in item_mst
                var itemCheck = _dbConnection.Query<dynamic>(@"
                SELECT stat FROM item_mst WHERE item = @Item", new { record.Item }).FirstOrDefault();

                if (itemCheck == null)
                {
                    validationErrors.Add($"Item {record.Item} does not exist in Syteline. You must remove this item from the CMA");
                }
                else if (itemCheck.stat == "O")
                {
                    validationErrors.Add($"Item {record.Item} is obsolete in Syteline. You must remove this item from the CMA");
                }
            }

            return validationErrors;
        }



    }

}
