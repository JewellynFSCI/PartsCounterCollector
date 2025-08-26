using System;
using System.Globalization;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using MySql.Data.MySqlClient;
using static PartsCounter.Model.Models;
using Microsoft.Extensions.Configuration;
using OfficeOpenXml;
using PartsCounter.Model;
using Dapper;
using System.Data;
using System.Dynamic;


namespace PartsCounter
{
    class Program
    {
        static string? connectionString;
        static string? logSource;
        static string? logError;
        static string? logArchive;
        static string logFilePath;

        static void Main()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            if (!InitializeConfiguration())
                return;

            ProcessFile();
        }

        #region InitializeConfiguration
        private static bool InitializeConfiguration()
        {
            try
            {
                var configuration = new ConfigurationBuilder()
                    .SetBasePath(Directory.GetCurrentDirectory())
                    .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
                    .Build();

                connectionString = configuration.GetConnectionString("DefaultConnection");

                string settingsFilePath = "FileSetting.cn"; // relative or absolute path
                var settings = LoadSettings(settingsFilePath);

                logSource = settings["LogsSourcePath"];
                logError = settings["ErrorLogsPath"];
                logArchive = settings["ArchiveLogsPath"];


                // Validate folders
                if (!Directory.Exists(logSource))
                {
                    Console.WriteLine($"Source Directory not found: {logSource}");
                    string message = $"Source Directory not found: {logSource}";
                    SaveErrorLog(message);
                    return false;
                }

                if (!Directory.Exists(logError))
                {
                    Console.WriteLine($"Error Directory not found: {logError}. It will automatically create error folder.");
                    string message = $"Error Directory not found: {logError}. It will automatically create error folder.";
                    SaveErrorLog(message);
                    return false;
                }

                if (!Directory.Exists(logArchive))
                {
                    Console.WriteLine($"Archive Directory not found: {logArchive}");
                    string message = $"Archive Directory not found: {logArchive}";
                    SaveErrorLog(message);
                    return false;
                }
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading configuration: {ex.Message}");
                string message = $"Error loading configuration: {ex.Message}";
                SaveErrorLog(message);
                return false;
            }
        }
        #endregion

        #region LoadSettings
        static Dictionary<string, string> LoadSettings(string filePath)
        {
            var settings = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            foreach (var line in File.ReadAllLines(filePath))
            {
                if (string.IsNullOrWhiteSpace(line) || line.TrimStart().StartsWith("#"))
                    continue; // skip empty lines or comments

                var parts = line.Split('=', 2);
                if (parts.Length == 2)
                {
                    settings[parts[0].Trim()] = parts[1].Trim();
                }
            }

            return settings;
        }
        #endregion

        #region ProcessFile
        private static void ProcessFile()
        {
            var excelFiles = Directory.GetFiles(logSource, "*.xlsx");
            if (excelFiles.Length == 0)
            {
                Console.WriteLine($"No XLSX files found in: {logSource}");
                string message = $"No XLSX files found in: {logSource}";
                SaveErrorLog(message);
                return;
            }

            Console.WriteLine($"Collecting data is on-going.");

            var allSummaries = new List<Models.Summary>();
            var allBreakdowns = new List<Models.Breakdown>();

            string destArchiveFolder = ArchiveFolder();
            string destErrorFolder = ErrorFolder();

            # region Process each CSV file
            foreach (var file in excelFiles)
            {
                var fileName = Path.GetFileName(file); // Get only file name
                try
                {
                    int partsCounterNo = GetPartsCounterNoFromFile(fileName);

                    using (var package = new ExcelPackage(new FileInfo(file)))
                    {
                        var worksheet = package.Workbook.Worksheets.FirstOrDefault();
                        if (worksheet == null)
                        {
                            Console.WriteLine($"No worksheet found in file {file}");
                            string message = $"No worksheet found in file {file}";
                            SaveErrorLog(message);
                            MoveFileToErrorFolder(file, destErrorFolder);
                            continue;
                        }

                        #region Parse summary from second row (row 2)
                        // Get headers (row 1)
                        var headers = Enumerable.Range(1, worksheet.Dimension.End.Column)
                            .Select(c => worksheet.Cells[1, c].Text)
                            .ToArray();

                        // Build header map (case-insensitive)
                        var headerMap = headers
                            .Select((h, i) => new { h, i })
                            .ToDictionary(x => x.h.Trim(), x => x.i, StringComparer.OrdinalIgnoreCase);

                        // Get row 2 values
                        var summaryCols = Enumerable.Range(1, worksheet.Dimension.End.Column)
                            .Select(c => worksheet.Cells[2, c].Text)
                            .ToArray();

                        // Now call your parser
                        var summary = ParseSummary(summaryCols, headerMap, partsCounterNo);
                        allSummaries.Add(summary);
                        #endregion

                        #region Parse breakdown starting from fourth row (row 4)
                        // Build header map once (row 3 = headers)
                        var breakdownHeaders = Enumerable.Range(1, worksheet.Dimension.End.Column)
                            .Select(c => worksheet.Cells[3, c].Text)
                            .ToArray();

                        var breakdownHeaderMap = breakdownHeaders
                            .Select((h, i) => new { h = h.Trim(), i })
                            .Where(x => !string.IsNullOrWhiteSpace(x.h))        // remove empty headers
                            .GroupBy(x => x.h, StringComparer.OrdinalIgnoreCase) // group duplicates
                            .ToDictionary(g => g.Key, g => g.First().i, StringComparer.OrdinalIgnoreCase); // take first index


                        // Loop through data rows (row 4+)
                        for (int row = 4; row <= worksheet.Dimension.End.Row; row++)
                        {
                            var cols = Enumerable.Range(1, worksheet.Dimension.End.Column)
                                                 .Select(c => worksheet.Cells[row, c].Text)
                                                 .ToArray();

                            var breakdown = ParseBreakdown(cols, breakdownHeaderMap, partsCounterNo);
                            allBreakdowns.Add(breakdown);
                        }
                        #endregion

                        //Check id data exist in DB
                        bool checkDuplicate = CheckSummaryDuplicate(allSummaries, connectionString);
                        if (checkDuplicate)
                        {
                            //with exisiting data in db
                            Console.WriteLine($"This log already exists in database! ({fileName})");
                            string message = $"This log already exists in database! ({fileName})";
                            SaveErrorLog(message);
                            MoveFileToErrorFolder(file, destErrorFolder);
                            allSummaries.Clear();
                            allBreakdowns.Clear();
                            continue;
                        }

                        #region Save file data to DB
                        int IDSummary = SaveFileSummary(allSummaries, connectionString);
                        SaveFileBreakdown(allBreakdowns, connectionString, IDSummary);
                        #endregion

                        MoveFileToArchieveFolder(file, destArchiveFolder);
                        allSummaries.Clear();
                        allBreakdowns.Clear();
                        Console.WriteLine($"Success: {fileName}");
                    }
                }
                catch (Exception ex)
                {
                    allSummaries.Clear();
                    allBreakdowns.Clear();
                    MoveFileToErrorFolder(file, destErrorFolder);
                    Console.WriteLine($"Error processing file - {fileName}: {ex.Message}");
                    string message = $"Error processing file - {fileName}: {ex.Message}";
                    SaveErrorLog(message);
                    continue;
                }
            }
            #endregion

            Console.WriteLine($"Collecting data is successful!");
        }
        #endregion

        #region parseSummary
        static Models.Summary ParseSummary(string[] cols, Dictionary<string, int> headerMap, int partsCounterNo)
        {
            // Helper to safely get a column by header name
            string Get(string header)
            {
                if (!headerMap.TryGetValue(header, out int index) || index >= cols.Length)
                    return string.Empty;
                return cols[index];
            }

            string Require(string header)
            {
                var value = Get(header);
                if (string.IsNullOrWhiteSpace(value))
                    throw new ArgumentException($"Column '{header}' cannot be empty.");
                return value;
            }

            // Parse numeric values
            int.TryParse(Require("No. of Blocks"), out int blocksCount);
            int.TryParse(Require("Actual Count"), out int actualCount);
            int.TryParse(Require("NG Mark"), out int ngMark);
            int.TryParse(Require("Unacc"), out int unacc);

            return new Models.Summary
            {
                log_datetime = DateTime.ParseExact(Require("Date & Time"), "ddMMyyyy HH:mm:ss", CultureInfo.InvariantCulture),
                log_wos = Get("WOS"),
                log_item_code = Require("Item Code"),
                log_batch_no = Require("Batch No."),
                log_sublot_no = Require("Sublot No."),
                log_blocks_count = blocksCount,
                log_actual_count = actualCount,
                log_ng_mark = ngMark,
                log_unacc = unacc,
                log_reason = Get("Reason"),
                log_high_unacc_reason = Require("High Unacc. Reason"),
                log_part_counter_no = partsCounterNo,
            };
        }
        #endregion

        #region Parse Breakdown
        static Models.Breakdown ParseBreakdown(string[] cols, Dictionary<string, int> headerMap, int partsCounterNo)
        {
            // Helper to safely get a column by header name
            string Get(string header)
            {
                if (!headerMap.TryGetValue(header, out int index) || index >= cols.Length)
                    return string.Empty;
                return cols[index];
            }
            string Require(string header)
            {
                var value = Get(header);
                if (string.IsNullOrWhiteSpace(value))
                    throw new ArgumentException($"Column '{header}' cannot be empty.");
                return value;
            }

            // Parse numeric values
            int.TryParse(Require("Pallet No."), out int palletno);
            int.TryParse(Require("Actual Count"), out int actualcount);

            return new Models.Breakdown
            {
                log_datetime = DateTime.ParseExact(Require("Date & Time"), "ddMMyyyy HH:mm:ss", CultureInfo.InvariantCulture),
                log_wos = Get("WOS"),
                log_item_code = Require("Item Code"),
                log_batch_no = Require("Batch No."),
                log_sublot_no = Require("Sublot No."),
                log_pallet_no = palletno,
                log_actual_count = actualcount,
                log_op_number = Require("OP Number"),
                log_parts_counter_no = partsCounterNo,
                summaryID = 0
            };
        }
        #endregion

        #region Helper method to move file to error folder safely
        static void MoveFileToErrorFolder(string file, string destErrorFolder)
        {
            try
            {
                var fileName = Path.GetFileName(file);
                var destPath = Path.Combine(destErrorFolder, fileName);

                // If file exists in destination, rename it with timestamp to avoid overwrite
                if (File.Exists(destPath))
                {
                    var timestamp = DateTime.Now.ToString("yyyyMMddHHmm");
                    var newFileName = $"{Path.GetFileNameWithoutExtension(fileName)}_{timestamp}{Path.GetExtension(fileName)}";
                    destPath = Path.Combine(destErrorFolder, newFileName);
                }

                File.Move(file, destPath);
                Console.WriteLine($"Moved file '{fileName}' to error folder.");
            }
            catch (Exception moveEx)
            {
                Console.WriteLine($"Failed to move file '{file}' to error folder: {moveEx.Message}");
                string message = $"Failed to move file '{file}' to error folder: {moveEx.Message}";
                SaveErrorLog(message);
            }
        }
        #endregion

        #region Helper method to move file to archive folder safely
        static void MoveFileToArchieveFolder(string file, string destArchiveFolder)
        {
            try
            {
                var fileName = Path.GetFileName(file);
                var destPath = Path.Combine(destArchiveFolder, fileName);

                // If file exists in destination, rename it with timestamp to avoid overwrite
                if (File.Exists(destPath))
                {
                    var timestamp = DateTime.Now.ToString("yyyyMMddHHmm");
                    var newFileName = $"{Path.GetFileNameWithoutExtension(fileName)}_{timestamp}{Path.GetExtension(fileName)}";
                    destPath = Path.Combine(destArchiveFolder, newFileName);
                }

                File.Move(file, destPath);
                Console.WriteLine($"Moved file '{fileName}' to archive folder.");
            }
            catch (Exception moveEx)
            {
                Console.WriteLine($"Failed to move file '{file}' to error folder: {moveEx.Message}");
                string message = $"Failed to move file '{file}' to error folder: {moveEx.Message}";
                SaveErrorLog(message);
            }
        }
        #endregion

        #region CheckSummaryDuplicate
        private static bool CheckSummaryDuplicate(List<Models.Summary> allSummaries, string connectionString)
        {
            using (var connection = new MySqlConnection(connectionString))
            {
                connection.Open();
                string storedProc = "sp_CheckDuplicateSummary";

                foreach (var summary in allSummaries)
                {
                    var param = new DynamicParameters();
                    param.Add("p_log_item_code", summary.log_item_code);
                    param.Add("p_log_batch_no", summary.log_batch_no);
                    param.Add("p_log_sublot_no", summary.log_sublot_no);

                    // Assume SP returns 1 if duplicate exists, 0 otherwise
                    int result = connection.QuerySingle<int>(
                        storedProc,
                        param,
                        commandType: CommandType.StoredProcedure
                    );

                    if (result == 1)
                    {
                        return true; // Duplicate found
                    }
                }
            }
            return false; // No duplicates found
        }
        #endregion

        #region SaveFileSummary
        private static int SaveFileSummary(List<Models.Summary> allSummaries, string connectionString)
        {
            int insertedId = 0;

            using (var connection = new MySqlConnection(connectionString))
            {
                connection.Open();
                string storedProc = "sp_saveSummary"; // Stored procedure must return LAST_INSERT_ID()

                foreach (var summary in allSummaries)
                {
                    var param = new DynamicParameters();
                    param.Add("p_log_datetime", summary.log_datetime);
                    param.Add("p_log_wos", string.IsNullOrEmpty(summary.log_wos) ? "" : summary.log_wos);
                    param.Add("p_log_item_code", summary.log_item_code);
                    param.Add("p_log_batch_no", summary.log_batch_no);
                    param.Add("p_log_sublot_no", summary.log_sublot_no);
                    param.Add("p_log_blocks_count", summary.log_blocks_count);
                    param.Add("p_log_actual_count", summary.log_actual_count);
                    param.Add("p_log_ng_mark", summary.log_ng_mark);
                    param.Add("p_log_unacc", summary.log_unacc);
                    param.Add("p_log_reason", summary.log_reason);
                    param.Add("p_log_high_unacc_reason", summary.log_high_unacc_reason);
                    param.Add("p_log_part_counter_no", summary.log_part_counter_no);

                    // Stored procedure should SELECT LAST_INSERT_ID() as the result
                    insertedId = connection.QuerySingle<int>(
                        storedProc,
                        param,
                        commandType: CommandType.StoredProcedure
                    );
                }
            }
            return insertedId;
        }
        #endregion

        #region SaveFileBreakdown
        private static void SaveFileBreakdown(List<Models.Breakdown> allBreakdowns, string connectionString, int summaryID)
        {
            using (var connection = new MySqlConnection(connectionString))
            {
                connection.Open();
                string storedProc = "sp_saveBreakdown"; // Replace with your actual stored procedure name

                foreach (var breakdown in allBreakdowns)
                {
                    var param = new
                    {
                        p_log_datetime = breakdown.log_datetime,
                        p_log_wos = string.IsNullOrEmpty(breakdown.log_wos) ? "" : breakdown.log_wos,
                        p_log_item_code = breakdown.log_item_code,
                        p_log_batch_no = breakdown.log_batch_no,
                        p_log_sublot_no = breakdown.log_sublot_no,
                        p_log_pallet_no = breakdown.log_pallet_no,
                        p_log_actual_count = breakdown.log_actual_count,
                        p_log_op_number = breakdown.log_op_number,
                        p_log_parts_counter_no = breakdown.log_parts_counter_no,
                        p_summaryID = summaryID
                    };
                    connection.Execute(storedProc, param, commandType: CommandType.StoredProcedure);
                }
            }
        }
        #endregion

        #region destination folders
        private static string ArchiveFolder()
        {
            DateTime now = DateTime.Now;
            string year = now.Year.ToString();
            string month = now.ToString("MMMM");

            string baseFolder = logArchive;
            string destArchiveFolder = Path.Combine(baseFolder, year, month);
            Directory.CreateDirectory(destArchiveFolder);
            return destArchiveFolder;
        }

        private static string ErrorFolder()
        {
            DateTime now = DateTime.Now;
            string year = now.Year.ToString();
            string month = now.ToString("MMMM");

            string baseFolder = logError;
            string destErrorFolder = Path.Combine(baseFolder, year, month);
            Directory.CreateDirectory(destErrorFolder);
            return destErrorFolder;
        }


        private static string ErrorLogFolder()
        {
            DateTime now = DateTime.Now;
            string year = now.Year.ToString();
            string month = now.ToString("MMMM");
            string day = now.ToString("dd");

            //Comment by: Jpguillermo  |   Aug 26, 2025
            //string errorlog = $"{year}/{month} - ErrorLog.log";
            //End of comment

            string baseFolder = logError;

            //Modified by: Jpguillermo  |   Aug 26, 2025
            //Purpose: Change to Log File and encode only to 1 log file.
            //string destErrorLogFolder = Path.Combine(baseFolder, year, month, errorLog);
            string destErrorLogFolder = Path.Combine(baseFolder, year, month);
            //End ogf modification

            Directory.CreateDirectory(destErrorLogFolder);
            return destErrorLogFolder;
        }
        #endregion

        #region GetPartsCounterNoFromFile
        private static int GetPartsCounterNoFromFile(string fileName)
        {
            int underscoreIndex = fileName.IndexOf('_');
            int dotIndex = fileName.LastIndexOf(".xlsx", StringComparison.OrdinalIgnoreCase);

            if (underscoreIndex != -1 && dotIndex != -1 && underscoreIndex < dotIndex)
            {
                string between = fileName.Substring(underscoreIndex + 1, dotIndex - (underscoreIndex + 1));

                // Extract only digits
                var digitsOnly = new string(between.Where(char.IsDigit).ToArray());

                if (!string.IsNullOrEmpty(digitsOnly))
                {
                    return int.Parse(digitsOnly);
                }
            }

            // Throw or return a special value to indicate error
            throw new ArgumentException($"PartsCounterNo is missing or invalid in file name: {fileName}");
        }
        #endregion

        #region 'SaveErrorLog'
        private static string SaveErrorLog(string message)
        {
            string destErrorLogFolder = ErrorLogFolder();

            // If logFilePath is not set yet, create a new file for this run
            if (string.IsNullOrEmpty(logFilePath))
            {
                //Modified by: Jpguillermo  |   Aug 26, 2025
                //Purpose: Change to Log File and encode only to 1 log file.
                //string logFileName = "errorLog_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".txt";
                string logFileName = "errorLog_" + DateTime.Now.ToString("yyyyMM") + ".log";
                //End Modification Aug 26, 2025

                logFilePath = Path.Combine(destErrorLogFolder, logFileName);
            }

            // Append message with timestamp
            string logEntry = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}";
            File.AppendAllText(logFilePath, logEntry + Environment.NewLine);

            return logFilePath;
        }
        #endregion
    }
}

