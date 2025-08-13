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

                logSource = configuration["FileSettings:LogsSourcePath"];
                logError = configuration["FileSettings:ErrorLogsPath"];
                logArchive = configuration["FileSettings:ArchiveLogsPath"];

                // Validate folders
                if (!Directory.Exists(logSource))
                {
                    Console.WriteLine("Source Directory not found: " + logSource);
                    return false;
                }

                if (!Directory.Exists(logError))
                {
                    Console.WriteLine("Error Directory not found: " + logError);
                    return false;
                }

                if (!Directory.Exists(logArchive))
                {
                    Console.WriteLine("Archive Directory not found: " + logArchive);
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading configuration: {ex.Message}");
                return false;
            }
        }
        #endregion

        #region ProcessFile
        private static void ProcessFile()
        {
            var excelFiles = Directory.GetFiles(logSource, "*.xlsx");
            if (excelFiles.Length == 0)
            {
                Console.WriteLine("No XLSX files found in: " + logSource);
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
                try
                {
                    var fileName = Path.GetFileName(file); // Get only file name
                    int partsCounterNo = GetPartsCounterNoFromFile(fileName);

                    using (var package = new ExcelPackage(new FileInfo(file)))
                    {
                        var worksheet = package.Workbook.Worksheets.FirstOrDefault();
                        if (worksheet == null)
                        {
                            Console.WriteLine($"No worksheet found in file {file}");
                            MoveFileToErrorFolder(file, destErrorFolder);
                            continue;
                        }

                        #region Parse summary from second row (row 2)
                        var summaryCols = Enumerable.Range(1, worksheet.Dimension.End.Column)
                                            .Select(c => worksheet.Cells[2, c].Text)
                                            .ToArray();
                        var summary = ParseSummary(summaryCols, partsCounterNo);
                        allSummaries.Add(summary);

                        #endregion

                        #region Parse breakdown starting from fourth row (row 4)
                        for (int row = 4; row <= worksheet.Dimension.End.Row; row++)
                        {
                            var cols = Enumerable.Range(1, worksheet.Dimension.End.Column)
                                                 .Select(c => worksheet.Cells[row, c].Text)
                                                 .ToArray();
                            var breakdown = ParseBreakdown(cols, partsCounterNo);
                            allBreakdowns.Add(breakdown);
                        }
                        #endregion

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
                    Console.WriteLine($"Error processing file {file}: {ex.Message}");
                    continue;
                }
            }
            #endregion

            Console.WriteLine($"Collecting data is successful!");
        }
        #endregion

        #region parseSummary
        static Models.Summary ParseSummary(string[] cols, int partsCounterNo)
        {
            // Check each column for empty/whitespace
            for (int i = 0; i < 7; i++)
            {
                if (string.IsNullOrWhiteSpace(cols[i]))
                {
                    throw new ArgumentException($"Column index {i} cannot be empty.");
                }
            }

            int.TryParse(cols[4], out int blocksCount);
            int.TryParse(cols[5], out int actualCount);
            int.TryParse(cols[6], out int ngMark);
            int.TryParse(cols[7], out int unacc);

            return new Models.Summary
            {
                log_datetime = DateTime.ParseExact(cols[0], "ddMMyyyy HH:mm:ss", CultureInfo.InvariantCulture),
                log_order_no = "",
                log_item_code = cols[1],
                log_batch_no = cols[2],
                log_sublot_no = cols[3],
                log_blocks_count = blocksCount,
                log_actual_count = actualCount,
                log_ng_mark = ngMark,
                log_unacc = unacc,
                log_reason = cols.Length > 8 ? cols[8] : "",
                log_high_unacc_reason = cols.Length > 9 ? cols[9] : "",
                log_part_counter_no = partsCounterNo,
            };
        }
        #endregion

        #region Parse Breakdown
        static Models.Breakdown ParseBreakdown(string[] cols, int partsCounterNo)
        {
            // Check each column for empty/whitespace
            for (int i = 0; i < 6; i++)
            {
                if (string.IsNullOrWhiteSpace(cols[i]))
                {
                    throw new ArgumentException($"Column index {i} cannot be empty.");
                }
            }

            int.TryParse(cols[4], out int palletno);
            int.TryParse(cols[5], out int actualcount);
            return new Models.Breakdown
            {
                log_datetime = DateTime.ParseExact(cols[0], "ddMMyyyy HH:mm:ss", CultureInfo.InvariantCulture),
                log_order_no = "",
                log_item_code = cols[1],
                log_batch_no = cols[2],
                log_sublot_no = cols[3],
                log_pallet_no = palletno,
                log_actual_count = actualcount,
                log_op_number = cols[6],
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
                    var timestamp = DateTime.Now.ToString("yyyyMMddHHmmssfff");
                    var newFileName = $"{Path.GetFileNameWithoutExtension(fileName)}_{timestamp}{Path.GetExtension(fileName)}";
                    destPath = Path.Combine(destErrorFolder, newFileName);
                }

                File.Move(file, destPath);
                Console.WriteLine($"Moved file '{fileName}' to error folder.");
            }
            catch (Exception moveEx)
            {
                Console.WriteLine($"Failed to move file '{file}' to error folder: {moveEx.Message}");
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
                    var timestamp = DateTime.Now.ToString("yyyyMMddHHmmssfff");
                    var newFileName = $"{Path.GetFileNameWithoutExtension(fileName)}_{timestamp}{Path.GetExtension(fileName)}";
                    destPath = Path.Combine(destArchiveFolder, newFileName);
                }

                File.Move(file, destPath);
                Console.WriteLine($"Moved file '{fileName}' to archive folder.");
            }
            catch (Exception moveEx)
            {
                Console.WriteLine($"Failed to move file '{file}' to error folder: {moveEx.Message}");
            }
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
                    param.Add("p_log_order_no", string.IsNullOrEmpty(summary.log_order_no) ? " " : summary.log_order_no);
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
                        p_log_order_no = breakdown.log_order_no,
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

    }
}

