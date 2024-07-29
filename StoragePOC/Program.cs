using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Data.Entity.Infrastructure;
using OfficeOpenXml;
using System.Data.Entity.Validation;

namespace StoragePOC
{
    class BulkOperation
    {
        // Main function - M
        static void Main(string[] args)
        {
            //Testing
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            //Get the Excel file location
            //@@raja Change this path and ask user to upload a excel file 
            string filePath = ConfigurationManager.AppSettings["ImportDataFilePath"];
            //Get the Error log file location
            //@@raja try to integrate with existing log file path 
            string errorFilePath = ConfigurationManager.AppSettings["ErrorDataFilePath"];
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"The file {filePath} does not exist.");
                Console.WriteLine("Press any key to exit...");
                Console.ReadKey();
                return;
            }
            //dataTable to store valid records from Registraion sheet
            DataTable dataTable = new DataTable();
            CreateDataTableSchema(dataTable);
            // DataTable to store valid records from Storage sheet
            DataTable secondTable = new DataTable();
            CreateSecondTableSchema(secondTable);
            //dataTable to store invalid records from both sheets
            DataTable errorDetailsTable = new DataTable();
            CreateErrorDataTableSchema(errorDetailsTable);
            // Read the Excel
            try
            {
                using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
                {
                    // Read and Fill Registration sheet
                    ExcelWorksheet workSheet = package.Workbook.Worksheets[0];
                    int rowCount = workSheet.Dimension.Rows;
                    Dictionary<string, int> columnIndexes = GetColumnIndexes(workSheet);
                    Console.WriteLine();
                    // Define expected columns
                    string[] expectedColumns = {
                    "PatientId", "SiteNo", "VisitName",
                    "DOBSCollection", "TOBSCollection", "RequisitionId", "BarCodeValue",
                    "TypeOfSample", "Remarks", "SubjectId", "PatientInitials",
                    "PatientName", "Height", "Weight", "Gender","ReceivedDateTime","ReceivedCondition",
                        "Status","DateOfBirth","ClientTubeNo","StudyId","ProcessingDate"
                    };
                    // Add columns to DataTable and check for missing columns
                    foreach (var column in expectedColumns)
                    {
                        if (columnIndexes.ContainsKey(column))
                        {
                            if (!dataTable.Columns.Contains(column)) // Check if column already exists
                            {
                                dataTable.Columns.Add(column);
                            }
                        }
                        else
                        {
                            Console.WriteLine(column);
                        }
                    }
                    for (int row = 2; row <= rowCount; row++)
                    {
                        DataRow dataRow = dataTable.NewRow();
                        // Check if the row is empty or not, if empty then the row won't be added else it will be added
                        bool rowIsEmpty = FillDataRow(dataRow, workSheet, columnIndexes, row, errorDetailsTable);
                        if (!rowIsEmpty && dataRow.Table.Columns.Contains("BarCodeValue") && !string.IsNullOrEmpty(dataRow["BarCodeValue"].ToString()))
                        {
                            dataTable.Rows.Add(dataRow);
                        }
                    }
                    Console.WriteLine($"Total No Of Valid records from Registration Sheet: {dataTable.Rows.Count}");
                    Console.WriteLine("");
                    //Read and Fill Storage sheet
                    ExcelWorksheet secondSheet = package.Workbook.Worksheets[1];
                    int secondRowCount = secondSheet.Dimension.Rows;
                    Dictionary<string, int> secondColumnIndexes = GetColumnIndexes(secondSheet);

                    string[] secondSheetColumns = {"PatientId","SiteNo", "VisitName", "TOBSCollection", "RequisitionId",
                                                   "BarCodeValue", "TypeOfSample", "Remarks", "DOBSCollection","ReceivedCondition"};
                    foreach(var column in secondSheetColumns)
                    {
                        if (secondColumnIndexes.ContainsKey(column))
                        {
                            if (!secondTable.Columns.Contains(column))
                            {
                                secondTable.Columns.Add(column);
                            }
                        }
                        else
                        {
                            Console.WriteLine(column);
                        }
                    }
                    for(int row = 2; row <= secondRowCount; row++)
                    {
                        DataRow dataRow = secondTable.NewRow();
                        bool secondTableIsEmpty = FillDataRowSecondTable(dataRow, secondSheet, secondColumnIndexes, row, errorDetailsTable);
                        if (!secondTableIsEmpty)
                        {
                            secondTable.Rows.Add(dataRow);
                        }
                    }
                    Console.WriteLine($"Total No Of Valid records from Storage sheet: {secondTable.Rows.Count}");
                    Console.WriteLine();
                }
                /// Call Update TRF_RegistrationInfo
                UpdateRegistrationInfos(dataTable, errorDetailsTable);
                // Call Update TRF_Reg_BarCodes only if the UpdateRegistrationInfos executed successfully
                if (UpdateRegistrationInfos(dataTable, errorDetailsTable).success)
                {
                    Console.WriteLine($"Total No Of Updated records in TRF_RegistrationInfo Table - {UpdateRegistrationInfos(dataTable, errorDetailsTable).rowCount}");
                    Console.WriteLine("");
                    int rowCount = UpdateRegBarCodes(dataTable, errorDetailsTable);
                    Console.WriteLine($"Total No Of Updated records in TRF_Reg_BarCode Table - {rowCount}");
                    Console.WriteLine("");
                }
                // Call InsertOrUpdate function for Storage
                InsertOrUpdateSampleData(secondTable, errorDetailsTable);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
            // Store the error log file in a csv file 
            string directoryPath = ConfigurationManager.AppSettings["ErrorDataFilePath"]; ;
            string uploadedFileName = Path.GetFileNameWithoutExtension(filePath);
            string dateTimeValue = DateTime.Now.ToString("dd-MM-yyyy");
            string combinedFileName = $"{uploadedFileName}_{dateTimeValue}";
            string outputFile = Path.Combine(directoryPath, $"{combinedFileName}.csv");
            try
            {
                if (!Directory.Exists(directoryPath))
                {
                    Directory.CreateDirectory(directoryPath);
                }
                bool isNewFile = !File.Exists(outputFile);
                using (StreamWriter writer = new StreamWriter(outputFile, true))
                {
                    if (isNewFile)
                    {
                       writer.WriteLine("SheetName,RowNumber,ColumnName,ErrorDescription,LoggedDateTime");
                    }
                    foreach (DataRow row in errorDetailsTable.Rows)
                    {
                        for (int i = 0; i < errorDetailsTable.Columns.Count; i++)
                        {
                            writer.Write(row[i]);
                            if (i < errorDetailsTable.Columns.Count - 1)
                            {
                                writer.Write(",");
                            }
                        }
                        writer.Write($",{DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                        writer.WriteLine();
                    }
                }
                Console.WriteLine($"Total number of invalid records from Excel: {errorDetailsTable.Rows.Count}");
            }
            catch(Exception ex)
            {
                Console.WriteLine($"An error occurred while writing to the file: {ex.Message}");
            }
            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }

        // Create Schema to add valid rows from Registration sheet - Help M1
        static void CreateDataTableSchema(DataTable dataTable)
        {
            dataTable.Columns.Add("StudyId", typeof(string));
            dataTable.Columns.Add("PatientId", typeof(string));
            dataTable.Columns.Add("SiteNo", typeof(string));
            dataTable.Columns.Add("VisitName", typeof(string));
            dataTable.Columns.Add("DOBSCollection", typeof(DateTime));
            dataTable.Columns.Add("TOBSCollection", typeof(TimeSpan));
            dataTable.Columns.Add("CollectionDate", typeof(DateTime));
            dataTable.Columns.Add("RequisitionId", typeof(string));
            dataTable.Columns.Add("BarCodeValue", typeof(string));
            dataTable.Columns.Add("TypeOfSample", typeof(string));
            dataTable.Columns.Add("Remarks", typeof(string));
            dataTable.Columns.Add("SubjectId", typeof(string));
            dataTable.Columns.Add("PatientInitials", typeof(string));
            dataTable.Columns.Add("PatientName", typeof(string));
            dataTable.Columns.Add("Gender", typeof(string));
            dataTable.Columns.Add("Height", typeof(string));
            dataTable.Columns.Add("Weight", typeof(string));
            dataTable.Columns.Add("RegField3", typeof(string));
            dataTable.Columns.Add("ReceivedDateTime", typeof(DateTime));
            dataTable.Columns.Add("ReceivedCondition", typeof(string));
            dataTable.Columns.Add("Status", typeof(string));
            dataTable.Columns.Add("DateOfBirth", typeof(DateTime));
            dataTable.Columns.Add("ProcessingDate", typeof(DateTime));
            dataTable.Columns.Add("ApprovedStatus", typeof(bool));
            dataTable.Columns.Add("ClientTubeNo", typeof(string));
            dataTable.Columns.Add("SystemId", typeof(long));
            dataTable.Columns.Add("Cryobox", typeof(string));
            dataTable.Columns.Add("CryoboxWellPosition", typeof(int));
            dataTable.Columns.Add("ProjectCode", typeof(string));
        }

        // Create Schema to add invalid rows from excel - Help M2
        static void CreateErrorDataTableSchema(DataTable errorDetailsTable)
        {
            errorDetailsTable.Columns.Add("SheetName", typeof(string));
            errorDetailsTable.Columns.Add("RowIndex", typeof(int));
            errorDetailsTable.Columns.Add("ErrorField", typeof(string));
            errorDetailsTable.Columns.Add("ErrorDescription", typeof(string));
        }

        // Create Schema to add valid rows from Storage sheet - Help M3
        static void CreateSecondTableSchema(DataTable secondTable)
        {
            secondTable.Columns.Add("PatientId", typeof(string));
            secondTable.Columns.Add("SiteNo", typeof(string));
            secondTable.Columns.Add("VisitName", typeof(string));
            secondTable.Columns.Add("TOBSCollection", typeof(TimeSpan));
            secondTable.Columns.Add("DOBSCollection", typeof(DateTime));
            secondTable.Columns.Add("RequisitionId", typeof(string));
            secondTable.Columns.Add("BarCodeValue", typeof(string));
            secondTable.Columns.Add("TypeOfSample", typeof(string));
            secondTable.Columns.Add("Remarks", typeof(string));
            secondTable.Columns.Add("ApprovedStatus", typeof(bool));
            secondTable.Columns.Add("ProjectCode", typeof(string));
            secondTable.Columns.Add("SystemId", typeof(long));
            secondTable.Columns.Add("Cryobox", typeof(string));
            secondTable.Columns.Add("CryoboxWellPosition", typeof(int));
            secondTable.Columns.Add("ReceivedCondition", typeof(string));
            secondTable.Columns.Add("Status", typeof(string));
            secondTable.Columns.Add("Location", typeof(string));
        }

        // Get the Column and Indexes - Help M4
        static Dictionary<string, int> GetColumnIndexes(ExcelWorksheet workSheet)
        {
            // Retrieve column indexes from Excel header row
            Dictionary<string, int> columnIndexes = new Dictionary<string, int>();
            var headersRow = workSheet.Cells[1, 1, 1, workSheet.Dimension.Columns];
            foreach (var cell in headersRow)
            {
                columnIndexes[cell.Text] = cell.Start.Column;
            }
            return columnIndexes;
        }

        // Checks the value from the Registration sheet and fills the data - Help M5
        static bool FillDataRow(DataRow dataRow, ExcelWorksheet workSheet, Dictionary<string, int> columnIndexes, int row, DataTable errorDetailsTable)
        {
            bool rowIsEmpty = false;
            void AddError(string fieldName, string errorDescription)
            {
                DataRow errorRow = errorDetailsTable.NewRow();
                errorRow["SheetName"] = "Registration";
                errorRow["RowIndex"] = row-1;
                errorRow["ErrorField"] = fieldName;
                errorRow["ErrorDescription"] = errorDescription;
                errorDetailsTable.Rows.Add(errorRow);
                rowIsEmpty = true;
            }
            string GetValue(string columnName)
            {
                var value = workSheet.Cells[row, columnIndexes[columnName]].Value?.ToString();
                if (string.IsNullOrWhiteSpace(value))
                {
                    AddError(columnName, "String is Spaces");
                    rowIsEmpty = true;
                    return null;
                }
                return value;
            }
            DateTime? GetNullableDate(string columnName)
            {
                var value = workSheet.Cells[row, columnIndexes[columnName]].Value?.ToString();
                if (string.IsNullOrWhiteSpace(value))
                {
                    AddError(columnName, "Date Value is missing");
                    return null;
                }

                if (IsValidDate(value, columnName, out DateTime? parsedDate))
                {
                    return parsedDate;
                }
                else
                {
                    rowIsEmpty = true;
                    return null;
                }
            }
            bool IsValidDate(string dateStr, string columnName, out DateTime? parsedDate)
            {
                parsedDate = null;
                // Attempt to parse the input date string
                if (!DateTime.TryParse(dateStr, out DateTime dateValue))
                {
                    AddError(columnName, "Invalid date format");
                    return false;
                }
                DateTime minDate = new DateTime(1900, 1, 1);
                DateTime maxDate = DateTime.Now; // Today's date or any other logical upper limit
                // Check if the date is within the acceptable range
                if (dateValue < minDate || dateValue > maxDate)
                {
                    AddError(columnName, "Date out of range");
                    return false;
                }
                // If all checks pass, return the parsed date
                parsedDate = dateValue;
                return true;
            }
            TimeSpan? GetNullableTime(string columnName)
            {
                var value = workSheet.Cells[row, columnIndexes[columnName]].Value?.ToString();
                if (string.IsNullOrWhiteSpace(value))
                {
                    AddError(columnName, "Time Value is missing");
                    rowIsEmpty = true;
                    return null;
                }
                if (IsValidTime(value, columnName, out TimeSpan? parsedTime))
                {
                    return parsedTime;
                }
                else
                {
                    rowIsEmpty = true;
                    return null;
                }
            }
            bool IsValidTime(string timeStr, string columnName, out TimeSpan? parsedTime)
            {
                parsedTime = null;

                // Attempt to parse as DateTime
                if (DateTime.TryParse(timeStr, out DateTime dateTimeValue))
                {
                    parsedTime = dateTimeValue.TimeOfDay;
                    return true;
                }
                // Attempt to parse as fractional day time
                if (double.TryParse(timeStr, out double fractionDay))
                {
                    TimeSpan timeSpan = TimeSpan.FromDays(fractionDay);
                    parsedTime = timeSpan;
                    return true;
                }
                // Split the time string by ':'
                string[] timeParts = timeStr.Split(':');

                // Ensure the time string is in the expected format
                if (timeParts.Length != 3)
                {
                    AddError(columnName, "Time format is incorrect, expected hh:mm:ss");
                    return false;
                }
                // Validate and parse hour
                if (!int.TryParse(timeParts[0], out int hour) || hour < 0 || hour > 23)
                {
                    AddError(columnName, "Hour out of range (0-23)");
                    return false;
                }
                // Validate and parse minute
                if (!int.TryParse(timeParts[1], out int minute) || minute < 0 || minute > 59)
                {
                    AddError(columnName, "Minute out of range (0-59)");
                    return false;
                }
                // Validate and parse second
                if (!int.TryParse(timeParts[2], out int second) || second < 0 || second > 59)
                {
                    AddError(columnName, "Second out of range (0-59)");
                    return false;
                }
                // Attempt to parse using TimeSpan
                if (!TimeSpan.TryParseExact(timeStr, "hh\\:mm\\:ss", null, out TimeSpan timeValue))
                {
                    AddError(columnName, "Invalid time format, expected hh:mm:ss");
                    return false;
                }
                // If all checks pass, set the parsedTime
                parsedTime = timeValue;
                return true;
            }
            void ValidateDateRange(DateTime? date, string columnName)
            {
                DateTime minDate = new DateTime(1900, 1, 1);
                DateTime maxDate = DateTime.Today;
                if (date.HasValue && (date.Value < minDate || date.Value > maxDate))
                {
                    AddError(columnName, "Date is out of the valid range");
                }
            }
            void ValidateTimeRange(TimeSpan? time, string columnName)
            {
                TimeSpan minTime = TimeSpan.Zero;
                TimeSpan maxTime = new TimeSpan(23, 59, 59);
                if (time.HasValue && (time.Value < minTime || time.Value > maxTime))
                {
                    AddError(columnName, "Time is out of the valid range");
                }
            }
            void ValidateDecimalRange(decimal? value, string columnName, decimal minValue, decimal maxValue)
            {
                if (value.HasValue && (value.Value < minValue || value.Value > maxValue))
                {
                    AddError(columnName, "Value is out of the valid range");
                }
            }
            void ValidateStringLength(string value, string columnName, int minLength, int maxLength)
            {
                if (!string.IsNullOrWhiteSpace(value) && (value.Length < minLength || value.Length > maxLength))
                {
                    AddError(columnName, $"String length is out of the valid range ({minLength}-{maxLength})");
                }
            }
            string studyId = GetValue("StudyId");
            ValidateStringLength(studyId, "StudyId", 1, 50);
            dataRow["StudyId"] = studyId ?? (object)DBNull.Value;
            string gender = GetValue("Gender");
            ValidateStringLength(gender, "Gender", 1, 10);
            dataRow["Gender"] = gender ?? (object)DBNull.Value;
            string patientId = GetValue("PatientId");
            ValidateStringLength(patientId, "PatientId", 1, 50);
            dataRow["PatientId"] = patientId ?? (object)DBNull.Value;
            string siteNo = GetValue("SiteNo");
            ValidateStringLength(siteNo, "SiteNo", 1, 50);
            dataRow["SiteNo"] = siteNo ?? (object)DBNull.Value;
            string visitName = GetValue("VisitName");
            ValidateStringLength(visitName, "VisitName", 1, 100);
            dataRow["VisitName"] = visitName ?? (object)DBNull.Value;
            dataRow["DOBSCollection"] = (object)GetNullableDate("DOBSCollection") ?? DBNull.Value;
            dataRow["TOBSCollection"] = (object)GetNullableTime("TOBSCollection") ?? DBNull.Value;
            string requisitionId = GetValue("RequisitionId");
            ValidateStringLength(requisitionId, "RequisitionId", 1, 50);
            dataRow["RequisitionId"] = requisitionId ?? (object)DBNull.Value;
            string barCodeValue = GetValue("BarCodeValue");
            ValidateStringLength(barCodeValue, "BarCodeValue", 1, 50);
            dataRow["BarCodeValue"] = barCodeValue ?? (object)DBNull.Value;
            string typeOfSample = GetValue("TypeOfSample");
            ValidateStringLength(typeOfSample, "TypeOfSample", 1, 50);
            dataRow["TypeOfSample"] = typeOfSample ?? (object)DBNull.Value;
            string remarks = GetValue("Remarks");
            ValidateStringLength(remarks, "Remarks", 0, 500); // Remarks can be empty
            dataRow["Remarks"] = remarks ?? (object)DBNull.Value;
            string subjectId = GetValue("SubjectId");
            ValidateStringLength(subjectId, "SubjectId", 1, 50);
            dataRow["SubjectId"] = subjectId ?? (object)DBNull.Value;
            string patientInitials = GetValue("PatientInitials");
            ValidateStringLength(patientInitials, "PatientInitials", 1, 50);
            dataRow["PatientInitials"] = patientInitials ?? (object)DBNull.Value;
            string patientName = GetValue("PatientName");
            ValidateStringLength(patientName, "PatientName", 1, 100);
            dataRow["PatientName"] = patientName ?? (object)DBNull.Value;
            string height = GetValue("Height");
            ValidateStringLength(height, "Height", 1, 10);
            dataRow["Height"] = height ?? (object)DBNull.Value;
            string weight = GetValue("Weight");
            ValidateStringLength(weight, "Weight", 1, 10);
            dataRow["Weight"] = weight ?? (object)DBNull.Value;
            dataRow["ReceivedDateTime"] = (object)GetNullableDate("ReceivedDateTime") ?? DBNull.Value;
            dataRow["ProcessingDate"] = (object)GetNullableDate("ProcessingDate") ?? DBNull.Value;
            string receivedCondition = GetValue("ReceivedCondition");
            ValidateStringLength(receivedCondition, "ReceivedCondition", 1, 100);
            dataRow["ReceivedCondition"] = receivedCondition ?? (object)DBNull.Value;
            string status = GetValue("Status");
            ValidateStringLength(status, "Status", 1, 50);
            dataRow["Status"] = status ?? (object)DBNull.Value;
            dataRow["DateOfBirth"] = (object)GetNullableDate("DateOfBirth") ?? DBNull.Value;
            string clientTubeNo = GetValue("ClientTubeNo");
            ValidateStringLength(clientTubeNo, "ClientTubeNo", 1, 50);
            dataRow["ClientTubeNo"] = clientTubeNo ?? (object)DBNull.Value;
            // Validate date ranges
            ValidateDateRange(dataRow.Field<DateTime?>("DOBSCollection"), "DOBSCollection");
            ValidateDateRange(dataRow.Field<DateTime?>("ReceivedDateTime"), "ReceivedDateTime");
            ValidateDateRange(dataRow.Field<DateTime?>("ProcessingDate"), "ProcessingDate");
            ValidateDateRange(dataRow.Field<DateTime?>("DateOfBirth"), "DateOfBirth");
            // Validate time ranges
            ValidateTimeRange(dataRow.Field<TimeSpan?>("TOBSCollection"), "TOBSCollection");
            // Validate decimal ranges for Height and Weight (assuming example ranges)
            if (decimal.TryParse(height, out decimal heightValue))
            {
                ValidateDecimalRange(heightValue, "Height", 0m, 300m);
            }
            else
            {
                AddError("Height", "Invalid decimal value");
            }
            if (decimal.TryParse(weight, out decimal weightValue))
            {
                ValidateDecimalRange(weightValue, "Weight", 0m, 500m);
            }
            else
            {
                AddError("Weight", "Invalid decimal value");
            }
            // Merge DOBSCollection and TOBSCollection into CollectionDate
            DateTime? dob = dataRow.Field<DateTime?>("DOBSCollection");
            TimeSpan? tob = dataRow.Field<TimeSpan?>("TOBSCollection");
            if (dob.HasValue && tob.HasValue)
            {
                dataRow["CollectionDate"] = dob.Value.Date + tob.Value;
            }
            else if (dob.HasValue)
            {
                dataRow["CollectionDate"] = dob.Value;
            }
            else if (tob.HasValue)
            {
                dataRow["CollectionDate"] = DateTime.Today + tob.Value;
            }
            else
            {
                dataRow["CollectionDate"] = DateTime.Now;
            }
            // Calculate RegField3 if both CollectionDate and DateOfBirth are present
            DateTime? collectionDate = dataRow.Field<DateTime?>("CollectionDate");
            DateTime? dateOfBirth = dataRow.Field<DateTime?>("DateOfBirth");
            if (collectionDate.HasValue && dateOfBirth.HasValue)
            {
                double daysDifference = (collectionDate.Value - dateOfBirth.Value).TotalDays;
                dataRow["RegField3"] = Math.Round(daysDifference, 0).ToString();
            }
            else
            {
                dataRow["RegField3"] = DBNull.Value;
            }
            return rowIsEmpty;
        }

        // Checks the value from the Storage sheet and fills the data - Help M6
        static bool FillDataRowSecondTable(DataRow dataRow, ExcelWorksheet workSheet, Dictionary<string, int> columnIndexes, int row, DataTable errorDetailsTable)
        {
            bool rowIsEmpty = false;
            void AddError(string fieldName, string errorDescription)
            {
                DataRow errorRow = errorDetailsTable.NewRow();
                errorRow["SheetName"] = "Storage";
                errorRow["RowIndex"] = row - 1;
                errorRow["ErrorField"] = fieldName;
                errorRow["ErrorDescription"] = errorDescription;
                errorDetailsTable.Rows.Add(errorRow);
                rowIsEmpty = true;
            }
            string GetValue(string columnName)
            {
                try
                {
                    var value = workSheet.Cells[row, columnIndexes[columnName]].Value?.ToString();
                    if (string.IsNullOrWhiteSpace(value))
                    {
                        AddError(columnName, "String is Spaces");
                        rowIsEmpty = true;
                        return null;
                    }
                    return value;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Exception in {columnName}: {ex.Message}");
                    AddError(columnName, ex.Message);
                    return null;
                }
            }
            DateTime? GetNullableDate(string columnName)
            {
                try
                {
                    var value = workSheet.Cells[row, columnIndexes[columnName]].Value?.ToString();
                    if (string.IsNullOrWhiteSpace(value))
                    {
                        AddError(columnName, "Date Value is missing");
                        return null;
                    }

                    if (IsValidDate(value, columnName, out DateTime? parsedDate))
                    {
                        return parsedDate;
                    }
                    else
                    {
                        rowIsEmpty = true;
                        return null;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Exception in {columnName}: {ex.Message}");
                    AddError(columnName, ex.Message);
                    return null;
                }
            }
            bool IsValidDate(string dateStr, string columnName, out DateTime? parsedDate)
            {
                parsedDate = null;

                if (!DateTime.TryParse(dateStr, out DateTime dateValue))
                {
                    AddError(columnName, "Invalid date format");
                    return false;
                }

                DateTime minDate = new DateTime(1900, 1, 1);
                DateTime maxDate = DateTime.Now;

                if (dateValue < minDate || dateValue > maxDate)
                {
                    AddError(columnName, "Date out of range");
                    return false;
                }
                parsedDate = dateValue;
                return true;
            }      
            TimeSpan? GetNullableTime(string columnName)
            {
                try
                {
                    var value = workSheet.Cells[row, columnIndexes[columnName]].Value?.ToString();
                    if (string.IsNullOrWhiteSpace(value))
                    {
                        AddError(columnName, "Time Value is missing");
                        rowIsEmpty = true;
                        return null;
                    }

                    if (IsValidTime(value, columnName, out TimeSpan? parsedTime))
                    {
                        return parsedTime;
                    }
                    else
                    {
                        rowIsEmpty = true;
                        return null;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Exception in {columnName}: {ex.Message}");
                    AddError(columnName, ex.Message);
                    return null;
                }
            }
            bool IsValidTime(string timeStr, string columnName, out TimeSpan? parsedTime)
            {
                parsedTime = null;

                if (DateTime.TryParse(timeStr, out DateTime dateTimeValue))
                {
                    parsedTime = dateTimeValue.TimeOfDay;
                    return true;
                }
                if (double.TryParse(timeStr, out double fractionDay))
                {
                    TimeSpan timeSpan = TimeSpan.FromDays(fractionDay);
                    parsedTime = timeSpan;
                    return true;
                }
                string[] timeParts = timeStr.Split(':');
                if (timeParts.Length != 3)
                {
                    AddError(columnName, "Time format is incorrect, expected hh:mm:ss");
                    return false;
                }
                if (!int.TryParse(timeParts[0], out int hour) || hour < 0 || hour > 23)
                {
                    AddError(columnName, "Hour out of range (0-23)");
                    return false;
                }
                if (!int.TryParse(timeParts[1], out int minute) || minute < 0 || minute > 59)
                {
                    AddError(columnName, "Minute out of range (0-59)");
                    return false;
                }
                if (!int.TryParse(timeParts[2], out int second) || second < 0 || second > 59)
                {
                    AddError(columnName, "Second out of range (0-59)");
                    return false;
                }
                if (!TimeSpan.TryParseExact(timeStr, "hh\\:mm\\:ss", null, out TimeSpan timeValue))
                {
                    AddError(columnName, "Invalid time format, expected hh:mm:ss");
                    return false;
                }
                parsedTime = timeValue;
                return true;
            }
            try
            {
                string patientId = GetValue("PatientId");
                dataRow["PatientId"] = patientId ?? (object)DBNull.Value;
                string siteNo = GetValue("SiteNo");
                dataRow["SiteNo"] = siteNo ?? (object)DBNull.Value;
                string visitName = GetValue("VisitName");
                dataRow["VisitName"] = visitName ?? (object)DBNull.Value;
                dataRow["DOBSCollection"] = (object)GetNullableDate("DOBSCollection") ?? DBNull.Value;
                dataRow["TOBSCollection"] = (object)GetNullableTime("TOBSCollection") ?? DBNull.Value;
                string requisitionId = GetValue("RequisitionId");
                dataRow["RequisitionId"] = requisitionId ?? (object)DBNull.Value;
                string barCodeValue = GetValue("BarCodeValue");
                dataRow["BarCodeValue"] = barCodeValue ?? (object)DBNull.Value;
                string typeOfSample = GetValue("TypeOfSample");
                dataRow["TypeOfSample"] = typeOfSample ?? (object)DBNull.Value;
                string remarks = GetValue("Remarks");
                dataRow["Remarks"] = remarks ?? (object)DBNull.Value;
                string receivedCondition = GetValue("ReceivedCondition");
                dataRow["ReceivedCondition"] = receivedCondition ?? (object)DBNull.Value;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Exception in row processing: {ex.Message}");
                rowIsEmpty = true;
            }
            // For testing I'm passing test values in actual the four cases will perform
            //dataRow["Cryobox"] = "Box2";
            return rowIsEmpty;
        }

        // Update TRF_RegistrationInfos - R
        static (bool success, int rowCount) UpdateRegistrationInfos(DataTable dataTable, DataTable errorDetailsTable)
        {
            bool success = true;
            int rowCount = 0;
            string ModuleName = "Registration";
            string PageName = "RegistrationUpdate";
            string UserRemarks = "Update TRF_RegistrationInfo";
            using (var context = new LIMSDevContext())
            {
                foreach (DataRow row in dataTable.Rows)
                {
                    int rowIndex = dataTable.Rows.IndexOf(row);
                    try
                    {
                        string storeSystemRemarks = "";
                        string requisitionId = row["RequisitionId"].ToString();
                        var registrationInfo = context.TRF_RegistrationInfos.FirstOrDefault(r => r.RequisitionId == requisitionId && r.ApprovedStatus == false);
                        row["ProjectCode"] = registrationInfo?.ProjectCode;
                        long systemId = 0;
                        if (registrationInfo?.SystemId != null && long.TryParse(registrationInfo.SystemId.ToString(), out systemId))
                        {
                            row["SystemId"] = systemId;
                        }
                        else
                        {
                            row["SystemId"] = DBNull.Value; 
                        }
                        if (registrationInfo != null)
                        {
                                registrationInfo.PatientId = row["PatientId"].ToString();
                                registrationInfo.PatientInitials = row["PatientInitials"].ToString();
                                registrationInfo.VisitName = row["VisitName"].ToString();
                                registrationInfo.RequisitionId = row["RequisitionId"].ToString();
                                registrationInfo.SubjectId = row["SubjectId"].ToString();
                                registrationInfo.PatientName = row["PatientName"].ToString();
                                registrationInfo.Gender = row["Gender"].ToString();
                                registrationInfo.CollectionDate = Convert.ToDateTime(row["CollectionDate"]);
                                registrationInfo.DateOfBirth = Convert.ToDateTime(row["DateOfBirth"]);
                                registrationInfo.ProcessingDate = Convert.ToDateTime(row["ProcessingDate"]);
                                registrationInfo.ReceivedDateTime = Convert.ToDateTime(row["ReceivedDateTime"]);
                                registrationInfo.RegField3 = row["RegField3"].ToString();
                                registrationInfo.RegistrationStatus = true;
                                registrationInfo.RegistrationCheckStatus = false;
                                registrationInfo.ResultCheckStatus = false;
                                registrationInfo.ApprovedStatus = false;
                                registrationInfo.HoldStatus = false;
                                registrationInfo.IsLock = false;
                                registrationInfo.ReportGeneratedStatus = false;
                                registrationInfo.CreatedOn = DateTime.Now;
                                // Parse Height and Weight as decimal?
                                if (decimal.TryParse(row["Height"].ToString(), out decimal height))
                                {
                                    registrationInfo.Height = height;
                                }
                                else
                                {
                                    registrationInfo.Height = null; // or default value
                                }
                                if (decimal.TryParse(row["Weight"].ToString(), out decimal weight))
                                {
                                    registrationInfo.Weight = weight;
                                }
                                else
                                {
                                    registrationInfo.Weight = null; // or default value
                                }
                                // Calculate AgeValue based on DateOfBirth
                                if (row["DateOfBirth"] != DBNull.Value)
                                {
                                    DateTime dateOfBirth = (DateTime)row["DateOfBirth"];
                                    TimeSpan age = DateTime.Now - dateOfBirth;
                                    registrationInfo.AgeValue = (long)age.TotalDays / 365; // Calculate age in years
                                }
                                else
                                {
                                    registrationInfo.AgeValue = null; // Handle null case
                                }
                                //Study Id 
                                if (string.IsNullOrEmpty(registrationInfo.StudyId))
                                {
                                    registrationInfo.StudyId = row["StudyId"].ToString();
                                }
                                // Save changes
                                context.SaveChanges();
                                storeSystemRemarks = ConvertUpdatedRegistrationInfoDataToSystemRemarks(registrationInfo);
                                bool auditResult = SaveAuditTrail(row["ProjectCode"].ToString(), systemId, ModuleName, PageName, UserRemarks, storeSystemRemarks);
                                rowCount++;
                            }                   
                        else
                        {
                            Console.WriteLine($"TRF_RegistrationInfos not found for RequisitionId: {requisitionId}");
                            success = false; // Update not happened
                        }
                    }
                    catch (DbUpdateException ex)
                    {
                        Console.WriteLine($"DbUpdateException occurred while updating TRF_RegistrationInfos: {ex.InnerException?.Message}");
                        string fieldName = GetFieldNameCausingError(row);
                        DataRow errorRow = errorDetailsTable.NewRow();
                        errorRow["SheetName"] = "RegistrationUpdate";
                        errorRow["RowIndex"] = rowIndex;
                        errorRow["ErrorField"] = fieldName;
                        errorRow["ErrorDescription"] = ex.Message;
                        errorDetailsTable.Rows.Add(errorRow);
                        success = false; // Update not happened
                        throw;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"An error occurred while updating TRF_RegistrationInfos: {ex.Message}");
                        string fieldName = GetFieldNameCausingError(row);
                        DataRow errorRow = errorDetailsTable.NewRow();
                        errorRow["SheetName"] = "RegistrationUpdate";
                        errorRow["RowIndex"] = rowIndex;
                        errorRow["ErrorField"] = fieldName;
                        errorRow["ErrorDescription"] = ex.Message;
                        errorDetailsTable.Rows.Add(errorRow);
                        success = false; // Update not happened
                        continue;
                    }
                }
            }
            return (success, rowCount);
        }

        // Convert the updated data for TRF_registrationinfo - Help R
        static string ConvertUpdatedRegistrationInfoDataToSystemRemarks(TRF_RegistrationInfo data)
        {
            string updatedRegSystemRemarks = "PatientId" + "\t" + data.PatientId + "\t" + Environment.NewLine +
                                          "PatientInitials" + "\t" + data.PatientInitials + "\t" + Environment.NewLine +
                                          "VisitName" + "\t" + data.VisitName + "\t" + Environment.NewLine +
                                          "RequisitionId" + "\t" + data.RequisitionId + "\t" + Environment.NewLine +
                                          "SubjectId" + "\t" + data.SubjectId + "\t" + Environment.NewLine +
                                          "StudyId" + "\t" + data.StudyId + "\t" + Environment.NewLine +
                                          "PatientName" + "\t" + data.PatientName + "\t" + Environment.NewLine +
                                          "Gender" + "\t" + data.Gender + "\t" + Environment.NewLine +
                                          "CollectionDate" + "\t" + data.CollectionDate + "\t" + Environment.NewLine +
                                          "DateOfBirth" + "\t" + data.DateOfBirth + "\t" + Environment.NewLine +
                                          "ProcessingDate" + "\t" + data.ProcessingDate + "\t" + Environment.NewLine +
                                          //need to change for the raja version
                                          "ReceivedDateTime" + "\t" + data.ReceivedDateTime + "\t" + Environment.NewLine +
                                          "RegField3" + "\t" + data.RegField3 + "\t" + Environment.NewLine +
                                          "RegistrationStatus" + "\t" + data.RegistrationStatus + "\t" + Environment.NewLine +
                                          "RegistrationCheckStatus" + "\t" + data.RegistrationCheckStatus + "\t" + Environment.NewLine +
                                          "ResultCheckStatus" + "\t" + data.ResultCheckStatus + "\t" + Environment.NewLine +
                                          "ApprovedStatus" + "\t" + data.ApprovedStatus + "\t" + Environment.NewLine +
                                          "HoldStatus" + "\t" + data.HoldStatus + "\t" + Environment.NewLine +
                                          "IsLock" + "\t" + data.IsLock + "\t" + Environment.NewLine +
                                          "ReportGeneratedStatus" + "\t" + data.ReportGeneratedStatus + "\t" + Environment.NewLine +
                                          "CreatedOn" + "\t" + data.CreatedOn + "\t" + Environment.NewLine +
                                          "Height" + "\t" + data.Height + "\t" + Environment.NewLine +
                                          "Weight" + "\t" + data.Weight + "\t" + Environment.NewLine +
                                          "AgeValue" + "\t" + data.AgeValue + "\t" + Environment.NewLine;

            return updatedRegSystemRemarks;

            /*var values = new List<string>
                {
                    data.PatientId,
                    data.PatientInitials,
                    data.VisitName,
                    data.RequisitionId,
                    data.SubjectId,
                    data.StudyId,
                    data.PatientName,
                    data.Gender,
                    data.CollectionDate?.ToString("yyyy-mm-dd HH:mm:ss")??string.Empty,
                    data.DateOfBirth?.ToString("yyyy-mm-dd HH:mm:ss")??string.Empty,
                    data.ProcessingDate?.ToString("yyyy-mm-dd HH:mm:ss")??string.Empty,
                    data.RegField3,
                    data.RegistrationStatus.ToString(),
                    data.RegistrationCheckStatus.ToString(),
                    data.ResultCheckStatus.ToString(),
                    data.ApprovedStatus.ToString(),
                    data.HoldStatus.ToString(),
                    data.IsLock.ToString(),
                    data.ReportGeneratedStatus.ToString(),
                    data.CreatedOn?.ToString("yyyy-mm-dd HH:mm:ss")??string.Empty,
                    data.Height.ToString(),
                    data.Weight.ToString(),
                    data.AgeValue.ToString()
                };
                return string.Join(",", values);*/
        }

        // Update TRF_Reg_BarCodes - RB
        static int UpdateRegBarCodes(DataTable dataTable, DataTable errorDetailsTable)
        {
            int rowCount = 0;
            string ModuleName = "TRF_Reg_BarCodes";
            string PageName = "TRF_Reg_BarCodesUpdate";
            string UserRemarks = "Update TRF_Reg_BarCodesUpdate";
            using (var context = new LIMSDevContext())
            {
                foreach (DataRow row in dataTable.Rows)
                {
                    int rowIndex = dataTable.Rows.IndexOf(row);
                    try
                    {
                        string storeSystemRemarks = "";
                        string barCodeValue = row["BarCodeValue"].ToString();
                        var regBarCode = context.TRF_Reg_BarCodes.FirstOrDefault(r => r.BarCodeValue == barCodeValue);
                        row["ProjectCode"] = regBarCode?.ProjectCode;
                        long systemId = 0;
                        if (regBarCode?.SystemId != null && long.TryParse(regBarCode.SystemId.ToString(), out systemId))
                        {
                            row["SystemId"] = systemId;
                        }
                        else
                        {
                            row["SystemId"] = DBNull.Value; // or handle the conversion failure as needed
                        }

                        if (regBarCode != null)
                        {
                            regBarCode.RequisitionId = row["RequisitionId"].ToString();
                            regBarCode.BarCodeValue = row["BarCodeValue"].ToString();
                            regBarCode.Remarks = row["Remarks"].ToString();
                            regBarCode.ReceivedCondition = row["ReceivedCondition"].ToString();
                            regBarCode.Status = row["Status"].ToString();
                            regBarCode.CollectionDate = Convert.ToDateTime(row["CollectionDate"]);
                            regBarCode.CustomField1 = row["ClientTubeNo"].ToString();
                            regBarCode.CreatedDateTime = DateTime.Now;
                            context.SaveChanges();
                            storeSystemRemarks = ConvertUpdatedRegBarCodesDataToSystemRemarks(regBarCode);
                            bool auditResult = SaveAuditTrail(row["ProjectCode"].ToString(), systemId, ModuleName, PageName, UserRemarks, storeSystemRemarks);
                            rowCount++;
                        }
                        else
                        {
                            Console.WriteLine($"RegBarCode not found for BarCodeValue: {barCodeValue}");
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"An error occurred while updating RegBarCode: {ex.Message}");
                        string fieldName = GetFieldNameCausingError(row);
                        DataRow errorRow = errorDetailsTable.NewRow();
                        errorRow["SheetName"] = "RegBarCodesUpdate";
                        errorRow["RowIndex"] = rowIndex;
                        errorRow["ErrorField"] = fieldName;
                        errorRow["ErrorDescription"] = ex.Message;
                        errorDetailsTable.Rows.Add(errorRow);
                        continue;
                    }
                }
            }
            return rowCount;
        }

        // Convert the updated data for TRF_regbarcodes - Help RB
        static string ConvertUpdatedRegBarCodesDataToSystemRemarks(TRF_Reg_BarCode data)
        {
            string updatedRegSystemRemarks = "RequisitionId" + "\t" + data.RequisitionId + "\t" + Environment.NewLine +
                                          "BarCodeValue" + "\t" + data.BarCodeValue + "\t" + Environment.NewLine +
                                          "Remarks" + "\t" + data.Remarks + "\t" + Environment.NewLine +
                                          "ReceivedCondition" + "\t" + data.ReceivedCondition + "\t" + Environment.NewLine +
                                          "Status" + "\t" + data.Status + "\t" + Environment.NewLine +
                                          "CollectionDate" + "\t" + data.CollectionDate + "\t" + Environment.NewLine +
                                          "CustomField1" + "\t" + data.CustomField1 + "\t" + Environment.NewLine +
                                          "CreatedDateTime" + "\t" + data.CreatedDateTime + "\t" + Environment.NewLine;

            return updatedRegSystemRemarks;


            /*var values = new List<string>
            {
                data.RequisitionId,
                data.BarCodeValue,
                data.Remarks,
                data.ReceivedCondition,
                data.Status,
                data.CollectionDate?.ToString("yyyy-mm-dd HH:mm:ss")?? string.Empty,
                data.CustomField1,
                data.CreatedDateTime?.ToString("yyyy-mm-dd HH:mm:ss")?? string.Empty
            };
            return string.Join(",", values);*/
        }

        // Storage Insert/Update - S
        static void InsertOrUpdateSampleData(DataTable dataTable, DataTable errorDetailsTable)
        {
            int insertCount = 0;
            int updateCount = 0;
            //@@raja string Cryobox ="" get it from User
            string Cryobox = "Box2";
            string ModuleName = "Storage";
            string PageName = "StorageInsertOrUpdate";
            string UserRemarks = "Insert/Update Storage";
            using (var context = new LIMSDevContext())
            {
                foreach (DataRow row in dataTable.Rows)
                {
                    int rowIndex = dataTable.Rows.IndexOf(row);
                    try
                    {
                        string storeSystemRemarks = "";
                        string barCodeValue = row["BarCodeValue"].ToString();
                        var regBarCode = context.TRF_Reg_BarCodes.FirstOrDefault(r => r.BarCodeValue == barCodeValue);
                        row["ProjectCode"] = regBarCode?.ProjectCode;
                        // Add the harcode values for Location
                        row["Location"] = "Banglore";
                        long systemId = 0;
                        if (regBarCode?.SystemId != null && long.TryParse(regBarCode.SystemId.ToString(), out systemId))
                        {
                            row["SystemId"] = systemId;
                        }
                        else
                        {
                            row["SystemId"] = DBNull.Value; // or handle the conversion failure as needed
                        }
                        try
                        {
                            if (BarcodeExists(context, barCodeValue))
                            {
                                var existingData = context.StorageDetails.SingleOrDefault(d => d.BarCodeValue == barCodeValue);
                                if (existingData != null)
                                {
                                    UpdateSampleDataFromRow(existingData, row);
                                    if (CryoBoxExists(context, barCodeValue, Cryobox))
                                    {
                                        //Case 4: Existing BarCodeValue & Existing Cryobox
                                        existingData.CryoboxWellPosition = GetMaxCryoBoxWellPositions(context, Cryobox) + 1 ?? 1;
                                        storeSystemRemarks = ConvertUpdatedDataToSystemRemarks(existingData);
                                        bool auditResult = SaveAuditTrail(row["ProjectCode"].ToString(), systemId, ModuleName, PageName, UserRemarks, storeSystemRemarks);
                                    }
                                    else
                                    {
                                        //Case 3: Existing BarCodeValue & New Cryobox
                                        existingData.Cryobox = Cryobox;
                                        existingData.CryoboxWellPosition = GetMaxCryoBoxWellPositions(context, Cryobox) + 1 ?? 1;
                                        storeSystemRemarks = ConvertUpdatedDataToSystemRemarks(existingData);
                                        bool auditResult = SaveAuditTrail(row["ProjectCode"].ToString(), systemId, ModuleName, PageName, UserRemarks, storeSystemRemarks);
                                    }
                                    context.SaveChanges();
                                    updateCount++;
                                }
                            }
                            else
                            {
                                row["Status"] = "Stored";
                                var newData = CreateSampleDataFromRow(row);
                                try
                                {
                                    if (CryoBoxExistsOnly(context, Cryobox))
                                    {
                                        //Case 2: Inserted New BarCodeValue & Existing Cryobox
                                        newData.Cryobox = Cryobox;
                                        newData.CryoboxWellPosition = GetMaxCryoBoxWellPositions(context, Cryobox) + 1 ?? 1;
                                        storeSystemRemarks = ConvertInsertedDataToSystemRemarks(newData);
                                        bool auditResult = SaveAuditTrail(row["ProjectCode"].ToString(), systemId, ModuleName, PageName, UserRemarks, storeSystemRemarks);
                                    }
                                    else
                                    {
                                        //Case 1: Inserted New BarCodeValue & New Cryobox
                                        newData.CryoboxWellPosition = 1; // Set initial position for new CryoBox
                                        newData.Cryobox = Cryobox;
                                        storeSystemRemarks = ConvertInsertedDataToSystemRemarks(newData);
                                        bool auditResult = SaveAuditTrail(row["ProjectCode"].ToString(), systemId, ModuleName, PageName, UserRemarks, storeSystemRemarks);
                                    }
                                    context.StorageDetails.Add(newData);
                                    context.SaveChanges();
                                    insertCount++;
                                }
                                catch (DbUpdateException ex)
                                {
                                    Console.WriteLine("DbUpdateException: " + ex.Message);
                                    string fieldName = GetFieldNameCausingError(row);
                                    DataRow errorRow = errorDetailsTable.NewRow();
                                    errorRow["SheetName"] = "StorageInsertUpdate";
                                    errorRow["RowIndex"] = rowIndex;
                                    errorRow["ErrorField"] = fieldName;
                                    errorRow["ErrorDescription"] = ex.Message;
                                    errorDetailsTable.Rows.Add(errorRow);
                                    throw;
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Inner Catch : Error processing row : {ex.Message}");
                            string fieldName = GetFieldNameCausingError(row);
                            DataRow errorRow = errorDetailsTable.NewRow();
                            errorRow["SheetName"] = "StorageInsertUpdate";
                            errorRow["RowIndex"] = rowIndex;
                            errorRow["ErrorField"] = fieldName;
                            errorRow["ErrorDescription"] = ex.Message;
                            errorDetailsTable.Rows.Add(errorRow);
                            continue;
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Outer Catch : Error processing row: {ex.Message}");
                        string fieldName = GetFieldNameCausingError(row);
                        DataRow errorRow = errorDetailsTable.NewRow();
                        errorRow["SheetName"] = "StorageInsertUpdate";
                        errorRow["RowIndex"] = rowIndex;
                        errorRow["ErrorField"] = fieldName;
                        errorRow["ErrorDescription"] = ex.Message;
                        errorDetailsTable.Rows.Add(errorRow);
                        continue;
                    }
                }
                Console.WriteLine($"Total No Of records Successfully Inserted - {insertCount} ");
                Console.WriteLine($"Total No Of records Successfully Updated - {updateCount} ");
            }
        }

        // Helper function for Storage Insert/Update - Help S1
        static StorageDetail CreateSampleDataFromRow(DataRow row)
        {
            StorageDetail newData = new StorageDetail();
            newData.PatientId = GetValueAsString(row, "PatientId");
            newData.SiteNo = GetValueAsString(row, "SiteNo");
            newData.VisitName = GetValueAsString(row, "VisitName");
            newData.DOBSCollection = row.Field<DateTime?>("DOBSCollection");
            newData.TOBSCollection = row.Field<TimeSpan?>("TOBSCollection");
            newData.RequisitionId = GetValueAsString(row, "RequisitionId");
            newData.BarCodeValue = GetValueAsString(row, "BarCodeValue");
            newData.TypeOfSample = GetValueAsString(row, "TypeOfSample");
            newData.Remarks = GetValueAsString(row, "Remarks");
            newData.ReceivedCondition = GetValueAsString(row, "ReceivedCondition");
            newData.ProjectCode = GetValueAsString(row, "ProjectCode");
            newData.Status = GetValueAsString(row, "Status");
            newData.Location = GetValueAsString(row, "Location");
            return newData;
        }

        // Helper function for Storage Insert/Update - Help S2
        private static void UpdateSampleDataFromRow(StorageDetail updateData, DataRow row)
        {
            updateData.PatientId = GetValueAsString(row, "PatientId");
            updateData.SiteNo = GetValueAsString(row, "SiteNo");
            updateData.VisitName = GetValueAsString(row, "VisitName");
            updateData.DOBSCollection = row.Field<DateTime?>("DOBSCollection");
            updateData.TOBSCollection = row.Field<TimeSpan?>("TOBSCollection");
            updateData.RequisitionId = GetValueAsString(row, "RequisitionId");
            updateData.BarCodeValue = GetValueAsString(row, "BarCodeValue");
            updateData.TypeOfSample = GetValueAsString(row, "TypeOfSample");
            updateData.Remarks = GetValueAsString(row, "Remarks");
            updateData.ReceivedCondition = GetValueAsString(row, "ReceivedCondition");
            updateData.ProjectCode = GetValueAsString(row, "ProjectCode");
            updateData.Location = GetValueAsString(row, "Location");
        }

        // Helper function for Storage Insert/Update - Help S3
        static string GetValueAsString(DataRow row, string columnName)
        {
            try
            {
                return row[columnName]?.ToString();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error setting {columnName}: {ex.Message}");
                return null;
            }
        }

        // Helper function for Storage Insert/Update - Help S4
        static bool BarcodeExists(LIMSDevContext context, string barCodeValue)
        {
            return context.StorageDetails.Count(sd => sd.BarCodeValue == barCodeValue) > 0;
        }

        // Helper function for Storage Insert/Update - Help S5
        static bool CryoBoxExists(LIMSDevContext context, string barCodeValue, string cryoBoxValue)
        {
            return context.StorageDetails.Count(sd => sd.BarCodeValue == barCodeValue && sd.Cryobox == cryoBoxValue) > 0;
        }

        // Helper function for Storage Insert/Update - Help S6
        static int? GetMaxCryoBoxWellPositions(LIMSDevContext context, string cryoBoxValue)
        {
            return context.StorageDetails.Where(sd => sd.Cryobox == cryoBoxValue).Max(sd => (int?)sd.CryoboxWellPosition);
        }

        // Helper function for Storage Insert/Update - Help S7
        static bool CryoBoxExistsOnly(LIMSDevContext context, string cryoBoxValue)
        {
            return context.StorageDetails.Count(sd => sd.Cryobox == cryoBoxValue) > 0;
        }

        // Helper function for Storage Insert/Update - Help S8
        static string ConvertInsertedDataToSystemRemarks(StorageDetail data)
        {
            string insertedStorageSystemRemarks = "PatientId" + "\t" + data.PatientId + "\t" + Environment.NewLine +
                                                  "SiteNo" + "\t" + data.SiteNo + "\t" + Environment.NewLine +
                                                  "VisitName" + "\t" + data.VisitName + "\t" + Environment.NewLine +
                                                  "DOBSCollection" + "\t" + data.DOBSCollection + "\t" + Environment.NewLine +
                                                  "TOBSCollection" + "\t" + data.TOBSCollection + "\t" + Environment.NewLine +
                                                  "RequisitionId" + "\t" + data.RequisitionId + "\t" + Environment.NewLine +
                                                  "BarCodeValue" + "\t" + data.BarCodeValue + "\t" + Environment.NewLine +
                                                  "TypeOfSample" + "\t" + data.TypeOfSample + "\t" + Environment.NewLine +
                                                  "Remarks" + "\t" + data.Remarks + "\t" + Environment.NewLine +
                                                  "ReceivedCondition" + "\t" + data.ReceivedCondition + "\t" + Environment.NewLine +
                                                  "Cryobox" + "\t" + data.Cryobox + "\t" + Environment.NewLine +
                                                  "CryoboxWellPosition" + "\t" + data.CryoboxWellPosition + "\t" + Environment.NewLine +
                                                  "ProjectCode" + "\t" + data.ProjectCode + "\t" + Environment.NewLine +
                                                  "Status" + "\t" + data.Status + "\t" + Environment.NewLine +
                                                  "Location" + "\t" + data.Location + "\t" + Environment.NewLine;

            return insertedStorageSystemRemarks;

            /*var values = new List<string>
            {
                data.PatientId,
                data.SiteNo,
                data.VisitName,
                data.DOBSCollection?.ToString("yyyy-mm-dd HH:mm:ss")?? string.Empty,
                data.TOBSCollection?.ToString(@"hh\:mm\:ss")?? string.Empty,
                data.RequisitionId,
                data.BarCodeValue,
                data.TypeOfSample,
                data.Remarks,
                data.ReceivedCondition,
                data.Cryobox,
                data.CryoboxWellPosition?.ToString()?? string.Empty,
                data.ProjectCode,
                data.Status,
                data.Location
            };
            return string.Join(",", values);*/
        }

        // Helper function for Storage Insert/Update - Help S9
        static string ConvertUpdatedDataToSystemRemarks(StorageDetail data)
        {
            string updatedStorageSystemRemarks = "PatientId" + "\t" + data.PatientId + "\t" + Environment.NewLine +
                                                  "SiteNo" + "\t" + data.SiteNo + "\t" + Environment.NewLine +
                                                  "VisitName" + "\t" + data.VisitName + "\t" + Environment.NewLine +
                                                  "DOBSCollection" + "\t" + data.DOBSCollection + "\t" + Environment.NewLine +
                                                  "TOBSCollection" + "\t" + data.TOBSCollection + "\t" + Environment.NewLine +
                                                  "RequisitionId" + "\t" + data.RequisitionId + "\t" + Environment.NewLine +
                                                  "BarCodeValue" + "\t" + data.BarCodeValue + "\t" + Environment.NewLine +
                                                  "TypeOfSample" + "\t" + data.TypeOfSample + "\t" + Environment.NewLine +
                                                  "Remarks" + "\t" + data.Remarks + "\t" + Environment.NewLine +
                                                  "ReceivedCondition" + "\t" + data.ReceivedCondition + "\t" + Environment.NewLine +
                                                  "Cryobox" + "\t" + data.Cryobox + "\t" + Environment.NewLine +
                                                  "CryoboxWellPosition" + "\t" + data.CryoboxWellPosition + "\t" + Environment.NewLine +
                                                  "ProjectCode" + "\t" + data.ProjectCode + "\t" + Environment.NewLine +
                                                  "Location" + "\t" + data.Location + "\t" + Environment.NewLine;

            return updatedStorageSystemRemarks;
            
            /*var values = new List<string>
            {
                data.PatientId,
                data.SiteNo,
                data.VisitName,
                data.DOBSCollection?.ToString("yyyy-mm-dd HH:mm:ss")?? string.Empty,
                data.TOBSCollection?.ToString(@"hh\:mm\:ss")?? string.Empty,
                data.RequisitionId,
                data.BarCodeValue,
                data.TypeOfSample,
                data.Remarks,
                data.ReceivedCondition,
                data.Cryobox,
                data.CryoboxWellPosition?.ToString()?? string.Empty,
                data.ProjectCode,
                data.Location
            };
            return string.Join(",", values);*/
        }

        // Helper function for SystemId Insert/Update - Help S10
        static long GetValueAsLong(DataRow row, string columnName)
        {
            try
            {
                if(row[columnName] == DBNull.Value)
                {
                    return 0;
                }
                return Convert.ToInt64(row[columnName]);
            }
            catch (InvalidCastException)
            {
                return 0; 
            }
            catch (Exception)
            {
                return 0; 
            }
        }

        // AuditHistory log function
        static bool SaveAuditTrail(string ProjectCode, long TransId, string ModuleName, string PageName, string UserRemarks, string storedSystemRemarks)
        {
            bool auditHistoryResult = false;
            using (var context = new LIMSDevContext())
            {
                try
                {
                    AuditHistory auditHistoryObj = new AuditHistory
                    {
                        LoginId = "Krishna H K", // For testing
                        ProjectCode = ProjectCode,
                        TransId = TransId,
                        ModuleName = ModuleName,
                        PageName = PageName,
                        UserRemarks = UserRemarks,
                        SystemRemarks = storedSystemRemarks,
                        CreatedDateTime = DateTime.Now
                };       
                    context.AuditHistorys.Add(auditHistoryObj);
                    int result = context.SaveChanges();
                    if (result > 0)
                    {
                        auditHistoryResult = true;
                    }
                    else
                    {
                        Console.WriteLine("No changes were saved to the database.");
                    }
                }
                catch (DbEntityValidationException dbEx)
                {
                    foreach (var validationErrors in dbEx.EntityValidationErrors)
                    {
                        foreach (var validationError in validationErrors.ValidationErrors)
                        {
                            Console.WriteLine($"Property: {validationError.PropertyName} Error: {validationError.ErrorMessage}");
                        }
                    }
                }
                catch (DbUpdateException dbUpEx)
                {
                    var innerException = dbUpEx.InnerException?.InnerException;
                    Console.WriteLine($"DbUpdateException error: {dbUpEx.Message}");
                    if (innerException != null)
                    {
                        Console.WriteLine($"Inner exception: {innerException.Message}");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"An error occurred while saving changes: {ex.Message}");
                }
            }
            return auditHistoryResult;
        }

        // Get the Column name for the ErrorTable 
        static string GetFieldNameCausingError(DataRow row)
        {
            foreach (DataColumn column in row.Table.Columns)
            {
                try
                {
                    var value = row[column.ColumnName]; // Access the column value
                }
                catch (Exception)
                {
                    return column.ColumnName; // Return the column name causing the error
                }
            }
            return string.Empty; // Return an empty string if no error found
        }
    }
}