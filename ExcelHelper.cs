using System;
using System.Data;
using System.IO;
using System.Collections.Generic;
using ClosedXML.Excel;

namespace HostelManagementSystem
{
    public class ExcelHelper : IDisposable
    {
        private readonly string _filePath;
        private bool _disposed = false;

        public ExcelHelper(string filePath)
        {
            _filePath = filePath ?? throw new ArgumentNullException(nameof(filePath));

            if (!File.Exists(_filePath))
            {
                CreateInitialExcelFile();
            }
            else if (!IsValidExcelFile())
            {
                File.Delete(_filePath);
                CreateInitialExcelFile();
            }
        }

        private void CreateInitialExcelFile()
        {
            DataTable staffTable = CreateStaffDataTable();
            DataTable studentsTable = CreateStudentsTable();
            DataTable complaintsTable = CreateComplaintsDataTable();
            DataTable roomsTable = CreateRoomsDataTable();
            DataTable feesTable = CreateFeesDataTable();

            CreateExcelFile(new DataTable[] { staffTable, studentsTable, complaintsTable, roomsTable, feesTable });
        }

        private DataTable CreateStaffDataTable()
        {
            DataTable table = new DataTable("Staff");
            table.Columns.Add("ID", typeof(string));
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("Email", typeof(string));
            table.Columns.Add("Password", typeof(string));
            table.Columns.Add("Position", typeof(string));

            DataRow row = table.NewRow();
            row["ID"] = "ADM-001";
            row["Name"] = "System Admin";
            row["Email"] = "admin@hostel.com";
            row["Password"] = "admin123";
            row["Position"] = "Admin";
            table.Rows.Add(row);

            return table;
        }

        private DataTable CreateStudentsTable()
        {
            DataTable table = new DataTable("Student");
           // DataTable table = new DataTable(StudentsSheet);
            table.Columns.Add("ID", typeof(string));
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("Email", typeof(string));
            table.Columns.Add("Password", typeof(string));
            table.Columns.Add("RoomNumber", typeof(string));
            table.Columns.Add("FeesPaid", typeof(decimal)); // Make sure this column exists
            table.Columns.Add("JoinDate", typeof(DateTime));
            //ExcelHelper.CreateExcelFile(new[] { table });
            return table;
        }

        private DataTable CreateComplaintsDataTable()
        {
            DataTable table = new DataTable("Complaints");
            table.Columns.Add("ID", typeof(string));
            table.Columns.Add("StudentID", typeof(string));
            table.Columns.Add("Description", typeof(string));
            table.Columns.Add("Status", typeof(string));
            table.Columns.Add("Date", typeof(DateTime));
            return table;
        }

        private DataTable CreateRoomsDataTable()
        {
            DataTable table = new DataTable("Rooms");
            table.Columns.Add("RoomNumber", typeof(int));
            table.Columns.Add("Capacity", typeof(int));
            table.Columns.Add("Occupied", typeof(int));

            for (int i = 1; i <= 8; i++)
            {
                DataRow row = table.NewRow();
                row["RoomNumber"] = i;
                row["Capacity"] = 4;
                row["Occupied"] = 0;
                table.Rows.Add(row);
            }

            return table;
        }

        private DataTable CreateFeesDataTable()
        {
            DataTable table = new DataTable("Fees");
            table.Columns.Add("StudentID", typeof(string));
            table.Columns.Add("Amount", typeof(decimal));
            table.Columns.Add("PaymentDate", typeof(DateTime));
            table.Columns.Add("DueDate", typeof(DateTime));
            return table;
        }

        private bool IsValidExcelFile()
        {
            try
            {
                using (var workbook = new XLWorkbook(_filePath))
                {
                    return workbook.Worksheets.Count > 0;
                }
            }
            catch
            {
                return false;
            }
        }

        public List<string> GetSheetNames()
        {
            try
            {
                using (var workbook = new XLWorkbook(_filePath))
                {
                    var sheetNames = new List<string>();
                    foreach (var worksheet in workbook.Worksheets)
                    {
                        sheetNames.Add(worksheet.Name);
                    }
                    return sheetNames;
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Error getting sheet names: " + ex.Message, ex);
            }
        }

        public bool SheetExists(string sheetName)
        {
            try
            {
                using (var workbook = new XLWorkbook(_filePath))
                {
                    return workbook.Worksheets.TryGetWorksheet(sheetName, out _);
                }
            }
            catch
            {
                return false;
            }
        }

        public void CreateExcelFile(DataTable[] dataTables)
        {
            try
            {
                using (var workbook = new XLWorkbook())
                {
                    foreach (DataTable table in dataTables)
                    {
                        var worksheet = workbook.Worksheets.Add(table);
                        worksheet.Name = table.TableName;
                    }
                    workbook.SaveAs(_filePath);
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Error creating Excel file: " + ex.Message, ex);
            }
        }

        public DataTable ReadData(string sheetName)
        {
            try
            {
                using (var workbook = new XLWorkbook(_filePath))
                {
                    var worksheet = workbook.Worksheet(sheetName);
                    var table = new DataTable(worksheet.Name);

                    // Read header
                    var firstRow = worksheet.FirstRowUsed();
                    foreach (var cell in firstRow.Cells())
                    {
                        table.Columns.Add(cell.Value.ToString());
                    }

                    // Read data
                    var rows = worksheet.RowsUsed().Skip(1);
                    foreach (var row in rows)
                    {
                        DataRow dataRow = table.NewRow();
                        for (int i = 0; i < table.Columns.Count; i++)
                        {
                            dataRow[i] = row.Cell(i + 1).Value;
                        }
                        table.Rows.Add(dataRow);
                    }

                    return table;
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error reading data from sheet '{sheetName}': {ex.Message}", ex);
            }
        }

        public void WriteData(string sheetName, DataTable dataTable)
        {
            try
            {
                using (var workbook = new XLWorkbook(_filePath))
                {
                    if (workbook.Worksheets.TryGetWorksheet(sheetName, out var existingWorksheet))
                    {
                        existingWorksheet.Delete();
                    }

                    workbook.Worksheets.Add(dataTable, sheetName);
                    workbook.Save();
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error writing data to sheet '{sheetName}': {ex.Message}", ex);
            }
        }

        public void AppendData(string sheetName, DataRow newRow)
        {
            try
            {
                var table = ReadData(sheetName);
                table.Rows.Add(newRow.ItemArray);
                WriteData(sheetName, table);
            }
            catch (Exception ex)
            {
                throw new Exception($"Error appending data to sheet '{sheetName}': {ex.Message}", ex);
            }
        }

        public void UpdateRow(string sheetName, string idColumn, string idValue, Dictionary<string, string> updates)
        {
            try
            {
                var table = ReadData(sheetName);
                foreach (DataRow row in table.Rows)
                {
                    if (row[idColumn].ToString() == idValue)
                    {
                        foreach (var update in updates)
                        {
                            row[update.Key] = update.Value;
                        }
                        break;
                    }
                }
                WriteData(sheetName, table);
            }
            catch (Exception ex)
            {
                throw new Exception($"Error updating row in sheet '{sheetName}': {ex.Message}", ex);
            }
        }

        public void DeleteRow(string sheetName, string idColumn, string idValue)
        {
            try
            {
                var table = ReadData(sheetName);
                for (int i = table.Rows.Count - 1; i >= 0; i--)
                {
                    if (table.Rows[i][idColumn].ToString() == idValue)
                    {
                        table.Rows.RemoveAt(i);
                        break;
                    }
                }
                WriteData(sheetName, table);
            }
            catch (Exception ex)
            {
                throw new Exception($"Error deleting row from sheet '{sheetName}': {ex.Message}", ex);
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    // Dispose managed resources if needed
                }
                _disposed = true;
            }
        }

        ~ExcelHelper()
        {
            Dispose(false);
        }
    }
}