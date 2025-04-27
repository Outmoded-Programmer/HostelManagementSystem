using System;
using System.Data;
using System.Linq;

namespace HostelManagementSystem
{
    public class Admin : Staff, IAdminOperations
    {
        private readonly IRoomManager _roomManager;
        private const string StudentsSheet = "Students";
        private const string StaffSheet = "Staff";
        private const string ComplaintsSheet = "Complaints";

        public Admin(string id, string name, string email, string password, ExcelHelper excelHelper)
            : base(id, name, email, password, excelHelper)
        {
            Position = "Admin";
            _roomManager = new HostelManager(excelHelper);
            EnsureTablesExist();
        }

        private void EnsureTablesExist()
        {
            try
            {
                if (!ExcelHelper.SheetExists(StudentsSheet))
                    CreateStudentsTable();

                if (!ExcelHelper.SheetExists(StaffSheet))
                    CreateStaffTable();

                if (!ExcelHelper.SheetExists(ComplaintsSheet))
                    CreateComplaintsTable();

                if (!ExcelHelper.SheetExists("Rooms"))
                    _roomManager.InitializeRooms();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error initializing tables: {ex.Message}");
            }
        }

        private void CreateStudentsTable()
        {
            var table = new DataTable(StudentsSheet);
            table.Columns.Add("ID", typeof(string));
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("Email", typeof(string));
            table.Columns.Add("Password", typeof(string));
            table.Columns.Add("RoomNumber", typeof(string));
            table.Columns.Add("FeesPaid", typeof(decimal));
            table.Columns.Add("JoinDate", typeof(DateTime));
            ExcelHelper.CreateExcelFile(new[] { table });
        }

        private void CreateStaffTable()
        {
            var table = new DataTable(StaffSheet);
            table.Columns.Add("ID", typeof(string));
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("Email", typeof(string));
            table.Columns.Add("Password", typeof(string));
            table.Columns.Add("Position", typeof(string));
            table.Columns.Add("JoinDate", typeof(DateTime));
            ExcelHelper.CreateExcelFile(new[] { table });
        }

        private void CreateComplaintsTable()
        {
            var table = new DataTable(ComplaintsSheet);
            table.Columns.Add("ID", typeof(string));
            table.Columns.Add("StudentID", typeof(string));
            table.Columns.Add("Description", typeof(string));
            table.Columns.Add("Status", typeof(string));
            table.Columns.Add("Date", typeof(DateTime));
            table.Columns.Add("Resolution", typeof(string));
            ExcelHelper.CreateExcelFile(new[] { table });
        }

        public void AddStudent(int id, string name, string email, string password)
        {
            try
            {
                DataTable studentsTable = ExcelHelper.ReadData(StudentsSheet);

                // Verify the table has all required columns
                if (!studentsTable.Columns.Contains("FeesPaid"))
                {
                    // Recreate the table if structure is incorrect
                    CreateStudentsTable();
                    studentsTable = ExcelHelper.ReadData(StudentsSheet);
                }

                DataRow newStudent = studentsTable.NewRow();
                newStudent["ID"] = id;
                newStudent["Name"] = name;
                newStudent["Email"] = email;
                newStudent["Password"] = password;
                newStudent["RoomNumber"] = "0"; // Unassigned
                newStudent["FeesPaid"] = 0m; // Initial fee
                newStudent["JoinDate"] = DateTime.Now;

                studentsTable.Rows.Add(newStudent);
                ExcelHelper.WriteData(StudentsSheet, studentsTable);
                Console.WriteLine($"\nStudent added successfully!\nID: {newStudent["ID"]}\nName: {name}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"\nError adding student: {ex.Message}");
                // Attempt to recreate the table if the error persists
                try
                {
                    CreateStudentsTable();
                    Console.WriteLine("Students table has been recreated. Please try adding the student again.");
                }
                catch (Exception recreateEx)
                {
                    Console.WriteLine($"\nFailed to recreate table: {recreateEx.Message}");
                }
            }
        }
        public void ViewAllStudents()
        {
            try
            {
                var studentsTable = ExcelHelper.ReadData(StudentsSheet);
                if (studentsTable.Rows.Count == 0)
                {
                    Console.WriteLine("No students found.");
                    return;
                }

                Console.WriteLine("\nStudent List:");
                Console.WriteLine("ID\tName\t\tEmail\t\tRoom\tFees");
                foreach (DataRow student in studentsTable.Rows)
                {
                    Console.WriteLine($"{student["ID"]}\t{student["Name"]}\t{student["Email"]}\t{student["RoomNumber"]}\t{student["FeesPaid"]}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error viewing students: {ex.Message}");
            }
        }
        //public void UpdateStudent(string studentId , string studentName , string studentEmail, string studentPassword)
        public void UpdateStudent(string studentId, string studentName = null, string studentEmail = null, string studentPassword = null)
        {
            var studentsTable = ExcelHelper.ReadData(StudentsSheet);
            var student = studentsTable.AsEnumerable().FirstOrDefault(row => row["ID"].ToString() == studentId);

            if (studentName != null) student["Name"] = studentName;
            if (studentEmail != null) student["Email"] = studentEmail;
            if (studentPassword != null) student["Password"] = studentPassword;

            ExcelHelper.WriteData(StudentsSheet, studentsTable);
        }


        public void AssignRoom(string studentId, int roomNumber)
        {
            try
            {
                var studentsTable = ExcelHelper.ReadData(StudentsSheet);
                var student = studentsTable.AsEnumerable()
                    .FirstOrDefault(row => row["ID"].ToString() == studentId);

                if (student == null)
                {
                    Console.WriteLine("Student not found!");
                    return;
                }

                if (_roomManager.AssignStudentToRoom(studentId, roomNumber))
                {
                    student["RoomNumber"] = roomNumber.ToString();
                    ExcelHelper.WriteData(StudentsSheet, studentsTable);
                    Console.WriteLine($"Room {roomNumber} assigned to student {studentId}");
                }
                else
                {
                    Console.WriteLine("Room assignment failed. Room may be full or doesn't exist.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error assigning room: {ex.Message}");
            }
        }

        public void RemoveStudent(string studentId)
        {
            try
            {
                var studentsTable = ExcelHelper.ReadData(StudentsSheet);
                var student = studentsTable.AsEnumerable()
                    .FirstOrDefault(row => row["ID"].ToString() == studentId);

                if (student == null)
                {
                    Console.WriteLine("Student not found!");
                    return;
                }

                if (student["RoomNumber"].ToString() != "0" && int.TryParse(student["RoomNumber"].ToString(), out int roomNum))
                {
                    _roomManager.RemoveStudentFromRoom(studentId, roomNum);
                }

                // In the AssignRoom method, add null checks:
                var currentRoom = student["RoomNumber"].ToString();
                if (!string.IsNullOrEmpty(currentRoom) && currentRoom != "0" && int.TryParse(currentRoom, out int currentRoomNum))
                {
                    _roomManager.RemoveStudentFromRoom(studentId, currentRoomNum);
                }

                studentsTable.Rows.Remove(student);
                ExcelHelper.WriteData(StudentsSheet, studentsTable);
                Console.WriteLine("Student removed successfully!");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error removing student: {ex.Message}");
            }
        }

        public void ViewRoomStatus()
        {
            _roomManager.DisplayRoomStatus();
        }

        public void ViewAllComplaints()
        {
            try
            {
                var complaintsTable = ExcelHelper.ReadData(ComplaintsSheet);
                if (complaintsTable.Rows.Count == 0)
                {
                    Console.WriteLine("No complaints found.");
                    return;
                }

                Console.WriteLine("\nComplaint List:");
                Console.WriteLine("ID\tStudentID\tStatus\tDate");
                foreach (DataRow complaint in complaintsTable.Rows)
                {
                    Console.WriteLine($"{complaint["ID"]}\t{complaint["StudentID"]}\t{complaint["Status"]}\t{complaint["Date"]}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error viewing complaints: {ex.Message}");
            }
        }

        public void UpdateComplaintStatus(string complaintId, string newStatus)
        {
            try
            {
                var complaintsTable = ExcelHelper.ReadData(ComplaintsSheet);
                var complaint = complaintsTable.AsEnumerable()
                    .FirstOrDefault(row => row["ID"].ToString() == complaintId);

                if (complaint == null)
                {
                    Console.WriteLine("Complaint not found!");
                    return;
                }

                complaint["Status"] = newStatus;
                ExcelHelper.WriteData(ComplaintsSheet, complaintsTable);
                Console.WriteLine("Complaint status updated successfully!");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error updating complaint: {ex.Message}");
            }
        }

        public void AddStaffManager(string name, string email, string password)
        {
            AddStaff(name, email, password, "Manager");
        }

        public void ViewAllStaff()
        {
            try
            {
                var staffTable = ExcelHelper.ReadData(StaffSheet);
                if (staffTable.Rows.Count == 0)
                {
                    Console.WriteLine("No staff members found.");
                    return;
                }

                Console.WriteLine("\nStaff List:");
                Console.WriteLine("ID\tName\t\tPosition");
                foreach (DataRow staff in staffTable.Rows)
                {
                    Console.WriteLine($"{staff["ID"]}\t{staff["Name"]}\t{staff["Position"]}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error viewing staff: {ex.Message}");
            }
        }

        public void AddStaff(string name, string email, string password, string position)
        {
            try
            {
                var staffTable = ExcelHelper.ReadData(StaffSheet);
                var newStaff = staffTable.NewRow();
                newStaff["ID"] = $"{position.Substring(0, 3).ToUpper()}-{(staffTable.Rows.Count + 1):D3}";
                newStaff["Name"] = name;
                newStaff["Email"] = email;
                newStaff["Password"] = password;
                newStaff["Position"] = position;
                newStaff["JoinDate"] = DateTime.Now;

                staffTable.Rows.Add(newStaff);
                ExcelHelper.WriteData(StaffSheet, staffTable);
                Console.WriteLine($"{position} added successfully! ID: {newStaff["ID"]}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error adding staff: {ex.Message}");
            }
        }

        public void RemoveStaff(string staffId)
        {
            try
            {
                var staffTable = ExcelHelper.ReadData(StaffSheet);
                var staff = staffTable.AsEnumerable()
                    .FirstOrDefault(row => row["ID"].ToString() == staffId);

                if (staff == null)
                {
                    Console.WriteLine("Staff member not found!");
                    return;
                }

                if (staff["Position"].ToString() == "Admin" &&
                    staffTable.AsEnumerable().Count(row => row["Position"].ToString() == "Admin") <= 1)
                {
                    Console.WriteLine("Cannot remove the last admin!");
                    return;
                }

                staffTable.Rows.Remove(staff);
                ExcelHelper.WriteData(StaffSheet, staffTable);
                Console.WriteLine("Staff member removed successfully!");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error removing staff: {ex.Message}");
            }
        }

        public void CollectFees(string studentId, decimal amount)
        {
            try
            {
                var studentsTable = ExcelHelper.ReadData(StudentsSheet);
                var student = studentsTable.AsEnumerable()
                    .FirstOrDefault(row => row["ID"].ToString() == studentId);

                if (student == null)
                {
                    Console.WriteLine("Student not found!");
                    return;
                }

                student["FeesPaid"] = Convert.ToDecimal(student["FeesPaid"]) + amount;
                ExcelHelper.WriteData(StudentsSheet, studentsTable);
                Console.WriteLine($"Successfully collected {amount:C} from student {studentId}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error collecting fees: {ex.Message}");
            }
        }
    }
}
