using System;
using System.Data;
using System.Linq;
using System.IO;
using ExcelDataReader;
using System.Text;

namespace HostelManagementSystem
{
    class Program
    {
        private static ExcelHelper _excelHelper;
        private static Admin _admin;
        private static StaffManager _staffManager;
        private static Staff _staff;
        private static Student _student;
        private const string DataFileName = "HostelData.xlsx";

        static void Main(string[] args)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            Console.WriteLine("Hostel Management System");
            Console.WriteLine("========================");

            InitializeSystem();

            while (true)
            {
                Console.WriteLine("\nMain Menu");
                Console.WriteLine("1. Login");
                Console.WriteLine("2. Exit");
                Console.Write("Enter choice: ");

                string input = Console.ReadLine();

                switch (input)
                {
                    case "1":
                        LoginMenu();
                        break;
                    case "2":
                        Console.WriteLine("Exiting system...");
                        return;
                    default:
                        Console.WriteLine("Invalid choice. Please try again.");
                        break;
                }
            }
        }

        private static void InitializeSystem()
        {
            if (!File.Exists(DataFileName))
            {
                CreateInitialExcelFile();
                Console.WriteLine("Created new database file with all required sheets.");
            }

            _excelHelper = new ExcelHelper(DataFileName);
            InitializeDefaultAccounts();
        }

        private static void CreateInitialExcelFile()
        {
            DataTable staffTable = new DataTable("Staff");
            staffTable.Columns.Add("ID", typeof(string));
            staffTable.Columns.Add("Name", typeof(string));
            staffTable.Columns.Add("Email", typeof(string));
            staffTable.Columns.Add("Password", typeof(string));
            staffTable.Columns.Add("Position", typeof(string));
            staffTable.Columns.Add("JoinDate", typeof(DateTime));

            DataTable studentsTable = new DataTable("Students");
            studentsTable.Columns.Add("ID", typeof(string));
            studentsTable.Columns.Add("Name", typeof(string));
            studentsTable.Columns.Add("Email", typeof(string));
            studentsTable.Columns.Add("Password", typeof(string));
            studentsTable.Columns.Add("RoomNumber", typeof(string));
            studentsTable.Columns.Add("FeesPaid", typeof(decimal));
            studentsTable.Columns.Add("JoinDate", typeof(DateTime));

            DataTable complaintsTable = new DataTable("Complaints");
            complaintsTable.Columns.Add("ID", typeof(string));
            complaintsTable.Columns.Add("StudentID", typeof(string));
            complaintsTable.Columns.Add("Description", typeof(string));
            complaintsTable.Columns.Add("Status", typeof(string));
            complaintsTable.Columns.Add("Date", typeof(DateTime));

            DataTable roomsTable = new DataTable("Rooms");
            roomsTable.Columns.Add("RoomNumber", typeof(int));
            roomsTable.Columns.Add("Capacity", typeof(int));
            roomsTable.Columns.Add("Occupied", typeof(int));

            // Initialize rooms
            for (int i = 1; i <= 20; i++)
            {
                DataRow room = roomsTable.NewRow();
                room["RoomNumber"] = i;
                room["Capacity"] = 4;
                room["Occupied"] = 0;
                roomsTable.Rows.Add(room);
            }

            ExcelHelper excelHelper = new ExcelHelper(DataFileName);
            excelHelper.CreateExcelFile(new[] { staffTable, studentsTable, complaintsTable, roomsTable });
        }

        private static void InitializeDefaultAccounts()
        {
            DataTable staffTable = _excelHelper.ReadData("Staff");

            // Create default admin if doesn't exist
            if (!staffTable.AsEnumerable().Any(row => row["Position"].ToString() == "Admin"))
            {
                DataRow admin = staffTable.NewRow();
                admin["ID"] = "ADM-001";
                admin["Name"] = "System Admin";
                admin["Email"] = "admin@hostel.com";
                admin["Password"] = "admin123";
                admin["Position"] = "Admin";
                admin["JoinDate"] = DateTime.Now;
                staffTable.Rows.Add(admin);
                Console.WriteLine("\nDefault admin account created (admin@hostel.com / admin123)");
            }

            // Create default manager if doesn't exist
            if (!staffTable.AsEnumerable().Any(row => row["Position"].ToString() == "Manager"))
            {
                DataRow manager = staffTable.NewRow();
                manager["ID"] = "MGR-001";
                manager["Name"] = "Default Manager";
                manager["Email"] = "manager@hostel.com";
                manager["Password"] = "manager123";
                manager["Position"] = "Manager";
                manager["JoinDate"] = DateTime.Now;
                staffTable.Rows.Add(manager);
                Console.WriteLine("Default manager account created (manager@hostel.com / manager123)");
            }

            // Create default staff if doesn't exist
            if (!staffTable.AsEnumerable().Any(row => row["Position"].ToString() == "Staff"))
            {
                DataRow staff = staffTable.NewRow();
                staff["ID"] = "STF-001";
                staff["Name"] = "Default Staff";
                staff["Email"] = "staff@hostel.com";
                staff["Password"] = "staff123";
                staff["Position"] = "Staff";
                staff["JoinDate"] = DateTime.Now;
                staffTable.Rows.Add(staff);
                Console.WriteLine("Default staff account created (staff@hostel.com / staff123)");
            }

            _excelHelper.WriteData("Staff", staffTable);

            // Create default student if doesn't exist
            DataTable studentsTable = _excelHelper.ReadData("Students");
            if (studentsTable.Rows.Count == 0)
            {
                DataRow student = studentsTable.NewRow();
                student["ID"] = "STU-001";
                student["Name"] = "Default Student";
                student["Email"] = "student@hostel.com";
                student["Password"] = "student123";
                student["RoomNumber"] = "0";
                student["FeesPaid"] = 0m;
                student["JoinDate"] = DateTime.Now;
                studentsTable.Rows.Add(student);
                _excelHelper.WriteData("Students", studentsTable);
                Console.WriteLine("Default student account created (student@hostel.com / student123)");
            }
        }

        private static void LoginMenu()
        {
            Console.Write("\nEnter Email: ");
            string email = Console.ReadLine();

            Console.Write("Enter Password: ");
            string password = Console.ReadLine();

            // Check Staff login
            DataTable staffTable = _excelHelper.ReadData("Staff");
            DataRow staff = staffTable.AsEnumerable()
                .FirstOrDefault(row => row["Email"].ToString().Equals(email, StringComparison.OrdinalIgnoreCase) &&
                                     row["Password"].ToString() == password);

            if (staff != null)
            {
                string position = staff["Position"].ToString();
                switch (position)
                {
                    case "Admin":
                        _admin = new Admin(
                            staff["ID"].ToString(),
                            staff["Name"].ToString(),
                            email,
                            password,
                            _excelHelper);
                        AdminMenu();
                        break;
                    case "Manager":
                        _staffManager = new StaffManager(
                            staff["ID"].ToString(),
                            staff["Name"].ToString(),
                            email,
                            password,
                            _excelHelper);
                        StaffManagerMenu();
                        break;
                    case "Staff":
                        _staff = new Staff(
                            staff["ID"].ToString(),
                            staff["Name"].ToString(),
                            email,
                            password,
                            _excelHelper);
                        StaffMenu();
                        break;
                }
                return;
            }

            // Check Student login
            DataTable studentsTable = _excelHelper.ReadData("Students");
            DataRow student = studentsTable.AsEnumerable()
                .FirstOrDefault(row => row["Email"].ToString().Equals(email, StringComparison.OrdinalIgnoreCase) &&
                                     row["Password"].ToString() == password);

            if (student != null)
            {
                _student = new Student(
                    student["ID"].ToString(),
                    student["Name"].ToString(),
                    email,
                    password,
                    _excelHelper);
                StudentMenu();
                return;
            }

            Console.WriteLine("Invalid email or password.");
        }

        private static void AdminMenu()
        {
            Console.WriteLine($"\nWelcome Admin: {_admin.Name}");

            while (true)
            {
                Console.WriteLine("\nAdmin Menu");
                Console.WriteLine("1. Manage Students");
                Console.WriteLine("2. Manage Rooms");
                Console.WriteLine("3. Manage Complaints");
                Console.WriteLine("4. Manage Staff");
                Console.WriteLine("5. View My Details");
                Console.WriteLine("6. Logout");
                Console.Write("Enter choice: ");

                string choice = Console.ReadLine();

                switch (choice)
                {
                    case "1":
                        ManageStudents();
                        break;
                    case "2":
                        ManageRooms();
                        break;
                    case "3":
                        ManageComplaints();
                        break;
                    case "4":
                        ManageStaff();
                        break;
                    case "5":
                        _admin.ViewDetails();
                        Console.WriteLine("Press any key to continue...");
                        Console.ReadKey();
                        break;
                    case "6":
                        return;
                    default:
                        Console.WriteLine("Invalid choice. Please try again.");
                        break;
                }
            }
        }

        private static void ManageStudents()
        {
            while (true)
            {
                Console.WriteLine("\nManage Students");
                Console.WriteLine("1. Add New Student");
                Console.WriteLine("2. View All Students");
                Console.WriteLine("3. Update Students");
                Console.WriteLine("4. Assign Room to Student");
                Console.WriteLine("5. Remove Student");
                Console.WriteLine("6. Back to Admin Menu");
                Console.Write("Enter choice: ");

                string choice = Console.ReadLine();

                switch (choice)
                {
                    case "1":
                        AddStudent();
                        break;
                    case "2":
                        _admin.ViewAllStudents();
                        Console.WriteLine("Press any key to continue...");
                        Console.ReadKey();
                        break;
                        case "3":
                        _admin.ViewAllStudents();
                        _admin.UpdateStudent();
                        break;
                    case "4":
                        AssignRoomToStudent();
                        break;
                    case "5":
                        RemoveStudent();
                        break;
                    case "6":
                        return;
                    default:
                        Console.WriteLine("Invalid choice. Please try again.");
                        break;
                }
            }
        }

        private static void AddStudent()
        {
            Console.Write("Enter Student ID (STU-XXX format): ");
            string id = Console.ReadLine();

            if (!id.StartsWith("STU-") || id.Length < 5 || !int.TryParse(id.Substring(4), out _))
            {
                Console.WriteLine("Invalid ID format. Must be STU- followed by numbers (e.g., STU-001)");
                return;
            }

            DataTable studentsTable = _excelHelper.ReadData("Students");
            if (studentsTable.AsEnumerable().Any(row => row["ID"].ToString() == id))
            {
                Console.WriteLine("Student ID already exists.");
                return;
            }

            Console.Write("Enter Student Name: ");
            string name = Console.ReadLine();

            Console.Write("Enter Student Email: ");
            string email = Console.ReadLine();

            Console.Write("Enter Student Password: ");
            string password = Console.ReadLine();

            DataRow newStudent = studentsTable.NewRow();
            newStudent["ID"] = id;
            newStudent["Name"] = name;
            newStudent["Email"] = email;
            newStudent["Password"] = password;
            newStudent["RoomNumber"] = "0";
            newStudent["FeesPaid"] = 0m;
            newStudent["JoinDate"] = DateTime.Now;

            studentsTable.Rows.Add(newStudent);
            _excelHelper.WriteData("Students", studentsTable);

            Console.WriteLine("Student added successfully!");
        }

        private static void AssignRoomToStudent()
        {
            _admin.ViewAllStudents();
            Console.Write("\nEnter Student ID: ");
            string studentId = Console.ReadLine();

            _admin.ViewRoomStatus();
            Console.Write("Enter Room Number to Assign: ");

            if (int.TryParse(Console.ReadLine(), out int roomNumber))
            {
                _admin.AssignRoom(studentId, roomNumber);
            }
            else
            {
                Console.WriteLine("Invalid room number.");
            }
        }

        private static void RemoveStudent()
        {
            _admin.ViewAllStudents();
            Console.Write("\nEnter Student ID to Remove: ");
            string idToRemove = Console.ReadLine();

            _admin.RemoveStudent(idToRemove);
        }

        private static void ManageRooms()
        {
            while (true)
            {
                Console.WriteLine("\nManage Rooms");
                Console.WriteLine("1. View Room Status");
                Console.WriteLine("2. Back to Admin Menu");
                Console.Write("Enter choice: ");

                string choice = Console.ReadLine();

                switch (choice)
                {
                    case "1":
                        _admin.ViewRoomStatus();
                        Console.WriteLine("Press any key to continue...");
                        Console.ReadKey();
                        break;
                    case "2":
                        return;
                    default:
                        Console.WriteLine("Invalid choice. Please try again.");
                        break;
                }
            }
        }

        private static void ManageComplaints()
        {
            while (true)
            {
                Console.WriteLine("\nManage Complaints");
                Console.WriteLine("1. View All Complaints");
                Console.WriteLine("2. Update Complaint Status");
                Console.WriteLine("3. Back to Admin Menu");
                Console.Write("Enter choice: ");

                string choice = Console.ReadLine();

                switch (choice)
                {
                    case "1":
                        _admin.ViewAllComplaints();
                        Console.WriteLine("Press any key to continue...");
                        Console.ReadKey();
                        break;
                    case "2":
                        UpdateComplaintStatus();
                        break;
                    case "3":
                        return;
                    default:
                        Console.WriteLine("Invalid choice. Please try again.");
                        break;
                }
            }
        }

        private static void UpdateComplaintStatus()
        {
            _admin.ViewAllComplaints();
            Console.Write("\nEnter Complaint ID: ");
            string complaintId = Console.ReadLine();

            Console.Write("Enter New Status (Pending/Resolved/Rejected): ");
            string newStatus = Console.ReadLine();

            _admin.UpdateComplaintStatus(complaintId, newStatus);
        }

        private static void ManageStaff()
        {
            while (true)
            {
                Console.WriteLine("\nManage Staff");
                Console.WriteLine("1. View All Staff");
                Console.WriteLine("2. Add Staff Member");
                Console.WriteLine("3. Remove Staff Member");
                Console.WriteLine("4. Back to Admin Menu");
                Console.Write("Enter choice: ");

                string choice = Console.ReadLine();

                switch (choice)
                {
                    case "1":
                        _admin.ViewAllStaff();
                        Console.WriteLine("Press any key to continue...");
                        Console.ReadKey();
                        break;
                    case "2":
                        AddStaffMember();
                        break;
                    case "3":
                        RemoveStaffMember();
                        break;
                    case "4":
                        return;
                    default:
                        Console.WriteLine("Invalid choice. Please try again.");
                        break;
                }
            }
        }

        private static void AddStaffMember()
        {
            Console.Write("Enter Staff ID (ADM-XXX, MGR-XXX, STF-XXX): ");
            string id = Console.ReadLine();

            if (!id.StartsWith("ADM-") && !id.StartsWith("MGR-") && !id.StartsWith("STF-"))
            {
                Console.WriteLine("Invalid ID format. Must start with ADM-, MGR-, or STF-");
                return;
            }

            DataTable staffTable = _excelHelper.ReadData("Staff");
            if (staffTable.AsEnumerable().Any(row => row["ID"].ToString() == id))
            {
                Console.WriteLine("Staff ID already exists.");
                return;
            }

            Console.Write("Enter Staff Name: ");
            string name = Console.ReadLine();

            Console.Write("Enter Staff Email: ");
            string email = Console.ReadLine();

            Console.Write("Enter Staff Password: ");
            string password = Console.ReadLine();

            string position = id.StartsWith("ADM-") ? "Admin" :
                            id.StartsWith("MGR-") ? "Manager" : "Staff";

            DataRow newStaff = staffTable.NewRow();
            newStaff["ID"] = id;
            newStaff["Name"] = name;
            newStaff["Email"] = email;
            newStaff["Password"] = password;
            newStaff["Position"] = position;
            newStaff["JoinDate"] = DateTime.Now;

            staffTable.Rows.Add(newStaff);
            _excelHelper.WriteData("Staff", staffTable);

            Console.WriteLine("Staff member added successfully!");
        }

        private static void RemoveStaffMember()
        {
            _admin.ViewAllStaff();
            Console.Write("\nEnter Staff ID to Remove: ");
            string staffId = Console.ReadLine();

            _admin.RemoveStaff(staffId);
        }

        private static void StaffManagerMenu()
        {
            Console.WriteLine($"\nWelcome Staff Manager: {_staffManager.Name}");

            while (true)
            {
                Console.WriteLine("\nStaff Manager Menu");
                Console.WriteLine("1. Manage Complaints");
                Console.WriteLine("2. Manage Staff");
                Console.WriteLine("3. View My Details");
                Console.WriteLine("4. Submit Complaint");
                Console.WriteLine("5. Logout");
                Console.Write("Enter choice: ");

                string choice = Console.ReadLine();

                switch (choice)
                {
                    case "1":
                        ManageComplaintsStaffManager();
                        break;
                    case "2":
                        ManageStaffStaffManager();
                        break;
                    case "3":
                        _staffManager.ViewDetails();
                        Console.WriteLine("Press any key to continue...");
                        Console.ReadKey();
                        break;
                    case "4":
                        SubmitComplaint();
                        break;
                    case "5":
                        return;
                    default:
                        Console.WriteLine("Invalid choice. Please try again.");
                        break;
                }
            }
        }

        private static void ManageComplaintsStaffManager()
        {
            while (true)
            {
                Console.WriteLine("\nManage Complaints");
                Console.WriteLine("1. View All Complaints");
                Console.WriteLine("2. Update Complaint Status");
                Console.WriteLine("3. Back to Staff Manager Menu");
                Console.Write("Enter choice: ");

                string choice = Console.ReadLine();

                switch (choice)
                {
                    case "1":
                        _staffManager.ViewAllComplaints();
                        Console.WriteLine("Press any key to continue...");
                        Console.ReadKey();
                        break;
                    case "2":
                        UpdateComplaintStatusStaffManager();
                        break;
                    case "3":
                        return;
                    default:
                        Console.WriteLine("Invalid choice. Please try again.");
                        break;
                }
            }
        }

        private static void UpdateComplaintStatusStaffManager()
        {
            _staffManager.ViewAllComplaints();
            Console.Write("\nEnter Complaint ID: ");
            string complaintId = Console.ReadLine();

            Console.Write("Enter New Status (Pending/Resolved/Rejected): ");
            string newStatus = Console.ReadLine();

            _staffManager.UpdateComplaintStatus(complaintId, newStatus);
        }

        private static void ManageStaffStaffManager()
        {
            while (true)
            {
                Console.WriteLine("\nManage Staff");
                Console.WriteLine("1. View All Staff");
                Console.WriteLine("2. Add Staff");
                Console.WriteLine("3. Remove Staff");
                Console.WriteLine("4. Back to Staff Manager Menu");
                Console.Write("Enter choice: ");

                string choice = Console.ReadLine();

                switch (choice)
                {
                    case "1":
                        _staffManager.ViewAllStaff();
                        Console.WriteLine("Press any key to continue...");
                        Console.ReadKey();
                        break;
                    case "2":
                        AddStaffByManager();
                        break;
                    case "3":
                        RemoveStaffByManager();
                        break;
                    case "4":
                        return;
                    default:
                        Console.WriteLine("Invalid choice. Please try again.");
                        break;
                }
            }
        }

        private static void AddStaffByManager()
        {
            Console.Write("Enter Staff ID (STF-XXX): ");
            string id = Console.ReadLine();

            if (!id.StartsWith("STF-"))
            {
                Console.WriteLine("Manager can only create staff with STF- prefix");
                return;
            }

            DataTable staffTable = _excelHelper.ReadData("Staff");
            if (staffTable.AsEnumerable().Any(row => row["ID"].ToString() == id))
            {
                Console.WriteLine("Staff ID already exists.");
                return;
            }

            Console.Write("Enter Staff Name: ");
            string name = Console.ReadLine();

            Console.Write("Enter Staff Email: ");
            string email = Console.ReadLine();

            Console.Write("Enter Staff Password: ");
            string password = Console.ReadLine();

            DataRow newStaff = staffTable.NewRow();
            newStaff["ID"] = id;
            newStaff["Name"] = name;
            newStaff["Email"] = email;
            newStaff["Password"] = password;
            newStaff["Position"] = "Staff";
            newStaff["JoinDate"] = DateTime.Now;

            staffTable.Rows.Add(newStaff);
            _excelHelper.WriteData("Staff", staffTable);

            Console.WriteLine("Staff member added successfully!");
        }

        private static void RemoveStaffByManager()
        {
            _staffManager.ViewAllStaff();
            Console.Write("\nEnter Staff ID to Remove: ");
            string staffId = Console.ReadLine();

            _staffManager.RemoveStaff(staffId);
        }

        private static void SubmitComplaint()
        {
            Console.Write("Enter Complaint Description: ");
            string description = Console.ReadLine();

            _staffManager.SubmitComplaint(description);
        }

        private static void StaffMenu()
        {
            Console.WriteLine($"\nWelcome Staff: {_staff.Name}");

            while (true)
            {
                Console.WriteLine("\nStaff Menu");
                Console.WriteLine("1. View My Details");
                Console.WriteLine("2. Submit Complaint");
                Console.WriteLine("3. Logout");
                Console.Write("Enter choice: ");

                string choice = Console.ReadLine();

                switch (choice)
                {
                    case "1":
                        _staff.ViewDetails();
                        Console.WriteLine("Press any key to continue...");
                        Console.ReadKey();
                        break;
                    case "2":
                        SubmitStaffComplaint();
                        break;
                    case "3":
                        return;
                    default:
                        Console.WriteLine("Invalid choice. Please try again.");
                        break;
                }
            }
        }

        private static void SubmitStaffComplaint()
        {
            Console.Write("Enter Complaint Description: ");
            string description = Console.ReadLine();

            _staff.SubmitComplaint(description);
        }

        private static void StudentMenu()
        {
            Console.WriteLine($"\nWelcome Student: {_student.Name}");

            while (true)
            {
                Console.WriteLine("\nStudent Menu");
                Console.WriteLine("1. View My Details");
                Console.WriteLine("2. View My Complaints");
                Console.WriteLine("3. Submit Complaint");
                Console.WriteLine("4. Pay Fees");
                Console.WriteLine("5. Logout");
                Console.Write("Enter choice: ");

                string choice = Console.ReadLine();

                switch (choice)
                {
                    case "1":
                        _student.ViewDetails();
                        Console.WriteLine("Press any key to continue...");
                        Console.ReadKey();
                        break;
                    case "2":
                        _student.ViewComplaints();
                        Console.WriteLine("Press any key to continue...");
                        Console.ReadKey();
                        break;
                    case "3":
                        SubmitStudentComplaint();
                        break;
                    case "4":
                        PayStudentFees();
                        break;
                    case "5":
                        return;
                    default:
                        Console.WriteLine("Invalid choice. Please try again.");
                        break;
                }
            }
        }

        private static void SubmitStudentComplaint()
        {
            Console.Write("Enter Complaint Description: ");
            string description = Console.ReadLine();

            _student.SubmitComplaint(description);
        }

        private static void PayStudentFees()
        {
            Console.Write("Enter Amount to Pay: ");
            if (decimal.TryParse(Console.ReadLine(), out decimal amount))
            {
                _student.PayFees(amount);
            }
            else
            {
                Console.WriteLine("Invalid amount.");
            }
        }
    }
}