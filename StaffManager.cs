using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HostelManagementSystem
{
    public class StaffManager : Staff, IStaffOperations
    {
        public StaffManager(string id, string name, string email, string password, ExcelHelper excelHelper)
            : base(id, name, email, password, excelHelper)
        {
            Position = "Manager";
        }

        public void ViewAllComplaints()
        {
            DataTable complaintsTable = ExcelHelper.ReadData("Complaints");

            Console.WriteLine("\nAll Complaints:");
            foreach (DataRow complaint in complaintsTable.Rows)
            {
                Console.WriteLine($"ID: {complaint["ID"]}");
                Console.WriteLine($"Student/Staff ID: {complaint["StudentID"]}");
                Console.WriteLine($"Date: {complaint["Date"]}");
                Console.WriteLine($"Description: {complaint["Description"]}");
                Console.WriteLine($"Status: {complaint["Status"]}\n");
            }
        }

        public void UpdateComplaintStatus(string complaintId, string newStatus)
        {
            ExcelHelper.UpdateRow("Complaints", "ID", complaintId,
                new Dictionary<string, string> { { "Status", newStatus } });
            Console.WriteLine("Complaint status updated successfully!");
        }

        public void ViewAllStaff()
        {
            DataTable staffTable = ExcelHelper.ReadData("Staff");

            Console.WriteLine("\nAll Staff Members:");
            foreach (DataRow staff in staffTable.Rows)
            {
                Console.WriteLine($"ID: {staff["ID"]}");
                Console.WriteLine($"Name: {staff["Name"]}");
                Console.WriteLine($"Email: {staff["Email"]}");
                Console.WriteLine($"Position: {staff["Position"]}\n");
            }
        }

        public void AddStaff(string name, string email, string password, string position)
        {
            DataTable staffTable = ExcelHelper.ReadData("Staff");
            DataRow newStaff = staffTable.NewRow();
            newStaff["ID"] = "STF-" + Guid.NewGuid().ToString().Substring(0, 8);
            newStaff["Name"] = name;
            newStaff["Email"] = email;
            newStaff["Password"] = password;
            newStaff["Position"] = position;

            ExcelHelper.AppendData("Staff", newStaff);
            Console.WriteLine($"Staff member {name} added successfully with ID: {newStaff["ID"]}");
        }

        public void RemoveStaff(string staffId)
        {
            ExcelHelper.DeleteRow("Staff", "ID", staffId);
            Console.WriteLine("Staff member removed successfully!");
        }
    }
}
