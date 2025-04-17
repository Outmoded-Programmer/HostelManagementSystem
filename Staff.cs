using System;
using System.Data;

namespace HostelManagementSystem
{
    public class Staff : Person, IComplaintHandler
    {
        public string Position { get; protected set; }

        public Staff(string id, string name, string email, string password, ExcelHelper excelHelper)
            : base(id, name, email, password, excelHelper)
        {
            DataRow staffData = FindPersonInSheet("Staff", id);
            if (staffData != null)
            {
                Position = staffData["Position"].ToString();
            }
        }

        public void SubmitComplaint(string description)
        {
            DataTable complaintsTable = ExcelHelper.ReadData("Complaints");
            DataRow newComplaint = complaintsTable.NewRow();
            newComplaint["ID"] = Guid.NewGuid().ToString();
            newComplaint["StudentID"] = "STAFF-" + this.ID;
            newComplaint["Description"] = description;
            newComplaint["Status"] = "Pending";
            newComplaint["Date"] = DateTime.Now.ToString("yyyy-MM-dd");

            ExcelHelper.AppendData("Complaints", newComplaint);
            Console.WriteLine("Complaint submitted successfully!");
        }

        public void ViewComplaints()
        {
            DataTable complaintsTable = ExcelHelper.ReadData("Complaints");
            var userComplaints = complaintsTable.Select($"StudentID = 'STAFF-{this.ID}'");

            Console.WriteLine("\nYour Complaints:");
            foreach (var complaint in userComplaints)
            {
                Console.WriteLine($"Date: {complaint["Date"]}");
                Console.WriteLine($"Description: {complaint["Description"]}");
                Console.WriteLine($"Status: {complaint["Status"]}\n");
            }
        }

        public override void ViewDetails()
        {
            Console.WriteLine($"ID: {ID}");
            Console.WriteLine($"Name: {Name}");
            Console.WriteLine($"Email: {Email}");
            Console.WriteLine($"Position: {Position}");
        }
    }
}