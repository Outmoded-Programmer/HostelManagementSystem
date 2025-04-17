using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HostelManagementSystem
{
    public abstract class ComplaintHandler : Person, IComplaintHandler
    {
        protected ComplaintHandler(string id, string name, string email, string password, ExcelHelper excelHelper)
            : base(id, name, email, password, excelHelper)
        {
        }

        public void SubmitComplaint(string description)
        {
            DataTable complaintsTable = ExcelHelper.ReadData("Complaints");
            DataRow newComplaint = complaintsTable.NewRow();
            newComplaint["ID"] = Guid.NewGuid().ToString();
            newComplaint["StudentID"] = this.ID;
            newComplaint["Description"] = description;
            newComplaint["Status"] = "Pending";
            newComplaint["Date"] = DateTime.Now.ToString("yyyy-MM-dd");

            ExcelHelper.AppendData("Complaints", newComplaint);
            Console.WriteLine("Complaint submitted successfully!");
        }

        public void ViewComplaints()
        {
            DataTable complaintsTable = ExcelHelper.ReadData("Complaints");
            var userComplaints = complaintsTable.Select($"StudentID = '{ID}'");

            Console.WriteLine("\nYour Complaints:");
            foreach (var complaint in userComplaints)
            {
                Console.WriteLine($"Date: {complaint["Date"]}");
                Console.WriteLine($"Description: {complaint["Description"]}");
                Console.WriteLine($"Status: {complaint["Status"]}\n");
            }
        }
    }
}
